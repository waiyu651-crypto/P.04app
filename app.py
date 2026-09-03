"""
GSO/GCTS Tools — web front end
================================
Lets you pick between three tools and run them from a browser instead of
the command line:

  1. GCTS Report generator   (fast, no network, always works hosted)
  2. GSO Model Extractor     (Playwright scraper, public pages, no login)
  3. GSO Product-ID Lookup   (needs a session file generated locally once)

Run locally with:   streamlit run app.py
Free hosting instructions are in README.md.
"""

import os
import io
import json
import subprocess
import sys
import tempfile
from pathlib import Path

import streamlit as st

st.set_page_config(page_title="GSO / GCTS Tools", page_icon="🧰", layout="centered")

# --------------------------------------------------------------------- #
# One-time setup: make sure the Playwright chromium browser is installed.
# On a fresh free-tier host there is no browser binary until this runs.
# --------------------------------------------------------------------- #
@st.cache_resource(show_spinner="Setting up browser engine (first run only)...")
def ensure_playwright_browser():
    try:
        subprocess.run(
            [sys.executable, "-m", "playwright", "install", "--with-deps", "chromium"],
            check=True, capture_output=True, text=True, timeout=600,
        )
        return True
    except Exception as e:
        return f"error: {e}"


st.title("🧰 GSO / GCTS Tools")
st.caption("Pick a tool below. Each one uploads a file, runs the script, and gives you a file back.")

tool = st.sidebar.radio(
    "Choose a tool",
    [
        "1. GCTS Report generator",
        "2. GSO Model Extractor",
        "3. GSO Product-ID Lookup",
    ],
)

WORKDIR = Path(tempfile.gettempdir()) / "gso_webapp_runs"
WORKDIR.mkdir(exist_ok=True)

# ======================================================================
# TOOL 1 — GCTS Report generator
# ======================================================================
if tool.startswith("1"):
    st.header("GCTS Lookup → Word report")
    st.write(
        "Upload the **All 35 GCTS Lookup** Excel workbook. Rows whose GCTS number "
        "is highlighted yellow or green are dropped; the rest are grouped by brand "
        "into a formatted Word report."
    )

    xlsx_file = st.file_uploader("Excel file (.xlsx)", type=["xlsx"], key="t1_xlsx")

    if xlsx_file and st.button("Generate report", type="primary"):
        sys.path.insert(0, str(Path(__file__).parent))
        import gcts_report as gr

        with st.spinner("Reading rows and building the document..."):
            tmp_in = WORKDIR / "t1_input.xlsx"
            tmp_in.write_bytes(xlsx_file.getvalue())

            rows = gr.load_rows(str(tmp_in))
            if not rows:
                st.error("No non-highlighted rows found — check the sheet name / highlight colors.")
            else:
                tmp_out = WORKDIR / "GCTS_Report.docx"
                gr.build_report(rows, str(tmp_out))
                st.success(f"Done — {len(rows)} rows across {len(gr.group_by_brand(rows))} brand table(s).")
                st.download_button(
                    "⬇️ Download GCTS_Report.docx",
                    data=tmp_out.read_bytes(),
                    file_name="GCTS_Report.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )

# ======================================================================
# TOOL 2 — GSO Model Extractor
# ======================================================================
elif tool.startswith("2"):
    st.header("GSO Model Extractor")
    st.write(
        "Upload an Excel file with a **GSO Product Page** URL column (or a "
        "**Product GCTS** number column). This scrapes each public product page "
        "for its Model Numbers table and Product Specification fields."
    )
    st.warning(
        "⚠️ This runs a real headless browser per row, so it's slow (roughly "
        "1–2 seconds/row) and free hosts will kill requests that run too long. "
        "Use the row limit below to test first — for a full multi-thousand-row "
        "run, do that locally instead of through the browser."
    )

    xlsx_file = st.file_uploader("Excel file (.xlsx)", type=["xlsx"], key="t2_xlsx")
    test_rows = st.number_input(
        "Only process this many rows (0 = no limit — not recommended on a free host)",
        min_value=0, value=5, step=1,
    )

    if xlsx_file and st.button("Run extractor", type="primary"):
        setup = ensure_playwright_browser()
        if setup is not True:
            st.error(f"Could not install the browser engine: {setup}")
        else:
            tmp_in = WORKDIR / "GSO_Batch2_Input.xlsx"
            tmp_in.write_bytes(xlsx_file.getvalue())

            env = os.environ.copy()
            env["WEBAPP_INPUT_FILE"] = str(tmp_in)
            if test_rows > 0:
                env["WEBAPP_TEST_ROWS"] = str(test_rows)

            log_box = st.empty()
            logs = []
            with st.spinner("Scraping pages..."):
                proc = subprocess.Popen(
                    [sys.executable, str(Path(__file__).parent / "gso_model_extractor.py")],
                    cwd=WORKDIR, env=env, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True,
                )
                for line in proc.stdout:
                    logs.append(line.rstrip())
                    log_box.code("\n".join(logs[-30:]))
                proc.wait()

            out_path = tmp_in.with_name(tmp_in.stem + "_filled.xlsx")
            if proc.returncode == 0 and out_path.exists():
                st.success("Done.")
                st.download_button(
                    "⬇️ Download results",
                    data=out_path.read_bytes(),
                    file_name=out_path.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                st.error(f"Script exited with code {proc.returncode}. See log above.")

# ======================================================================
# TOOL 3 — GSO Product-ID Lookup
# ======================================================================
else:
    st.header("GSO Product-ID → GCTS/NB Lookup")
    st.write(
        "This tool needs a **logged-in session** to call the GSO lookup API. "
        "A website can't do the interactive login step for you, so you generate "
        "the session once on your own computer, then upload it here."
    )

    with st.expander("How to get the session file (one-time, on your own computer)"):
        st.markdown(
            "1. Run `python gso_product_id_lookup.py` locally.\n"
            "2. A real browser window opens — log in to GSO normally.\n"
            "3. Press Enter in the terminal once logged in.\n"
            "4. This creates a `gso_state.json` file next to the script.\n"
            "5. Upload that file below. Sessions expire after a while — "
            "regenerate it if lookups start failing."
        )

    input_file = st.file_uploader("Input Excel file (.xlsx)", type=["xlsx"], key="t3_xlsx")
    state_file = st.file_uploader("Session file (gso_state.json)", type=["json"], key="t3_state")
    sheet_name = st.text_input("Sheet name", value="Model Lookup")
    id_col = st.text_input("Product ID column name", value="Product GCTS")
    test_limit = st.number_input(
        "Only process this many new rows (0 = no limit — not recommended on a free host)",
        min_value=0, value=10, step=1,
    )

    if input_file and state_file and st.button("Run lookup", type="primary"):
        try:
            json.loads(state_file.getvalue())
        except Exception:
            st.error("That doesn't look like a valid session JSON file.")
            st.stop()

        tmp_in = WORKDIR / "t3_input.xlsx"
        tmp_in.write_bytes(input_file.getvalue())
        tmp_state = WORKDIR / "gso_state.json"
        tmp_state.write_bytes(state_file.getvalue())
        tmp_out = WORKDIR / "t3_output.xlsx"
        tmp_progress = WORKDIR / "t3_progress.json"

        # patch the sheet/column names the script reads at import time
        script_src = (Path(__file__).parent / "gso_product_id_lookup.py").read_text()
        script_src = script_src.replace(
            'INPUT_SHEET = "Model Lookup"', f'INPUT_SHEET = {sheet_name!r}'
        ).replace(
            'PRODUCT_ID_COLUMN = "Product GCTS"', f'PRODUCT_ID_COLUMN = {id_col!r}'
        )
        run_script = WORKDIR / "_t3_run.py"
        run_script.write_text(script_src)

        env = os.environ.copy()
        env["WEBAPP_INPUT_FILE"] = str(tmp_in)
        env["WEBAPP_OUTPUT_FILE"] = str(tmp_out)
        env["WEBAPP_STATE_FILE"] = str(tmp_state)
        if test_limit > 0:
            env["WEBAPP_TEST_LIMIT"] = str(test_limit)

        # progress file lives next to the script (PROGRESS_FILE is not overridable)
        prog_target = WORKDIR / "gcts_lookup_progress.json"
        if prog_target.exists():
            prog_target.unlink()

        log_box = st.empty()
        logs = []
        with st.spinner("Looking up product IDs..."):
            proc = subprocess.Popen(
                [sys.executable, str(run_script)],
                cwd=WORKDIR, env=env, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True,
            )
            for line in proc.stdout:
                logs.append(line.rstrip())
                log_box.code("\n".join(logs[-30:]))
            proc.wait()

        if tmp_out.exists():
            st.success("Done — see log above for Found/Not Found counts.")
            st.download_button(
                "⬇️ Download results",
                data=tmp_out.read_bytes(),
                file_name="GSO_merged_with_gcts.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.error(f"Script exited with code {proc.returncode} and produced no output. See log above.")
