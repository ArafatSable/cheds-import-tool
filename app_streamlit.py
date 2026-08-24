"""
app_streamlit.py
Local web app: client picks the form from a dropdown, uploads their raw
Excel export, picks a mode, and downloads the result. Runs entirely on the
client's own machine -- nothing leaves it.

Two modes:
  - Report: a pure FINDINGS REPORT -- never touches, copies, or modifies the
    raw sheet. Two tables: "Column Changes" (renames/missing columns) and
    "Row Issues" (specific row/column/value problems, with the row number
    matching their own file). Hand this out FIRST so they fix their own data.
  - Final: the actual converted sheet, ready for Zoho Creator's bulk import.
    Generate this once their data is clean.

RUN IT (see run_web_app.bat for a double-click shortcut on Windows):
    pip install -r requirements.txt
    streamlit run app_streamlit.py
"""

import os
import tempfile
from pathlib import Path

import streamlit as st

from cheds_import_tool import build_import_ready, build_issue_report
from cheds_forms import FORMS, UNVERIFIED_NO_WORKFLOW

APP_DIR = Path(__file__).resolve().parent
HEDB_PATH = APP_DIR / "HEDB_.xlsx"

st.set_page_config(page_title="CHEDS Import Sheet Builder", page_icon="\U0001F4C4", layout="centered")
st.title("CHEDS Import Sheet Builder")
st.write("Pick the form and upload the raw Excel export for it.")

if not HEDB_PATH.exists():
    st.error(
        f"HEDB_.xlsx not found next to this app (expected at {HEDB_PATH}). "
        "Copy HEDB_.xlsx into this same folder and restart the app."
    )
    st.stop()

form_names = list(FORMS.keys())
form_name = st.selectbox("Which form is this?", form_names)

if form_name in UNVERIFIED_NO_WORKFLOW:
    st.warning(
        "Heads up: this form has no confirmed live push-to-CHEDS workflow yet "
        "(see this form's mapping notes). The column layout is still correct, "
        "but double-check with whoever owns the Zoho integration before relying "
        "on it for a real submission."
    )

mode = st.radio(
    "What do you want?",
    [
        "1. Report -- findings only, doesn't touch their sheet (give this to the client first)",
        "2. Final -- the actual converted import-ready sheet (once their data is clean)",
    ],
)
is_report = mode.startswith("1.")

uploaded = st.file_uploader("Raw Excel file for this form (.xlsx)", type=["xlsx"])
sheet_name = st.text_input(
    "Sheet name in the raw file (leave blank to use the first sheet)", value=""
)

button_label = "Generate report" if is_report else "Generate final import-ready sheet"
generate = st.button(button_label, type="primary", disabled=uploaded is None)

if generate and uploaded is not None:
    with tempfile.TemporaryDirectory() as tmp:
        raw_path = os.path.join(tmp, "raw_upload.xlsx")
        with open(raw_path, "wb") as f:
            f.write(uploaded.getbuffer())

        config_path = str(APP_DIR / FORMS[form_name])
        safe_name = form_name.split(". ", 1)[-1].replace(" ", "_").replace("/", "_")
        prefix = "1_Report" if is_report else "2_Final"
        suffix = "Report" if is_report else "Import_Ready"
        out_path = os.path.join(tmp, f"{prefix}_{safe_name}_{suffix}.xlsx")

        try:
            if is_report:
                stats = build_issue_report(
                    raw_path, config_path, str(HEDB_PATH), out_path,
                    sheet=sheet_name.strip() or None,
                )
            else:
                stats = build_import_ready(
                    raw_path, config_path, str(HEDB_PATH), out_path,
                    sheet=sheet_name.strip() or None,
                )
        except Exception as e:
            st.error(f"Something went wrong: {e}")
        else:
            if is_report:
                st.success(f"Done -- {stats['rows']} row(s) checked.")
                c1, c2, c3 = st.columns(3)
                c1.metric("Must fix", stats["errors"])
                c2.metric("Please check", stats["warnings"])
                c3.metric("Info only", stats["infos"])
                if stats["missing_columns"]:
                    st.error(
                        "These expected columns weren't found in the file at all:\n\n"
                        + "\n".join(
                            f"- expected header **{mc['raw_column']}** (for field {mc['target_field']})"
                            for mc in stats["missing_columns"]
                        )
                    )
                st.info(
                    "This file never touches the client's data -- it's a pure findings report. "
                    "**Column Changes** lists every column and whether it's renamed/missing. "
                    "**Row Issues** lists the specific rows/columns/values to fix, with row "
                    "numbers matching their own file."
                )
            else:
                if stats["sheet_load_errors"]:
                    st.warning(
                        "Some HEDB sheet names in this form's mapping aren't resolvable yet -- "
                        "those columns were left as raw values instead of being converted:\n\n"
                        + "\n".join(f"- **{k}**: {v}" for k, v in stats["sheet_load_errors"].items())
                    )
                st.success(f"Done -- {stats['rows']} row(s) processed.")
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Rows", stats["rows"])
                c2.metric("Lookups OK", stats["lookup_ok"])
                c3.metric("Needs review", stats["lookup_unresolved"] + stats["pending"])
                c4.metric("Dates fixed", stats["dates_fixed"])
                if stats["lookup_unresolved"] or stats["pending"]:
                    st.info(
                        "Some rows have a note in the output file's trailing **_NOTES** column -- "
                        "open it in Excel and check those before importing into Zoho."
                    )

            with open(out_path, "rb") as f:
                st.download_button(
                    f"Download {'report' if is_report else 'final import-ready Excel'}",
                    data=f.read(),
                    file_name=os.path.basename(out_path),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
