#!/usr/bin/env python3
"""
app_cli.py
Menu-driven version of the same tool: no browser needed. The client types
a number to pick the form, then a path to their raw Excel file, picks a
mode, and gets the result.

Two modes, saved into two separate folders next to the raw file so you
always have two clean deliverables to hand out at different times:
  1. Report  -> saved in a "1_Report" folder. A pure FINDINGS REPORT --
     it never touches, copies, or modifies the client's sheet in any way.
     Two tables: "Column Changes" (which raw columns get renamed or are
     missing) and "Row Issues" (which specific row/column/value needs
     fixing and why). Give this to the client FIRST so they fix their own
     source file.
  2. Final   -> saved in a "2_Final" folder. The actual converted sheet,
     ready to feed into Zoho Creator's bulk import. Generate this once
     their data is clean.

RUN IT (see run_cli.bat for a double-click shortcut on Windows):
    pip install -r requirements.txt
    python app_cli.py
"""

import os
import sys
from pathlib import Path

from cheds_import_tool import build_import_ready, build_issue_report
from cheds_forms import FORMS, UNVERIFIED_NO_WORKFLOW

APP_DIR = Path(__file__).resolve().parent
HEDB_PATH = APP_DIR / "HEDB_.xlsx"


def ask_form():
    names = list(FORMS.keys())
    print("\nWhich form are you working with?\n")
    for name in names:
        print(f"  {name}")
    while True:
        choice = input(f"\nEnter a number (1-{len(names)}): ").strip()
        if choice.isdigit() and 1 <= int(choice) <= len(names):
            return names[int(choice) - 1]
        print("Not a valid choice -- try again.")


def ask_mode():
    print("\nWhat do you want?\n")
    print("  1. Report   (findings only -- doesn't touch their sheet -- goes in a '1_Report' folder)")
    print("  2. Final    (the actual converted import-ready sheet -- goes in a '2_Final' folder)")
    while True:
        choice = input("\nEnter a number (1-2): ").strip()
        if choice in ("1", "2"):
            return "report" if choice == "1" else "build"
        print("Not a valid choice -- try again.")


def ask_raw_path():
    while True:
        raw_path = input("\nPath to the raw Excel file (.xlsx): ").strip().strip('"')
        if os.path.isfile(raw_path):
            return raw_path
        print(f"Can't find that file: {raw_path!r} -- try again.")


def main():
    print("=" * 60)
    print("CHEDS Import Sheet Builder")
    print("=" * 60)

    if not HEDB_PATH.exists():
        print(f"\nERROR: HEDB_.xlsx not found next to this script (expected at {HEDB_PATH}).")
        print("Copy HEDB_.xlsx into this same folder and try again.")
        sys.exit(1)

    form_name = ask_form()

    if form_name in UNVERIFIED_NO_WORKFLOW:
        print(
            "\nHeads up: this form has no confirmed live push-to-CHEDS workflow yet "
            "(see its mapping notes). The column layout is still correct, but double-"
            "check with whoever owns the Zoho integration before relying on it."
        )

    mode = ask_mode()
    raw_path = ask_raw_path()
    sheet = input("Sheet name in the raw file (press Enter to use the first sheet): ").strip() or None

    safe_name = form_name.split(". ", 1)[-1].replace(" ", "_").replace("/", "_")
    raw_dir = os.path.dirname(os.path.abspath(raw_path)) or "."

    if mode == "report":
        deliverable_dir = os.path.join(raw_dir, "1_Report")
        default_out = os.path.join(deliverable_dir, f"{safe_name}_Report.xlsx")
    else:
        deliverable_dir = os.path.join(raw_dir, "2_Final")
        default_out = os.path.join(deliverable_dir, f"{safe_name}_Import_Ready.xlsx")

    out_path = input(f"Output file path (press Enter for '{default_out}'): ").strip().strip('"') or default_out
    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)

    config_path = str(APP_DIR / FORMS[form_name])

    print("\nBuilding...")
    try:
        if mode == "report":
            stats = build_issue_report(raw_path, config_path, str(HEDB_PATH), out_path, sheet=sheet)
        else:
            stats = build_import_ready(raw_path, config_path, str(HEDB_PATH), out_path, sheet=sheet)
    except Exception as e:
        print(f"\nFAILED: {e}")
        sys.exit(1)

    if mode == "report":
        print(f"\nWrote {out_path}  ({stats['rows']} row(s) checked)")
        print(f"Must fix:      {stats['errors']}")
        print(f"Please check:  {stats['warnings']}")
        print(f"Info only:     {stats['infos']}")
        if stats["missing_columns"]:
            print("\nColumns expected but not found in this file at all:")
            for mc in stats["missing_columns"]:
                print(f"  - expected header {mc['raw_column']!r} (for field {mc['target_field']!r})")
        print("\nGive the client this report. 'Column Changes' lists every column and whether "
              "it's renamed/missing. 'Row Issues' lists the specific rows/columns/values to fix, "
              "with the row number matching their own file. Nothing in their sheet was touched.")
    else:
        print(f"\nWrote {out_path}  ({stats['rows']} row(s))")
        print(f"Lookup fields resolved OK: {stats['lookup_ok']}")
        print(f"Lookup fields UNRESOLVED (check the _NOTES column): {stats['lookup_unresolved']}")
        print(f"Pending/custom-function fields (check the _NOTES column): {stats['pending']}")
        print(f"Date values cleaned (time stripped): {stats['dates_fixed']}")
        if stats["sheet_load_errors"]:
            print("\nWARNING -- these hedb_sheet values in the mapping aren't usable yet:")
            for sheet_name_, err in stats["sheet_load_errors"].items():
                print(f"  - {sheet_name_!r}: {err}")
        print("\nThis is the final sheet -- ready to import into Zoho once its _NOTES column is clear.")


if __name__ == "__main__":
    main()
