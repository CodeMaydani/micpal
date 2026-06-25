"""
gen_company.py
CLI wrapper around engine.py -- same behavior as the original script,
now built on the shared engine module that app.py (Streamlit UI) also uses.

Usage:
    python3 gen_company.py [company] [data_dir] [out_dir] [year] [month]

All arguments are optional; defaults match the original script.
This only generates the template .txt and the .xlsx -- it does NOT insert
into Q8SRGL26.000. Use the insertion command from the project docs, or the
Streamlit UI (app.py), for that step.
"""

import datetime
import os
import sys

import config
import engine


def main(company, data_dir, out_dir, year, month):
    os.makedirs(out_dir, exist_ok=True)
    template_name = f"תבנית משכורות אוטומציה {company}"
    mifl_path = f"{data_dir}/Q8MIFL26.{company}"
    ovdm_path = f"{data_dir}/Q8OVDM26.{company}"

    print(f"Processing company {company}...")

    print("  Extracting components...")
    components, skipped = engine.extract_components(mifl_path)
    print(f"  Found {len(components)} components")
    if skipped:
        print(f"    ({skipped} candidate component records filtered out)")

    print("  Extracting employees...")
    employees, invalid = engine.extract_employees(ovdm_path)
    print(f"  Found {len(employees)} employees")
    if invalid:
        print(f"    ({invalid} employee blocks had an invalid תעודת זהות)")

    print("  Building template...")
    template_text, col_map, stats_cols = engine.build_template(
        template_name, components
    )

    template_path = f"{out_dir}/template_{company}.txt"
    with open(template_path, "w", encoding="utf-8") as f:
        f.write(template_text)
    print(f"  Template saved: {template_path}")

    excel_path = f"{out_dir}/template_{company}.xlsx"
    engine.build_excel(
        company, components, col_map, stats_cols, employees, excel_path, year, month
    )
    print(f"  Excel saved:    {excel_path}")

    print("\nDone.")
    print(f"  Components: {len(components)} | Employees: {len(employees)}")

    return template_path, excel_path


if __name__ == "__main__":
    today = datetime.date.today()

    cfg = config.load()

    args = sys.argv[1:]
    COMPANY = args[0] if len(args) > 0 else "083"
    DATA_DIR = args[1] if len(args) > 1 else cfg["data_dir"]
    OUT_DIR = args[2] if len(args) > 2 else cfg["out_dir"]
    YEAR = int(args[3]) if len(args) > 3 else today.year
    MONTH = int(args[4]) if len(args) > 4 else today.month

    main(COMPANY, DATA_DIR, OUT_DIR, YEAR, MONTH)
