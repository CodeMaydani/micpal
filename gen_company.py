"""
gen_company.py
CLI wrapper around engine.py -- same behavior as the original script,
now built on the shared engine module that app.py (Streamlit UI) also uses.

Usage:
    python3 gen_company.py [company] [data_dir] [out_dir] [year] [month] [options]

Positional arguments are optional; defaults match the original script.
Options:
    --carry-forward          Pre-fill each employee row with last month's
                             qty/price values, read from the CURRENT
                             Q8OVDM26.[company] (which holds the previous
                             month until the new sheet is imported). New
                             hires (absent last month) get a blank row with
                             their תז/שם highlighted yellow in the Excel.
    --no-carry N[,N...]      EXTRA component numbers to exclude from carry-
                             forward, on top of the auto-detected leave
                             default (vacation/sick/holiday/maternity).
    --carry-all              Carry forward everything, ignoring even the
                             leave default (overrides --no-carry).

This only generates the template .txt and the .xlsx -- it does NOT insert
into Q8SRGL26.000. Use the insertion command from the project docs, or the
Streamlit UI (app.py), for that step.
"""

import datetime
import os
import sys

import config
import engine


def main(
    company,
    data_dir,
    out_dir,
    year,
    month,
    carry_forward=False,
    no_carry_extra=None,
    carry_all=False,
):
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

    # --- carry-forward: read previous month's values from the same Q8OVDM26 ---
    prior_month = None
    no_carry = set()
    if carry_forward:
        print("  Reading previous month's values for carry-forward...")
        prior_month = engine.read_prior_month(ovdm_path)
        if carry_all:
            no_carry = set()
        else:
            no_carry = engine.default_no_carry_components(components)
            if no_carry_extra:
                no_carry |= set(no_carry_extra)
        print(
            f"    {len(prior_month)} employees had prior-month data; "
            f"{len(no_carry)} component(s) excluded from carry-forward"
        )

    excel_path = f"{out_dir}/template_{company}.xlsx"
    result = engine.build_excel(
        company,
        components,
        col_map,
        stats_cols,
        employees,
        excel_path,
        year,
        month,
        prior_month=prior_month,
        no_carry=no_carry,
    )
    print(f"  Excel saved:    {result['out_path']}")
    if carry_forward:
        print(f"    Cells carried forward: {result['carried']}")
        if result["new_hires"]:
            print(
                f"    New hires (blank row, highlighted yellow): "
                f"{len(result['new_hires'])}"
            )
            for tz, name in result["new_hires"]:
                print(f"      - {tz or '(no תז)'}  {name}")
        if result["no_activity"]:
            print(
                f"    Employees present last month with no activity "
                f"(blank row): {len(result['no_activity'])}"
            )

    print("\nDone.")
    print(f"  Components: {len(components)} | Employees: {len(employees)}")

    return template_path, result["out_path"]


def _parse_args(argv):
    """
    Split argv into positionals and options. Positionals keep their original
    order/meaning (company, data_dir, out_dir, year, month); options may
    appear anywhere after them.
    """
    positionals = []
    carry_forward = False
    carry_all = False
    no_carry_extra = []

    i = 0
    while i < len(argv):
        a = argv[i]
        if a == "--carry-forward":
            carry_forward = True
        elif a == "--carry-all":
            carry_all = True
            carry_forward = True
        elif a == "--no-carry":
            i += 1
            if i < len(argv):
                no_carry_extra = [
                    int(x) for x in argv[i].split(",") if x.strip().isdigit()
                ]
        elif a.startswith("--no-carry="):
            no_carry_extra = [
                int(x) for x in a.split("=", 1)[1].split(",") if x.strip().isdigit()
            ]
        else:
            positionals.append(a)
        i += 1

    return positionals, carry_forward, carry_all, no_carry_extra


if __name__ == "__main__":
    today = datetime.date.today()
    cfg = config.load()

    positionals, carry_forward, carry_all, no_carry_extra = _parse_args(sys.argv[1:])

    COMPANY = positionals[0] if len(positionals) > 0 else "083"
    DATA_DIR = positionals[1] if len(positionals) > 1 else cfg["data_dir"]
    OUT_DIR = positionals[2] if len(positionals) > 2 else cfg["out_dir"]
    YEAR = int(positionals[3]) if len(positionals) > 3 else today.year
    MONTH = int(positionals[4]) if len(positionals) > 4 else today.month

    main(
        COMPANY,
        DATA_DIR,
        OUT_DIR,
        YEAR,
        MONTH,
        carry_forward=carry_forward,
        no_carry_extra=no_carry_extra,
        carry_all=carry_all,
    )
