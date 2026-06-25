#!/usr/bin/env python3
"""
ui_check.py  --  generic per-company cross-check tool

Purpose
-------
Print every component for ONE company, formatted EXACTLY like מיכפל's own
dropdown ("%03d - name", confirmed from MxpEmpIE.dll's format string), in
the SAME order the file stores them. This lets you put the script's output
and the live dropdown side by side and diff by eye with zero mental
reformatting.

It shows BOTH the OLD (pre-fix) and NEW (corrected) number for every line,
so a mismatch against the UI is self-explanatory: either NEW matches the
UI (confirms the fix) or it doesn't (tells us something new).

It also writes a timestamped, full log file per run (every record, every
filter decision, both numbering rules) to ./ui_check_logs/ -- so nothing
needs to be re-derived later if a question comes up about a specific
company or component.

Usage
-----
    python3 ui_check.py <company_number> [--data-dir /mnt/Z/Msk8] [--show-skipped]

Examples
    python3 ui_check.py 083
    python3 ui_check.py 004 --data-dir /mnt/Z/Msk8
    python3 ui_check.py 176 --show-skipped

Output
------
  1. A printed table to the terminal: idx, OLD, NEW, "%03d - name" (NEW form)
  2. A full log file: ui_check_logs/ui_check_<company>_<timestamp>.log
     containing every record (including ones filtered out) with raw tail
     bytes, for later inspection if something doesn't match.

What to do with the output
---------------------------
See the companion file UI_CHECK_INSTRUCTIONS.md for the exact step-by-step
procedure (what to open in מיכפל, what to copy, how to compare).
"""

import argparse
import datetime
import os
import re
import sys

_ILLEGAL = re.compile(
    "[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x84\x86-\x9f\ufdd0-\ufddf\ufffe\uffff]"
)


def clean(s):
    return _ILLEGAL.sub("", s).strip()


def parse_file(path):
    """
    Parse Q8MIFL26.[company] and return a list of record dicts, in file
    order, with BOTH numbering rules computed:
      old_number = tail[10] - 1                     (gen_company's original rule)
      new_number = data[name_offset - 19]            (corrected rule; first
                                                        record falls back to old)
    Every record is returned, including ones that fail the validity filter
    (marked valid=False), so the log is complete -- filtering happens only
    at display/report time.
    """
    data = open(path, "rb").read()
    START, TAIL = 0x4E2E, 29

    records = []
    offset = START
    first = True
    while offset < len(data) - TAIL:
        try:
            name_end = data.index(b"\x00", offset)
        except ValueError:
            break
        raw_name = data[offset:name_end].decode("iso-8859-8", errors="replace").strip()
        name = clean(raw_name)
        tail = data[name_end + 1 : name_end + 1 + TAIL]
        if len(tail) < TAIL:
            break

        t10 = tail[10]
        kod_mahk = tail[-1]
        old_number = t10 - 1
        if first:
            new_number = t10 - 1
        else:
            new_number = (
                data[offset - 19] + 1 - 1
            )  # = data[offset-19], spelled out for clarity

        valid = bool(
            name
            and 0 < t10 < 200
            and "\ufffd" not in name
            and name != "?" * len(name)
            and kod_mahk < 100
        )

        records.append(
            {
                "file_offset": offset,
                "raw_name": raw_name,
                "name": name,
                "tail_hex": tail.hex(),
                "t10": t10,
                "kod_mahk": kod_mahk,
                "old_number": old_number,
                "new_number": new_number,
                "valid": valid,
            }
        )

        first = False
        offset = name_end + 1 + TAIL

    return records


def main():
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter
    )
    ap.add_argument(
        "company", help="Company number, e.g. 083 or 4 (zero-padding optional)"
    )
    ap.add_argument(
        "--data-dir",
        default="/mnt/Z/Msk8",
        help="Directory containing Q8MIFL26.[company] (default: /mnt/Z/Msk8)",
    )
    ap.add_argument(
        "--show-skipped",
        action="store_true",
        help="Also print records that failed the validity filter",
    )
    ap.add_argument(
        "--log-dir",
        default="./ui_check_logs",
        help="Where to write the full per-run log file (default: ./ui_check_logs)",
    )
    args = ap.parse_args()

    company = args.company.strip()
    # Accept "4" or "083" etc. -- pad to 3 digits to match the file naming convention.
    company_padded = company.zfill(3)
    path = os.path.join(args.data_dir, f"Q8MIFL26.{company_padded}")

    if not os.path.exists(path):
        # also try unpadded, in case the company number isn't 3 digits in this install
        alt = os.path.join(args.data_dir, f"Q8MIFL26.{company}")
        if os.path.exists(alt):
            path = alt
        else:
            print(f"ERROR: could not find {path} (also tried {alt})", file=sys.stderr)
            sys.exit(1)

    records = parse_file(path)

    os.makedirs(args.log_dir, exist_ok=True)
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = os.path.join(args.log_dir, f"ui_check_{company_padded}_{ts}.log")

    with open(log_path, "w", encoding="utf-8") as log:
        log.write(f"ui_check.py log\n")
        log.write(f"company:   {company_padded}\n")
        log.write(f"file:      {path}\n")
        log.write(f"timestamp: {ts}\n")
        log.write(f"total records parsed: {len(records)}\n\n")
        log.write(
            f"{'idx':>4} {'valid':>5} {'off':>7} {'t10':>4} {'mahk':>4} "
            f"{'old#':>5} {'new#':>5}  name | tail_hex\n"
        )
        for i, r in enumerate(records):
            log.write(
                f"{i:>4} {str(r['valid']):>5} {r['file_offset']:#07x} "
                f"{r['t10']:>4} {r['kod_mahk']:>4} {r['old_number']:>5} "
                f"{r['new_number']:>5}  {r['name']} | {r['tail_hex']}\n"
            )

    print(f"Company {company_padded}  --  file: {path}")
    print(f"Full log written to: {log_path}\n")

    valid_records = [r for r in records if r["valid"]]
    print(
        f"Valid components: {len(valid_records)}  "
        f"(total raw records scanned: {len(records)})\n"
    )

    diffs = [r for r in valid_records if r["old_number"] != r["new_number"]]
    if diffs:
        print(
            f"*** {len(diffs)} component(s) where OLD and NEW differ -- "
            f"these are the ones to prioritize checking against the live UI: ***"
        )
        for r in diffs:
            print(
                f"    OLD={r['old_number']:03d}  NEW={r['new_number']:03d}  "
                f'-- "{r["name"]}"'
            )
        print()
    else:
        print(
            "No OLD/NEW differences in this company (gap-free numbering) -- "
            "still worth spot-checking a couple of entries below.\n"
        )

    print("=" * 50)
    print("COPY FROM HERE — paste next to the live מיכפל dropdown")
    print("=" * 50)
    print(f"{'#':>3}  {'NEW (use this)':<14} {'OLD (for reference)':<10}")
    for i, r in enumerate(valid_records):
        marker = "  <-- CHECK THIS ONE" if r["old_number"] != r["new_number"] else ""
        print(
            f"{i:>3}  {r['new_number']:03d} - {r['name']:<20} "
            f"(old: {r['old_number']:03d}){marker}"
        )
    print("=" * 50)
    print("COPY ENDS HERE")
    print("=" * 50)

    if args.show_skipped:
        skipped = [r for r in records if not r["valid"]]
        if skipped:
            print(
                f"\n--- {len(skipped)} skipped/invalid record(s) (--show-skipped) ---"
            )
            for r in skipped:
                print(
                    f"    off={r['file_offset']:#x}  t10={r['t10']}  "
                    f"mahk={r['kod_mahk']}  raw_name={r['raw_name']!r}"
                )


if __name__ == "__main__":
    main()
