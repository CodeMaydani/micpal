"""
engine.py
Core extraction / generation / insertion logic for the מיכפל salary
import automation.

Pure functions only -- no printing, no hardcoded paths. Both
gen_company.py (CLI) and app.py (Streamlit UI) import from here so
behavior is identical between the two front ends.
"""

import datetime
import glob
import os
import re
import struct

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def col_letter(n):
    """Convert 1-based column index to Excel letter(s): 1->A, 27->AA, etc."""
    r = ""
    while n:
        n, x = divmod(n - 1, 26)
        r = chr(65 + x) + r
    return r


def valid_israeli_id(n):
    """Luhn-variant check for Israeli 9-digit ID numbers."""
    s = str(n).zfill(9)
    total = 0
    for i, c in enumerate(s):
        d = int(c) * (1 if i % 2 == 0 else 2)
        total += d - 9 if d > 9 else d
    return total % 10 == 0


# Excel/XML 1.0 forbids most control characters in cell strings. Allowed:
# tab (\x09), newline (\x0A), carriage return (\x0D), and the printable
# ranges above \x20 (excluding the surrogate/illegal blocks). openpyxl
# raises IllegalCharacterError on anything outside this set.
_ILLEGAL_XML_CHARS_RE = re.compile(
    "[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x84\x86-\x9f\ufdd0-\ufddf\ufffe\uffff]"
)


def sanitize_text(s):
    """Strip characters that are illegal in Excel/XML cell values."""
    return _ILLEGAL_XML_CHARS_RE.sub("", s).strip()


# ---------------------------------------------------------------------------
# Company discovery / template inspection
# ---------------------------------------------------------------------------


def discover_companies(data_dir):
    """
    Find company codes that have BOTH a Q8MIFL26.[company] and a
    Q8OVDM26.[company] file in data_dir.

    Returns a sorted list of company-code strings exactly as they appear
    in the filename suffix (leading zeros preserved).
    """
    mifl = {
        os.path.basename(p).split(".", 1)[1]
        for p in glob.glob(os.path.join(data_dir, "Q8MIFL26.*"))
    }
    ovdm = {
        os.path.basename(p).split(".", 1)[1]
        for p in glob.glob(os.path.join(data_dir, "Q8OVDM26.*"))
    }
    return sorted(mifl & ovdm)


def list_existing_templates(srgl_path):
    """
    Parse @N lines out of Q8SRGL26.000 and return the template names.
    Used for duplicate-name warnings before generating a new template.
    """
    with open(srgl_path, "rb") as f:
        content = f.read()
    text = content.decode("iso-8859-8", errors="replace")
    return [name.strip() for name in re.findall(r"^@N(.+)$", text, re.MULTILINE)]


# ---------------------------------------------------------------------------
# Binary extraction
# ---------------------------------------------------------------------------


def extract_components(path):
    """
    Read salary components from Q8MIFL26.[company].
    Returns (results, skipped_count).
      results       -- list of (rechiv_extracted, name, kod_mahk)
      skipped_count -- candidate records that looked like real components
                       but failed a filter (diagnostic only)

    NOTE: actual מיכפל code = rechiv_extracted - 1.

    NUMBERING FIX (was: rechiv_extracted = tail[10], the "-1 offset
    invariant"): tail[10] of a record is actually the NEXT record's
    number, not this one's -- it only looked like "current+1" because
    that holds whenever component numbering has no gaps. Verified against
    the live מיכפל UI across companies 003/004/083: the correct number is
    the PREVIOUS record's tail[10], read directly as
    data[name_offset - 19]. The first record has no predecessor, so it
    falls back to tail[10] itself (always correct there, since משכורת is
    always record 0).

    NOISE-RECORD GUARD (new): company 003's file contains a few garbled
    pseudo-records past the real component list that satisfied the old
    filters by coincidence. Excluded via looks_like_real_component_name()
    -- a real name is composed almost entirely of Hebrew letters, digits,
    and common punctuation; verified with zero false rejections across
    189 known-real components and 3/3 known garbage records caught.
    """
    data = open(path, "rb").read()
    TAIL_SIZE = 29
    START = 0x4E2E
    MAX_REXIVIM = 252  # confirmed in מיכפל's own UI: max רכיבי שכר per company
    results = []
    skipped = 0
    offset = START
    first = True
    while offset < len(data) - TAIL_SIZE:
        try:
            name_end = data.index(b"\x00", offset)
        except ValueError:
            break
        raw_name = data[offset:name_end].decode("iso-8859-8", errors="replace").strip()
        name = sanitize_text(raw_name)
        tail = data[name_end + 1 : name_end + 1 + TAIL_SIZE]
        if len(tail) < TAIL_SIZE:
            break
        t10 = tail[10]
        kod_mahk = tail[-1]

        # --- numbering fix: real number = previous record's tail[10] ---
        if first:
            rechiv_extracted = t10
        else:
            rechiv_extracted = data[offset - 19] + 1
        # -----------------------------------------------------------------

        actual = rechiv_extracted - 1
        if (
            name
            and t10 > 0
            and t10 < 200
            and 0 < actual <= MAX_REXIVIM
            and "\ufffd" not in name
            and name != "?" * len(name)
            and kod_mahk < 100
            and looks_like_real_component_name(name)  # noise-record guard
        ):
            results.append((rechiv_extracted, name, kod_mahk))
        elif raw_name and t10 > 0 and t10 < 200:
            skipped += 1
        first = False
        offset = name_end + 1 + TAIL_SIZE
    return results, skipped


_ALLOWED_PUNCT = set(" .,'\"%-~$()/+")
_MAX_NOISE_FRACTION = 0.15


def looks_like_real_component_name(name):
    """
    Content-based filter distinguishing real component names from garbage
    pseudo-records found past the end of a company's real component
    section (e.g. company 003 near the documented 0xCCCA GetLength-clamp
    boundary). An offset-gap approach was tried first and rejected: the
    transition isn't one clean jump -- the parser walks through many
    small steps (50-220 bytes) of padding-as-pseudo-records before
    reaching the actual garbage, so no gap-size threshold reliably marks
    the boundary. This content-based check works instead: a real name is
    built almost entirely from Hebrew letters, digits, and a small set of
    common punctuation; garbage records are dominated by other characters
    even when 1-2 Hebrew letters appear by coincidental byte alignment.
    """
    if not name or len(name) > 40:
        return False
    has_hebrew = any("\u05d0" <= ch <= "\u05ea" for ch in name)
    if not has_hebrew:
        return False
    noise = sum(
        1
        for ch in name
        if not ("\u05d0" <= ch <= "\u05ea" or ch.isdigit() or ch in _ALLOWED_PUNCT)
    )
    return (noise / len(name)) <= _MAX_NOISE_FRACTION


def extract_employees(path):
    """
    Read employee list from Q8OVDM26.[company].

    Returns (results, invalid_count).
      results       -- list of (tz_str, full_name) sorted by תעודת זהות
      invalid_count -- employee blocks with an invalid תז (diagnostic only)
    """
    data = open(path, "rb").read()
    BLOCK = 167936
    results = []
    total_blocks = len(data) // BLOCK
    invalid_ids = 0
    for b in range(total_blocks):
        block = data[b * BLOCK : (b + 1) * BLOCK]
        end = block.index(b"\x00", 9)
        last_name = sanitize_text(
            block[9:end].decode("iso-8859-8", errors="replace").strip()
        )
        end2 = block.index(b"\x00", 29)
        first_name = sanitize_text(
            block[29:end2].decode("iso-8859-8", errors="replace").strip()
        )
        tz_raw = struct.unpack_from("<I", block, 59)[0]
        if valid_israeli_id(tz_raw):
            tz = str(tz_raw).zfill(9)
        else:
            tz = ""
            invalid_ids += 1
        full_name = f"{first_name} {last_name}".strip()
        results.append((tz, full_name))

    results.sort(key=lambda x: x[0])
    return results, invalid_ids


# ---------------------------------------------------------------------------
# Template builder (@@QD@@ block for Q8SRGL26.000)
# ---------------------------------------------------------------------------


def build_template(template_name, components):
    """
    Build the @@QD@@ text block for insertion into Q8SRGL26.000.

    Returns (template_text, col_map, stats_cols).
      col_map    -- list of (actual_code, name, col_כמות, col_מחיר)
      stats_cols -- list of (field_code, label, col_index)
    """
    lines = []
    lines.append(f"@N{template_name}")
    lines.append("@V90072")
    lines.append("@YB2 C2 A2")
    lines.append("@F4")
    lines.append("@H")
    lines.append("@D-1 1 1 1")
    lines.append("@D1 -1 -1 1")
    lines.append("@D1 7 4 1")  # סכום -> col D=4
    lines.append("@D1 8 3 1")  # ברוטו/נטו -> col C=3
    lines.append("@Mבב")
    lines.append("@M* אין מיפוי *ג")
    lines.append("@Mננ")
    lines.append("@M* אין מיפוי *ע")
    lines.append("@M* אין מיפוי *פ")
    lines.append("@Mקק")
    # רכיב 001 (משכורת) fixed dual-entry at E=5, F=6
    lines.append("@D1 -1 -1 1")
    lines.append("@D1 5 5 1")
    lines.append("@D1 6 6 1")
    lines.append("@D1 8 3 1")

    col_map = []
    col = 7  # remaining components start at G=7
    for rechiv_extracted, name, kod_mahk in components:
        if rechiv_extracted == 2:  # משכורת already covered in fixed block
            continue
        actual = rechiv_extracted - 1  # THE -1 OFFSET -- universally verified
        col_k, col_m = col, col + 1
        col_map.append((actual, name, col_k, col_m))
        lines.append(f"@D{actual} -1 -1 1")
        lines.append(f"@D{actual} 5 {col_k} 1")
        lines.append(f"@D{actual} 6 {col_m} 1")
        lines.append(f"@D{actual} 8 3 1")
        col += 2

    stats = [
        (10, "ימי עבודה משולמים"),
        (14, "ימי עבודה בפועל"),
        (15, "שעות בפועל"),
        (16, "תקן ימים"),
        (17, "תקן שעות"),
        (18, "חופשה"),
        (19, "מחלה"),
    ]
    stats_cols = []
    for field, label in stats:
        lines.append(f"@D-1 {field} {col} 1")
        stats_cols.append((field, label, col))
        col += 1

    return "\n".join(lines) + "\n", col_map, stats_cols


# ---------------------------------------------------------------------------
# Excel builder
# ---------------------------------------------------------------------------


def build_excel(
    company, components, col_map, stats_cols, employees, out_path, year, month
):
    """
    Generate the .xlsx import template.

    Rows 1-2 : metadata (חברה / שנת מס / חודש דיווח)
    Row 4    : column headers
    Row 5+   : employees (תעודת זהות pre-filled in col A, name in col B)
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "ייבוא משכורות"
    ws.sheet_view.rightToLeft = True

    # --- styles ---------------------------------------------------------
    bold = Font(bold=True, name="Arial", size=10)
    normal = Font(name="Arial", size=10)
    header_fill = PatternFill("solid", start_color="D9E1F2")
    comp_fill = PatternFill("solid", start_color="E2EFDA")
    stats_fill = PatternFill("solid", start_color="FFF2CC")
    rechiv1_fill = PatternFill("solid", start_color="FCE4D6")
    emp_fill = PatternFill("solid", start_color="F2F2F2")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="right", vertical="center")
    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def style(cell, font=normal, fill=None, align=center):
        cell.font = font
        if fill:
            cell.fill = fill
        cell.alignment = align
        cell.border = border

    def safe_set(cell, value):
        """Set a cell value, sanitizing strings as a final safety net."""
        if isinstance(value, str):
            value = sanitize_text(value)
        cell.value = value
        return cell

    # --- rows 1-2: metadata ---------------------------------------------
    style(safe_set(ws["A1"], "חברה"), font=bold, fill=header_fill)
    style(safe_set(ws["B1"], "שנת מס"), font=bold, fill=header_fill)
    style(safe_set(ws["C1"], "חודש דיווח"), font=bold, fill=header_fill)

    style(safe_set(ws["A2"], int(company)))
    style(safe_set(ws["B2"], int(year)))  # read by @YB2 C2 A2
    style(safe_set(ws["C2"], int(month)))  # read by @YB2 C2 A2

    # --- row 4: column headers ------------------------------------------
    for col_idx, label in [
        (1, "מספר עובד"),
        (2, "שם עובד"),
        (3, "ברוטו/נטו"),
        (4, "סכום"),
    ]:
        style(safe_set(ws.cell(4, col_idx), label), font=bold, fill=header_fill)

    first_comp_name = components[0][1] if components else "משכורת"
    style(
        safe_set(ws.cell(4, 5), f"{first_comp_name}\nכמות"),
        font=bold,
        fill=rechiv1_fill,
    )
    style(
        safe_set(ws.cell(4, 6), f"{first_comp_name}\nמחיר"),
        font=bold,
        fill=rechiv1_fill,
    )

    for actual, name, col_k, col_m in col_map:
        style(safe_set(ws.cell(4, col_k), f"{name}\nכמות"), font=bold, fill=comp_fill)
        style(safe_set(ws.cell(4, col_m), f"{name}\nמחיר"), font=bold, fill=comp_fill)

    for field, label, col_idx in stats_cols:
        style(safe_set(ws.cell(4, col_idx), label), font=bold, fill=stats_fill)

    # --- rows 5+: employees ---------------------------------------------
    for row_idx, (tz, full_name) in enumerate(employees, 5):
        style(safe_set(ws.cell(row_idx, 1), tz), fill=emp_fill)
        style(safe_set(ws.cell(row_idx, 2), full_name), fill=emp_fill, align=left)

    # --- column widths --------------------------------------------------
    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 10
    ws.column_dimensions["E"].width = 10
    ws.column_dimensions["F"].width = 10
    for actual, name, col_k, col_m in col_map:
        ws.column_dimensions[col_letter(col_k)].width = 10
        ws.column_dimensions[col_letter(col_m)].width = 10
    for field, label, col_idx in stats_cols:
        ws.column_dimensions[col_letter(col_idx)].width = 13

    ws.row_dimensions[4].height = 45
    ws.freeze_panes = "E5"
    wb.save(out_path)
    return out_path


# ---------------------------------------------------------------------------
# Q8SRGL26.000 backup + insertion
# ---------------------------------------------------------------------------


def regenerate_and_insert_template(srgl_path, template_name, template_text, company):
    """
    Idempotently (re)place a company's template block in Q8SRGL26.000:
    remove any existing block with this exact name first, then insert
    the new template_text. Always backs up before any modification.

    Returns a dict:
        {
            "removed":       bool,           # was an old block found+removed?
            "remove_backup": str or None,     # backup path from the removal step
            "insert_backup": str,             # backup path from the insertion step
            "removed_block": str or None,     # the old block's text, for review
        }

    Raises ValueError if the template name is found more than once
    BEFORE touching the file at all (refuses to guess).
    """
    existing = list_existing_templates(srgl_path)
    matches = [n for n in existing if n == template_name]

    if len(matches) > 1:
        raise ValueError(
            f"Template {template_name!r} found {len(matches)} times in "
            f"{srgl_path} -- refusing to regenerate until the file is "
            f"de-duplicated by hand."
        )

    result = {
        "removed": False,
        "remove_backup": None,
        "insert_backup": None,
        "removed_block": None,
    }

    if len(matches) == 1:
        remove_backup, removed_block = remove_template(srgl_path, template_name)
        result["removed"] = True
        result["remove_backup"] = remove_backup
        result["removed_block"] = removed_block

    insert_backup = backup_and_insert(srgl_path, template_text, company)
    result["insert_backup"] = insert_backup

    return result


def remove_template(srgl_path, template_name):
    """
    Back up Q8SRGL26.000, then remove the single @N<template_name> block
    (from its @N line up to the next @N line or @@XL@@, whichever comes
    first) from the file.

    Raises ValueError BEFORE touching the original file if template_name
    is not found, or if it is found more than once (ambiguous -- this
    function will not guess which one to remove).

    Returns (backup_path, removed_block_text) on success.
    """
    if not os.path.exists(srgl_path):
        raise FileNotFoundError(f"{srgl_path} not found")

    with open(srgl_path, "rb") as f:
        content = f.read()

    text = content.decode("iso-8859-8", errors="replace")
    lines = text.splitlines(keepends=True)

    target_line = f"@N{template_name}"
    match_indices = [
        i for i, line in enumerate(lines) if line.rstrip("\r\n") == target_line
    ]

    if len(match_indices) == 0:
        raise ValueError(f"Template {template_name!r} not found -- nothing removed.")
    if len(match_indices) > 1:
        raise ValueError(
            f"Template {template_name!r} found {len(match_indices)} times -- "
            f"refusing to guess which to remove."
        )

    start = match_indices[0]
    end = len(lines)
    for j in range(start + 1, len(lines)):
        s = lines[j].rstrip("\r\n")
        if s.startswith("@N") or s.startswith("@@XL@@"):
            end = j
            break

    removed_block = "".join(lines[start:end])

    # --- backup BEFORE any modification, same pattern as backup_and_insert ---
    company_hint = "".join(ch for ch in template_name if ch.isdigit()) or "removal"
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = f"{srgl_path}.bak-{company_hint}-remove-{timestamp}"

    with open(backup_path, "wb") as f:
        f.write(content)
    with open(backup_path, "rb") as f:
        if f.read() != content:
            raise IOError("Backup verification failed -- original file left untouched.")
    # ---------------------------------------------------------------------------

    new_text = "".join(lines[:start]) + "".join(lines[end:])
    new_content = new_text.encode("iso-8859-8")

    with open(srgl_path, "wb") as f:
        f.write(new_content)

    return backup_path, removed_block


def backup_and_insert(srgl_path, template_text, company):
    """
    Back up Q8SRGL26.000, then insert template_text (the @@QD@@ block)
    immediately before the @@XL@@ marker.

    Backup naming:
        {srgl_path}.bak-{company zero-padded to 3 digits}-{YYYYMMDD_HHMMSS}

    All validation happens BEFORE the original file is touched. The backup
    is written and verified first; the marker is checked next; only then
    is the original file overwritten. Returns the backup path on success.
    """
    if not os.path.exists(srgl_path):
        raise FileNotFoundError(f"{srgl_path} not found")

    company_padded = str(company).zfill(3)
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = f"{srgl_path}.bak-{company_padded}-{timestamp}"

    with open(srgl_path, "rb") as f:
        content = f.read()

    # Backup BEFORE any modification attempt.
    with open(backup_path, "wb") as f:
        f.write(content)

    # Verify the backup actually matches before proceeding.
    with open(backup_path, "rb") as f:
        if f.read() != content:
            raise IOError("Backup verification failed -- original file left untouched.")

    marker = b"@@XL@@"
    pos = content.find(marker)
    if pos == -1:
        raise ValueError("Marker @@XL@@ not found -- refusing to modify file.")

    template_encoded = template_text.encode("iso-8859-8")
    new_content = content[:pos] + template_encoded + b"\n" + content[pos:]

    with open(srgl_path, "wb") as f:
        f.write(new_content)

    return backup_path
