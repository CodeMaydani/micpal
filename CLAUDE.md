# CLAUDE.md

## Project
Python automation for **מיכפל** (Michpal), an Israeli payroll system,
operating on its binary data files directly (no GUI automation). Two
outputs: (1) an Excel salary-import template pre-filled with employee
IDs/names, (2) a `@@QD@@` block written into מיכפל's own
`Q8SRGL26.000` template-definitions file so it recognizes the new
import template.

Full background/history: `מיכפל_project_knowledge_v6.md` in this
project. Read it if you need the *why*; this file is the *what to do
now and how*.

## Stack
- Python 3.13, venv at `~/workspace/micpal-app/.venv`
- `pip install -r requirements.txt --break-system-packages` (streamlit,
  openpyxl)
- **`sudo` drops the venv.** If a command needs `sudo` (writing to
  `/mnt/Z/Msk8/Q8SRGL26.000` usually does), call the venv interpreter
  explicitly: `sudo ~/workspace/micpal-app/.venv/bin/python3 -c "..."`
  — plain `sudo python3` will fail with `ModuleNotFoundError: openpyxl`.

## Files
- `engine.py` — **the only file with real logic.** Pure library, no
  printing, no hardcoded paths. Both `gen_company.py` and `app.py`
  import from it. Edit logic here, nowhere else.
- `gen_company.py` — CLI entry point: `main(company, data_dir, out_dir,
  year, month)`.
- `app.py` — Streamlit UI wrapper. Feature-completeness vs. its own
  `README.md`: **not yet verified, treat as TODO.**
- Data lives at `/mnt/Z/Msk8/` (NOT the מיכפל install folder, which is
  separate and irrelevant to this pipeline).
- `Q8MIFL26.[company]` — binary, per-company salary components.
- `Q8OVDM26.[company]` — binary, per-company employee list.
- `Q8SRGL26.000` — ONE shared file, holds every company's template
  definitions. This is what gets modified.

## Status: core pipeline is fixed and validated. Remaining work is rollout, not investigation.

Two real bugs were found and fixed in `engine.py::extract_components()`
this round (full evidence/validation in the project-knowledge doc,
Sections 5 and 11 — don't re-derive, it's settled):

1. **Numbering bug (fixed):** component numbers were read off by one
   record whenever a company's component numbering had a gap (common).
   Old: `actual = tail[10] - 1`. Fixed: `actual = data[name_offset - 19]`
   (previous record's `tail[10]`, read directly; first record falls back
   to its own `tail[10]`). Validated against the live מיכפל UI across 3
   companies (003/004/083), 9/9 corrections confirmed correct.
2. **Noise-record bug (fixed):** a few garbled pseudo-records near the
   `0xCCCA` boundary were slipping through old filters in some
   companies (confirmed in 003). Fixed via
   `looks_like_real_component_name()` — content-based: must contain a
   Hebrew letter, ≤15% non-Hebrew/digit/punctuation noise. Zero false
   rejections across 189 known-real names, 3/3 known garbage rejected.
   A 252-component hard ceiling (confirmed in מיכפל's own UI) is also
   enforced as defense in depth.

**Do not re-investigate either of these. Do not reintroduce an
offset/gap-based approach to the noise filter — it was tried and proven
unreliable** (the real data doesn't have one clean gap to detect on;
see project-knowledge Section 5 if you need the specifics).

`Q8SRGL26.000` insertion is now also safer:
- `remove_template(srgl_path, name)` — removes one named `@N` block.
  Backs up first. Refuses (no file change) if name isn't found exactly
  once.
- `regenerate_and_insert_template(srgl_path, name, text, company)` —
  **use this for all real regeneration**, not `backup_and_insert`
  directly. Removes any existing same-named block first, then inserts.
  Refuses outright, untouched, if the name is already ambiguously
  duplicated in the file.
- `backup_and_insert()` still exists (append-only, no dedup) — only use
  directly if you specifically want append-without-replace behavior.

## Immediate next steps, in order

1. **Finish company 004's reinsertion.** Its old (buggy) template block
   was already manually removed from the live `Q8SRGL26.000` — the file
   currently has zero templates for 004. Generate via `gen_company.py`,
   then call `regenerate_and_insert_template()` to insert. Verify via
   `list_existing_templates()` that exactly one `004` block exists
   afterward.
2. **Repeat for 083, 084, 049, 176** — every company previously
   processed under the old buggy rule. Same pattern each time: check
   current state → `regenerate_and_insert_template` → re-verify exactly
   one block.
3. **Wire `regenerate_and_insert_template()` into `gen_company.py` and
   `app.py`**, replacing any direct call to `backup_and_insert()`, so
   the remove-before-insert safety is automatic, not a manual step a
   human has to remember.
4. **Audit `app.py`** against its own `README.md` for feature
   completeness — not yet checked in this pass.

## Validation tools available (don't rebuild these)
- `ui_check.py <company>` — dumps a company's component list in the
  exact format the live מיכפל dropdown uses (`011 - name`), plus a full
  raw-record log to `ui_check_logs/`. Use when checking a new company
  against the live UI.
- `diff_ui.py <company> <pasted_ui_text.txt>` — automated diff between
  a UI copy-paste and the file-derived list. Preferred over eyeballing
  `ui_check.py` output if you have copy-paste access to the dropdown.
  (Screenshots, not copy-paste, were used for the validation already
  done — diff_ui.py wasn't actually exercised on real data yet, ui_check
  output was visually cross-checked against screenshots instead.)

## Hard constraints / things that will waste your time if ignored
- The on-disk format of `Q8MIFL26` is NOT MFC `CArchive`-serialized
  (length-prefixed `CString`s). That was a tested hypothesis based on
  DLL disassembly; it produced garbage against real bytes. The working
  model is the original null-terminated-name + 29-byte-tail format —
  don't switch to a different record model without re-validating
  against real bytes first, the way every change in this project has
  been.
- Don't trust visual/manual diffing of two pasted Hebrew text blocks —
  RTL rendering makes this error-prone. Use `diff_ui.py` (programmatic)
  or screenshots cross-checked line-by-line with explicit slot numbers,
  as already done.
- Never call `backup_and_insert()` for a company that might already
  have a template in the file. Always go through
  `regenerate_and_insert_template()` instead, or duplicates accumulate
  silently.
- Don't run `sudo python3` directly in this project — see Stack section.
