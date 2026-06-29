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
- `pip install -r requirements.txt --break-system-packages` — deps are now
  **pinned** (`streamlit==1.58.0`, `openpyxl==3.1.5`) for reproducible builds.
- **`sudo` drops the venv.** If a command needs `sudo` (writing to
  `/mnt/Z/Msk8/Q8SRGL26.000` usually does), call the venv interpreter
  explicitly: `sudo ~/workspace/micpal-app/.venv/bin/python3 -c "..."`
  — plain `sudo python3` will fail with `ModuleNotFoundError: openpyxl`.
  In the Streamlit UI this is why insertion is launched with
  `sudo ~/workspace/micpal-app/.venv/bin/streamlit run app.py`.
- **Paths are config-driven, never hardcoded.** `config.py` owns defaults
  (OS-aware: `/mnt/Z/Msk8` on Linux, `Z:\Msk8` on Windows) and persists
  user changes to `config.json` next to the app. Read paths via
  `config.load()`, not literals. `config.json` is git-ignored (per-machine).

## Files
- `engine.py` — **the only file with real logic.** Pure library, no
  printing, no hardcoded paths. Both `gen_company.py` and `app.py`
  import from it. Edit logic here, nowhere else.
- `gen_company.py` — CLI entry point: `main(company, data_dir, out_dir,
  year, month)`. Defaults for `data_dir`/`out_dir` come from `config.py`.
  Generates files only — does **not** insert into `Q8SRGL26.000`.
- `app.py` — Streamlit UI wrapper. **3-step flow** (was 4): (1) select
  company, (2) extract & generate — merged, one button, (3) insert.
  Insertion goes through `regenerate_and_insert_template` (replace-safe).
  Audited against `README.md` and reconciled — no longer a TODO.
- `config.py` — path defaults + persistence (see Stack).
- `launch.py` — cross-platform launcher: picks a port, starts Streamlit
  headless, opens the browser, skips starting a second server if one is up.
- `deploy/` — **Windows installer build** (Inno Setup wizard + bundled
  standalone Python). See Deployment section and `deploy/README.md`.
- `archive/` — one-off reverse-engineering / investigation scripts and
  dumps, kept for reference, git-ignored. NOT part of deployment. (This
  is where `diff_ui.py`, `scan_companies.py`, etc. now live.)
- Data lives at `/mnt/Z/Msk8/` on Linux (`Z:\Msk8` on Windows) — NOT the
  מיכפל install folder, which is separate and irrelevant to this pipeline.
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
3. **Build & test the Windows installer** on a Windows machine — run
   `deploy/TESTING.md` end to end (build → install → run → stop →
   uninstall). The build artifacts can't be exercised on the Linux dev
   box; this is the only remaining validation gap.

**Done (was on this list):** wiring `regenerate_and_insert_template` into
`app.py` (Step 3 of the UI uses it); auditing `app.py` vs `README.md`
(reconciled — README rewritten to match the 3-step replace-based flow).
`gen_company.py` was intentionally left generation-only (it never
inserts), so no insertion wiring is needed there.

## Deployment (Windows-only)
Decisions are settled (see `deploy/` and project memory):
- **Windows-only.** Linux stays the dev environment; no Linux installer.
- **Install to `Z:\Micpal`** — on the share, user-writable, so `config.json`
  lives next to the app with no permission tricks.
- **Data path set on first run** in the sidebar (defaults to `Z:\Msk8`); the
  installer does not ask for it.
- **Packaging:** bundle a relocatable python-build-standalone CPython +
  pip-installed pinned deps + app files, wrapped by an Inno Setup wizard.
  The build (`deploy/build_windows.ps1` + Inno Setup) **must run on Windows**
  — Windows wheels can't be reliably pip-installed from Linux.
- `launch.py` is the single entry point; `run_micpal.vbs` / `stop_micpal.vbs`
  are the windowless start/stop shortcuts.
- Full build + test instructions: `deploy/README.md` and `deploy/TESTING.md`.

## Validation tools available (don't rebuild these)
- `ui_check.py <company>` — dumps a company's component list in the
  exact format the live מיכפל dropdown uses (`011 - name`), plus a full
  raw-record log to `ui_check_logs/`. Use when checking a new company
  against the live UI. (This one is kept at the repo root.)
- `archive/diff_ui.py <company> <pasted_ui_text.txt>` — automated diff
  between a UI copy-paste and the file-derived list. Now lives under
  `archive/`. Preferred over eyeballing `ui_check.py` output if you have
  copy-paste access to the dropdown. (Screenshots, not copy-paste, were
  used for the validation already done — diff_ui.py wasn't actually
  exercised on real data yet, ui_check output was visually cross-checked
  against screenshots instead.)

## Hard constraints / things that will waste your time if ignored
- The on-disk format of `Q8MIFL26` is NOT MFC `CArchive`-serialized
  (length-prefixed `CString`s). That was a tested hypothesis based on
  DLL disassembly; it produced garbage against real bytes. The working
  model is the original null-terminated-name + 29-byte-tail format —
  don't switch to a different record model without re-validating
  against real bytes first, the way every change in this project has
  been.
- Don't trust visual/manual diffing of two pasted Hebrew text blocks —
  RTL rendering makes this error-prone. Use `archive/diff_ui.py`
  (programmatic) or screenshots cross-checked line-by-line with explicit
  slot numbers, as already done.
- Never call `backup_and_insert()` for a company that might already
  have a template in the file. Always go through
  `regenerate_and_insert_template()` instead, or duplicates accumulate
  silently.
- Don't run `sudo python3` directly in this project — see Stack section.
