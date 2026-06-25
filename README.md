# מיכפל Salary Import Template Automation

## Files

- `engine.py` — all extraction / generation / insertion logic (pure functions, no I/O side effects beyond reading/writing the files it's explicitly told to touch). Used by both front ends below.
- `gen_company.py` — original CLI workflow, now built on `engine.py`. Generates the `.txt` template and `.xlsx` only; does **not** insert into `Q8SRGL26.000`.
- `app.py` — Streamlit wizard UI: select company → extract & generate → insert (with automatic backup, replacing any existing block for that company).
- `config.py` — persistent settings (data/output folders) stored in `config.json`; OS-aware defaults.
- `launch.py` — cross-platform launcher: starts Streamlit and opens the browser.
- `requirements.txt` — pinned dependencies.
- `deploy/` — Windows installer build (Inno Setup wizard + bundled Python). See `deploy/README.md`.

## Setup (development)

```bash
pip install -r requirements.txt --break-system-packages
```

(Drop `--break-system-packages` if you're using a virtualenv.)

For the packaged Windows installer, see `deploy/README.md` instead — end users do not run pip.

## Running the UI

```bash
streamlit run app.py
# or, via the launcher (picks a port, opens the browser, avoids a second server):
python3 launch.py
```

This opens a browser tab (default `http://localhost:8501`). The sidebar **Data folder** and **Output folder** persist between runs (saved to `config.json`); they default to `/mnt/Z/Msk8` on Linux and `Z:\Msk8` on Windows.

## Running the old CLI

```bash
python3 gen_company.py [company] [data_dir] [out_dir] [year] [month]
# all args optional (data_dir/out_dir default to the saved config), e.g.:
python3 gen_company.py 083
```

## Wizard flow (app.py)

1. **Select company** — only companies with both a `Q8MIFL26.[n]` and `Q8OVDM26.[n]` file are listed. Shows an info note if a template with that company's name already exists in `Q8SRGL26.000` (it will be replaced, not duplicated).
2. **Extract & generate** — runs the binary extraction and immediately builds the `@@QD@@` text block and the `.xlsx`. Shows component/employee counts and tables plus diagnostics (filtered-out records, invalid ID blocks). The generated block and column mapping are collapsed by default; download the Excel and the `.txt` to inspect. Nothing is written to `Q8SRGL26.000` at this point.
3. **Insert** — separate, explicit step requiring a confirmation checkbox. Uses `regenerate_and_insert_template`: if a block with this company's template name already exists it is **removed first** (backed up to `…bak-{company}-remove-{timestamp}`), then the new block is inserted (backed up to `…bak-{company}-{timestamp}`) before the `@@XL@@` marker. The insert backup path is reported so a manual rollback is always one copy away.

## Notes

- The core extraction/insertion pipeline has been validated against the live מיכפל UI (see `CLAUDE.md` and the project-knowledge doc). When trying a brand-new company, still cross-check its component list against the UI with `ui_check.py` before trusting the output.
- Re-inserting the same company **replaces** its block rather than appending a duplicate. The duplicate-name note in Step 1 is informational, not a block.
- Insertion writes to the data folder, so the user running the app needs write access to it. On the Linux dev box that share typically needs elevation (`sudo .venv/bin/streamlit run app.py`); on the Windows deployment the user must have write access to the share. A permission failure is reported clearly and leaves the file unmodified.
