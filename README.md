# מיכפל Salary Import Template Automation

## Files

- `engine.py` — all extraction / generation / insertion logic (pure functions, no I/O side effects beyond reading/writing the files it's explicitly told to touch). Used by both front ends below.
- `gen_company.py` — original CLI workflow, now built on `engine.py`. Generates the `.txt` template and `.xlsx` only; does **not** insert into `Q8SRGL26.000`.
- `app.py` — Streamlit wizard UI: select company → extract & preview → generate → insert (with automatic backup).
- `requirements.txt` — dependencies.

## Setup

```bash
pip install -r requirements.txt --break-system-packages
```

(Drop `--break-system-packages` if you're using a virtualenv.)

## Running the UI

```bash
streamlit run app.py
```

This opens a browser tab (default `http://localhost:8501`). Set the data folder in the sidebar if it isn't at the default `/mnt/Z/Msk8`.

## Running the old CLI

```bash
python3 gen_company.py [company] [data_dir] [out_dir] [year] [month]
# all args optional, e.g.:
python3 gen_company.py 083
```

## Wizard flow (app.py)

1. **Select company** — only companies with both a `Q8MIFL26.[n]` and `Q8OVDM26.[n]` file are listed. Warns if a template with that company's name already exists in `Q8SRGL26.000`.
2. **Extract & preview** — runs the binary extraction, shows component/employee counts and tables, plus diagnostics (filtered-out records, invalid ID blocks).
3. **Generate** — builds the `@@QD@@` text block and the `.xlsx` file. Nothing is written to `Q8SRGL26.000` at this point. Download both to inspect before proceeding.
4. **Insert** — separate, explicit step. Requires checking a confirmation box. Creates a backup at:
   ```
   Q8SRGL26.000.bak-{company, zero-padded to 3 digits}-{YYYYMMDD_HHMMSS}
   ```
   before writing, verifies the backup matches the original, confirms the `@@XL@@` marker exists, then inserts. Reports the exact backup path so a manual rollback is always one `cp` away.

## Notes / things to verify on your end

- I don't have access to `/mnt/Z/Msk8` or any real `Q8MIFL26.*` / `Q8OVDM26.*` / `Q8SRGL26.000` files in this environment, so this hasn't been run against real binary data — only checked for logical consistency against the documented, already-validated `gen_company.py` behavior. Worth a first run against a company you've already validated (083, 084, 049, or 176) to confirm `engine.py` produces byte-identical output to what the original script produced for that company, before trusting it on a new company.
- Duplicate-template-name detection is a warning only, not a hard block — by design, in case you want to intentionally regenerate one.
