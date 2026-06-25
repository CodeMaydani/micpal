# Windows build & rollout test checklist

Run this on a **Windows 10/11 (x64)** machine after building the installer
(see `README.md`). The Linux dev box cannot validate the installer, the VBS
launchers, or the bundled Python — all of that is verified here.

## A. Build the installer

- [ ] `./deploy/build_windows.ps1` completes with no errors.
- [ ] `deploy\build\python\python.exe` exists.
- [ ] `deploy\build\python\python.exe -m streamlit --version` prints a version.
- [ ] `deploy\build\app\` contains `app.py`, `engine.py`, `gen_company.py`,
      `config.py`, `launch.py`, `ui_check.py`, `requirements.txt`, `README.md`.
- [ ] Inno Setup builds `deploy\Output\micpal-setup-<version>.exe` with no errors.

## B. Clean-machine install (ideally a fresh VM with **no** Python installed)

- [ ] Run `setup.exe`; wizard completes, installs to `Z:\Micpal`.
- [ ] No admin/UAC prompt was required (per-user install).
- [ ] Start Menu has: launch, **Stop**, and Uninstall shortcuts.
- [ ] Desktop icon present if that task was ticked.

## C. First run

- [ ] Launch shortcut opens the browser to the app (no console window lingers).
- [ ] Sidebar **Data folder** defaults to `Z:\Msk8`.
- [ ] Change the data/output folders; confirm `Z:\Micpal\config.json` is written.
- [ ] Close everything, relaunch — the changed folders persist.

## D. End-to-end against a TEST copy of the data share

> Point the data folder at a **copy** of `Z:\Msk8`, not production, for this.

- [ ] Step 1: company list populates from the data folder.
- [ ] Step 2: Extract & generate — component/employee counts look right; Excel
      and `@@QD@@` block download.
- [ ] Step 3: Insert — `regenerate_and_insert_template` runs; success message
      with a backup path; backup file exists next to `Q8SRGL26.000`.
- [ ] Re-insert the same company — confirm it replaces (one block, no duplicate)
      via the template list, not appends.
- [ ] If the share is read-only for this user, confirm the permission error
      message is the friendly one, then fix share perms (not code).

## E. Stop & uninstall

- [ ] **Stop** shortcut ends the app (browser can no longer reach it; no
      `pythonw.exe` running `launch.py`/`streamlit` remains in Task Manager).
- [ ] **Stop** when nothing is running shows the "not running" message.
- [ ] Uninstall removes `Z:\Micpal` app files and `config.json`.
- [ ] Uninstall does **not** touch `Z:\Msk8` or the generated
      `company_templates` folder.

## Notes / sign-off

- Tester: ____________________  Date: ____________
- Build version: ____________  Python runtime tag: ____________
- Issues found:
