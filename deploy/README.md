# Deployment (Windows)

The מיכפל app ships as a Windows installer (`setup.exe`) built with
[Inno Setup](https://jrsoftware.org/isinfo.php). It bundles its own Python, so
the target machine needs **no** pre-installed Python.

Deployment is Windows-only. Linux remains the development environment.

## What the user gets

- Installs to `Z:\Micpal` (on the share, so it's user-writable and `config.json`
  can live next to the app).
- A Start Menu shortcut (and optional desktop icon) that launches the app
  silently and opens the browser UI.
- On **first run** the user sets the data folder in the sidebar; it defaults to
  `Z:\Msk8`. The installer does not ask for it.

## How to build (on a Windows machine)

You need: Windows 10/11 (x64), [Inno Setup 6](https://jrsoftware.org/isdl.php),
and internet access for the first build (to download the Python runtime).

1. Clone/copy the repo onto the Windows machine.
2. In PowerShell, from the repo root:
   ```powershell
   ./deploy/build_windows.ps1
   ```
   This downloads a relocatable CPython, `pip install`s the pinned
   `requirements.txt` into it, and stages the app under `deploy\build\`.
3. Open `deploy\micpal.iss` in Inno Setup and **Build** (Ctrl+F9).
4. The installer is written to `deploy\Output\micpal-setup-<version>.exe`.

> The build **must** run on Windows: pip fetches platform-specific wheels
> (streamlit pulls in pyarrow, etc.), which have to be the Windows builds.

## Files in this folder

| File                | Purpose                                                       |
|---------------------|--------------------------------------------------------------|
| `build_windows.ps1` | Stages bundled Python + deps + app files into `build\`.       |
| `micpal.iss`        | Inno Setup wizard definition (produces `setup.exe`).         |
| `run_micpal.vbs`    | Silent launcher the shortcut points at (no console window).  |

`build/` and `Output/` are build outputs and are git-ignored.

## Stopping the app

The launcher runs Streamlit in the background (windowless `pythonw.exe`). Use
the **Stop Michpal Template Automation** Start Menu shortcut to stop it — it
terminates the launcher and its Streamlit child (and only those). Closing the
browser tab alone leaves the server idling in the background.

## Bumping versions

- App version: `#define AppVersion` in `micpal.iss`.
- Python runtime: `$PyVersion` / `$PyReleaseTag` in `build_windows.ps1`
  (see the python-build-standalone releases for current tags).
- Dependencies: `requirements.txt` in the repo root (pinned).
