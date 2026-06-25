<#
build_windows.ps1
Stage everything the Inno Setup wizard needs, on a Windows machine.

Produces deploy\build\ containing:
    python\   -- a relocatable CPython with streamlit+openpyxl installed
    app\      -- the application files (engine.py, app.py, ... , launch.py)

After this runs, compile deploy\micpal.iss with Inno Setup to get setup.exe.

WHY this must run on Windows:
    pip installs platform-specific wheels (streamlit pulls in pyarrow etc.).
    Those must be the Windows wheels, so the build host has to be Windows.
    The Linux dev box cannot produce this bundle reliably.

Usage (from the repo root, in PowerShell):
    ./deploy/build_windows.ps1
#>

$ErrorActionPreference = "Stop"

# --- versions / sources (bump as needed) ----------------------------------
# Relocatable CPython from astral-sh/python-build-standalone. The
# "install_only" asset extracts to a clean python\ folder that runs anywhere.
$PyVersion   = "3.13.1"
$PyReleaseTag = "20241206"   # the python-build-standalone release date tag
$PyAsset = "cpython-$PyVersion+$PyReleaseTag-x86_64-pc-windows-msvc-install_only.tar.gz"
$PyUrl   = "https://github.com/astral-sh/python-build-standalone/releases/download/$PyReleaseTag/$PyAsset"

# --- paths ----------------------------------------------------------------
$RepoRoot  = Split-Path -Parent $PSScriptRoot
$DeployDir = $PSScriptRoot
$BuildDir  = Join-Path $DeployDir "build"
$PyDir     = Join-Path $BuildDir "python"
$AppDir    = Join-Path $BuildDir "app"
$Tarball   = Join-Path $BuildDir $PyAsset

# App files that ship to the user (mirror the deployment set, exclude dev/archive).
$AppFiles = @(
    "engine.py",
    "app.py",
    "gen_company.py",
    "config.py",
    "launch.py",
    "ui_check.py",
    "requirements.txt",
    "README.md"
)

# --- clean & recreate build dir -------------------------------------------
Write-Host "==> Cleaning $BuildDir"
if (Test-Path $BuildDir) { Remove-Item -Recurse -Force $BuildDir }
New-Item -ItemType Directory -Force -Path $BuildDir | Out-Null
New-Item -ItemType Directory -Force -Path $AppDir   | Out-Null

# --- fetch + extract standalone Python ------------------------------------
Write-Host "==> Downloading $PyAsset"
Invoke-WebRequest -Uri $PyUrl -OutFile $Tarball

Write-Host "==> Extracting Python (tar is built into Windows 10+)"
tar -xzf $Tarball -C $BuildDir   # extracts a 'python' folder
Remove-Item $Tarball

$PyExe = Join-Path $PyDir "python.exe"
if (-not (Test-Path $PyExe)) { throw "python.exe not found at $PyExe after extract" }

# --- install dependencies into that Python --------------------------------
Write-Host "==> Installing dependencies"
& $PyExe -m pip install --upgrade pip
& $PyExe -m pip install -r (Join-Path $RepoRoot "requirements.txt")

# --- stage app files ------------------------------------------------------
Write-Host "==> Staging app files"
foreach ($f in $AppFiles) {
    Copy-Item (Join-Path $RepoRoot $f) (Join-Path $AppDir $f)
}

Write-Host ""
Write-Host "==> Build staged at $BuildDir"
Write-Host "    Next: open deploy\micpal.iss in Inno Setup and Build (Ctrl+F9)."
