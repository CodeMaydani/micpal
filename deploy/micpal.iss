; micpal.iss -- Inno Setup wizard for the מיכפל automation app (Windows only).
;
; Build steps:
;   1. Run deploy\build_windows.ps1 on Windows (stages deploy\build\).
;   2. Open this file in Inno Setup and Build (Ctrl+F9) -> deploy\Output\setup.exe
;
; Installs a self-contained app (bundled Python + deps) to Z:\Micpal and adds
; a Start Menu shortcut. The data folder is NOT asked for here -- the user sets
; it in the app sidebar on first run (it defaults to Z:\Msk8).

#define AppName "Michpal Template Automation"
#define AppVersion "1.0.0"
#define AppPublisher "Michpal Automation"

[Setup]
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
; Decided install location: on the share, user-writable (so config.json works).
DefaultDirName=Z:\Micpal
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
OutputDir=Output
OutputBaseFilename=micpal-setup-{#AppVersion}
Compression=lzma2
SolidCompression=yes
; Per-user install: no admin elevation needed, matches a user-writable target.
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

[Files]
; Bundled relocatable Python (with streamlit+openpyxl already installed).
Source: "build\python\*"; DestDir: "{app}\python"; Flags: recursesubdirs createallsubdirs ignoreversion
; Application files (engine.py, app.py, launch.py, ...).
Source: "build\app\*"; DestDir: "{app}"; Flags: recursesubdirs createallsubdirs ignoreversion
; Silent launcher + stopper the shortcuts point at.
Source: "run_micpal.vbs"; DestDir: "{app}"; Flags: ignoreversion
Source: "stop_micpal.vbs"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
; Start Menu + optional Desktop shortcut -> the silent VBS launcher.
Name: "{group}\{#AppName}"; Filename: "{app}\run_micpal.vbs"; WorkingDir: "{app}"
Name: "{group}\Stop {#AppName}"; Filename: "{app}\stop_micpal.vbs"; WorkingDir: "{app}"
Name: "{group}\Uninstall {#AppName}"; Filename: "{uninstallexe}"
Name: "{userdesktop}\{#AppName}"; Filename: "{app}\run_micpal.vbs"; WorkingDir: "{app}"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional shortcuts:"

[Run]
; Offer to launch right after install.
Filename: "{app}\run_micpal.vbs"; Description: "Launch {#AppName} now"; Flags: postinstall nowait skipifsilent shellexec

[UninstallDelete]
; Remove the per-machine config written next to the app, but never touch the
; data share (Z:\Msk8) or generated templates.
Type: files; Name: "{app}\config.json"
Type: filesandordirs; Name: "{app}\__pycache__"
