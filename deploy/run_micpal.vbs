' run_micpal.vbs
' Silent launcher: starts the app with no console window.
' The Start Menu / Desktop shortcut points here. It runs the bundled
' pythonw.exe (windowless) against launch.py, which starts Streamlit and
' opens the browser. To stop the app, end "pythonw.exe" in Task Manager.

Set fso = CreateObject("Scripting.FileSystemObject")
Set sh  = CreateObject("WScript.Shell")

' Folder this .vbs lives in == the install dir ({app}).
appDir = fso.GetParentFolderName(WScript.ScriptFullName)

pythonw = appDir & "\python\pythonw.exe"
launcher = appDir & "\launch.py"

' 0 = hidden window, False = don't wait (server keeps running in background).
sh.Run """" & pythonw & """ """ & launcher & """", 0, False
