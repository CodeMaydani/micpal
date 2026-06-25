' stop_micpal.vbs
' Stops the running app. The launcher starts a windowless pythonw.exe running
' launch.py, which in turn spawns a "streamlit run" child. This finds both via
' their command lines and terminates them -- without touching any unrelated
' Python process on the machine.

Set wmi = GetObject("winmgmts:\\.\root\cimv2")
Set procs = wmi.ExecQuery( _
    "SELECT ProcessId, CommandLine FROM Win32_Process " & _
    "WHERE Name = 'pythonw.exe' OR Name = 'python.exe'")

killed = 0
For Each p In procs
    cmd = p.CommandLine
    If Not IsNull(cmd) Then
        cmd = LCase(cmd)
        If InStr(cmd, "launch.py") > 0 Or InStr(cmd, "streamlit run") > 0 Then
            p.Terminate
            killed = killed + 1
        End If
    End If
Next

If killed = 0 Then
    MsgBox "Michpal is not running.", vbInformation, "Stop Michpal"
End If
