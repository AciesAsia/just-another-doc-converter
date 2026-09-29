Set oShell = CreateObject("WScript.Shell")
sScript = Replace(WScript.ScriptFullName, WScript.ScriptName, "") & "gdrive_sync_gui.py"
oShell.Run "pythonw """ & sScript & """", 0, False
