Set oShell = CreateObject("Wscript.Shell")
Dim strArgs
strArgs = "cmd /c son1.bat"
oShell.Run strArgs, 0, false