code = MsgBox("3" , vbCritical )
code = MsgBox("2" , vbCritical )
code = MsgBox("1" , vbCritical )

Dim Sh
Set Sh = CreateObject("WScript.Shell")
Sh.Run "4.vbs"
Set Sh = Nothing