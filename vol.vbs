Call Volume("MAX")
'***********************************************************************

Sub Volume(Param)
    set oShell = CreateObject("WScript.Shell") 
    Select Case Param 
    Case "MAX"
        oShell.SendKeys "{" & chr(175) & " 50}" ' volume maximum 100%
    End select
End Sub
