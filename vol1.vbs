T = Timer()  
Do while T + 276 > Timer() 
Call Volume("MAX")
WScript.Sleep 2000
loop
'***********************************************************************

Sub Volume(Param)
    set oShell = CreateObject("WScript.Shell") 
    Select Case Param 
    Case "MAX"
        oShell.SendKeys "{" & chr(175) & " 50}" ' volume maximum 100%
    End select
End Sub
