Attribute VB_Name = "Support_Functions"
Option Explicit

' http://www.fmsinc.com/microsoftaccess/modules/examples/avoiddoevents.asp
' http://analystcave.com/vba-sleep-vs-wait/
' http://www.exceltrick.com/formulas_macros/vba-wait-and-sleep-functions/
Public Sub WaitSeconds(intSeconds As Integer)
    Dim datTime As Date

    datTime = DateAdd("s", intSeconds, Now)

    Do
        ' Yield to other programs (better than using DoEvents which eats up all the CPU cycles)
        Sleep 100
        DoEvents
    Loop Until Now >= datTime
End Sub
