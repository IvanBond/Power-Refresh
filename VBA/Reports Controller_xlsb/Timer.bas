Attribute VB_Name = "Timer"
Option Explicit

' Attribute VB_Name = "timerProc"

#If VBA7 Or VBA8 Then
'  Code is running in the new VBA7 editor
    
    Public Declare PtrSafe Function SetTimer Lib "user32" ( _
        ByVal HWnd As Long, _
        ByVal nIDEvent As Long, _
        ByVal uElapse As Long, _
        ByVal lpTimerFunc As LongPtr) As Long

    Public Declare PtrSafe Function KillTimer Lib "user32" ( _
        ByVal HWnd As Long, _
        ByVal nIDEvent As Long) As Long
#Else

' Code is running in VBA version 6 or earlier
    Public Declare Function SetTimer Lib "user32" ( _
        ByVal HWnd As Long, _
        ByVal nIDEvent As Long, _
        ByVal uElapse As Long, _
        ByVal lpTimerFunc As Long) As Long

    Public Declare Function KillTimer Lib "user32" ( _
        ByVal HWnd As Long, _
        ByVal nIDEvent As Long) As Long

#End If

Public TimerID As Long
Public Const timerSeconds As Single = 60

Sub Test()
    With ThisWorkbook.Sheets("Control Panel").Shapes("StartStop Button")
        If .TextFrame2.TextRange.Characters.Text = "Start Processing" Then
            .TextFrame2.TextRange.Characters.Text = "Stop Processing"
            .Fill.ForeColor.RGB = RGB(209, 0, 36) ' Red
            
            Call Main.Check_And_Run
            Call StartTimer
        Else
            ' stop processing
            Call EndTimer
            .TextFrame2.TextRange.Characters.Text = "Start Processing"
            .Fill.ForeColor.RGB = RGB(0, 176, 80) ' Green
        End If
    End With
End Sub

Sub StartTimer()
    ' How often to "pop" the timer.
    TimerID = SetTimer(0&, 0&, timerSeconds * 1000&, AddressOf timerProc)
    'Call Build_Processing_Bar
End Sub

Sub EndTimer()
    On Error Resume Next
    KillTimer 0&, TimerID
    'deleteToolBar (tBarProgress)
End Sub

Sub timerProc(ByVal HWnd As Long, ByVal uMsg As Long, ByVal nIDEvent As Long, ByVal dwTimer As Long)
    ' Schedule the jobs the timer should do
    Call Main.Check_And_Run
End Sub
