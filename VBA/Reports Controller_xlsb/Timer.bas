Attribute VB_Name = "Timer"
Option Explicit

Public Const timerSeconds As Single = 60 ' seconds

' http://stackoverflow.com/questions/20269844/api-timers-in-vba-how-to-make-safe
' Declaration of API and Timer variable
#If VBA7 And Win64 Then    ' 64 bit Excel under 64-bit windows
                           ' Use LongLong and LongPtr

    Private Declare PtrSafe Function SetTimer Lib "user32" _
                                    (ByVal hwnd As LongPtr, _
                                     ByVal nIDEvent As LongPtr, _
                                     ByVal uElapse As LongLong, _
                                     ByVal lpTimerFunc As LongPtr _
                                     ) As LongLong

    Public Declare PtrSafe Function KillTimer Lib "user32" _
                                    (ByVal hwnd As LongPtr, _
                                     ByVal nIDEvent As LongPtr _
                                     ) As LongLong
    Public TimerID As LongPtr


#ElseIf VBA7 Then     ' 64 bit Excel in all environments
                      ' Use LongPtr only, LongLong is not available

    Private Declare PtrSafe Function SetTimer Lib "user32" _
                                    (ByVal hwnd As LongPtr, _
                                     ByVal nIDEvent As Long, _
                                     ByVal uElapse As Long, _
                                     ByVal lpTimerFunc As LongPtr) As LongPtr

    Private Declare PtrSafe Function KillTimer Lib "user32" _
                                    (ByVal hwnd As LongPtr, _
                                     ByVal nIDEvent As Long) As Long

    Public TimerID As LongPtr

#Else    ' 32 bit Excel

    Private Declare Function SetTimer Lib "user32" _
                            (ByVal hwnd As Long, _
                             ByVal nIDEvent As Long, _
                             ByVal uElapse As Long, _
                             ByVal lpTimerFunc As Long) As Long

    Public Declare Function KillTimer Lib "user32" _
                            (ByVal hwnd As Long, _
                             ByVal nIDEvent As Long) As Long

    Public TimerID As Long

#End If

#If VBA7 And Win64 Then     ' 64 bit Excel under 64-bit windows  ' Use LongLong and LongPtr
                            ' Note that wMsg is always the WM_TIMER message, which actually fits in a Long
    Public Sub TimerProc(ByVal hwnd As LongPtr, _
                         ByVal wMsg As LongLong, _
                         ByVal idEvent As LongPtr, _
                         ByVal dwTime As LongLong)
        On Error Resume Next
        Call Main.Check_And_Run
    
    'KillTimer hwnd, idEvent   ' Kill the recurring callback here, if that's what you want to do
                              ' Otherwise, implement a lobal KillTimer call on exit

    End Sub
#ElseIf VBA7 Then          ' 64 bit Excel in all environments
                           ' Use LongPtr only
    Public Sub TimerProc(ByVal hwnd As LongPtr, _
                         ByVal wMsg As Long, _
                         ByVal idEvent As LongPtr, _
                         ByVal dwTime As Long)
        On Error Resume Next
        Call Main.Check_And_Run
    
    End Sub

#Else    ' 32 bit Excel
    Public Sub TimerProc(ByVal hwnd As Long, _
                         ByVal wMsg As Long, _
                         ByVal idEvent As Long, _
                         ByVal dwTime As Long)
        On Error Resume Next
        Call Main.Check_And_Run
    End Sub
#End If

Sub Test()
    With ThisWorkbook.Sheets("ControlPanel").Shapes("StartStop Button")
        If .TextFrame2.TextRange.Characters.Text = "Start Processing" Then
            .TextFrame2.TextRange.Characters.Text = "Stop Processing"
            .Fill.ForeColor.RGB = RGB(209, 0, 36) ' Red
            
            Call Set_Global_Variables
            If Control_Table.DataBodyRange Is Nothing Then
                MsgBox "No reports for execution", vbExclamation + vbOKOnly, "Information"
                Exit Sub
            Else
                Call Main.Check_And_Run
                Call StartTimer
            End If
        Else
            ' stop processing
            Call EndTimer
            .TextFrame2.TextRange.Characters.Text = "Start Processing"
            .Fill.ForeColor.RGB = RGB(0, 176, 80) ' Green
        End If
    End With
End Sub

' Start / Kill Timer API
' ***********************************************************************************************
Sub StartTimer()
    On Error Resume Next
    TimerID = SetTimer(0&, 0&, timerSeconds * 1000&, AddressOf TimerProc)
    If Err.Number <> 0 Then
        Debug.Print Err.Number & ": " & Err.Description
    End If
End Sub

Sub EndTimer()
    On Error Resume Next
    KillTimer 0&, TimerID
    If Err.Number <> 0 Then
        Debug.Print Err.Number & ": " & Err.Description
    End If
End Sub
' ***********************************************************************************************
