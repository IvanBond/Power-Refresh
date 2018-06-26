Attribute VB_Name = "Schedule"
Option Explicit
Option Compare Text

' Acronyms
' EOM - end of month
' EOY - end of year
' last - key word for 'last Working Day'
' ALL - key word for all working days, can be used for Month or Year
'

' Schedule: string with pattern (sample)
'       1,3..10,last-5, last
' 3..10 is a range will be converted into string 3,4,5,6,7,8,9,10
' 'last' is a keyword, which means 'last working day', will be calculated
' accordingly to country / time frame (can be last WD in a week / month / year)

Sub ReCalcNextRunForAllReports()
    Dim cell As Range
    
    If vbYes = MsgBox("Are you sure you want to re-calculate Next Run DateTime for all reports?", _
                        vbYesNoCancel + vbDefaultButton3 + vbQuestion, "Question") Then
        ' let remaining code to be executed
    Else
        Exit Sub
    End If
    
    On Error GoTo ErrHandler
    
    Application.ScreenUpdating = False
    If Control_Table Is Nothing Then
        Call Set_Global_Variables
    End If
    
    For Each cell In Control_Table.ListColumns("Report ID *").DataBodyRange
        Call SetReportParameter(cell.Row, "Next Run", GetScheduledRunTime(cell.Row))
    Next cell

Exit_sub:
    On Error Resume Next
    Application.ScreenUpdating = True
    Exit Sub
    
ErrHandler:
    Debug.Print Now, "ReCalcNextRunForAllReports", Err.Number & ": " & Err.Description
    Err.Clear
    GoTo Exit_sub
    Resume
End Sub

Sub ttt()
    Call SetReportParameter(15, "Next Run", GetScheduledRunTime(15))
End Sub
    

Function GetScheduledRunTime(Report_Row_ID As Long) As Date
    Dim ResultingDateTime As Date
    
    If Control_Table Is Nothing Then
        Call Set_Global_Variables
    End If
    
    If (GetReportParameter(Report_Row_ID, "Only Working Days") = "Y" Or _
        GetReportParameter(Report_Row_ID, "Month Working Days") <> vbNullString) And _
            GetReportParameter(Report_Row_ID, "WD Country") = vbNullString Then
        Call SetReportParameter(Report_Row_ID, "Schedule status", "'WD Country' is empty")
    End If
    
    ' Standard basic schedule - every X days
    If Control_Table.Parent.Cells(Report_Row_ID, _
        Control_Table.ListColumns("Recur every X days").Range.Column).Value <> vbNullString Then
        
        ResultingDateTime = GetClosestDateTimeRecurXDays( _
                                GetReportParameter(Report_Row_ID, "Start Date"), _
                                GetReportParameter(Report_Row_ID, "Recur every X days"), _
                                GetReportParameter(Report_Row_ID, "Execution Time"), _
                                GetReportParameter(Report_Row_ID, "Recur every X Minutes"), _
                                GetReportParameter(Report_Row_ID, "To Time"), _
                                Now(), _
                                Report_Row_ID)
    End If
    
    ' GetScheduledRunTime = 0 at this point
    If ResultingDateTime <> 0 Then
        GetScheduledRunTime = ResultingDateTime
    End If
    
    ' schedule by X Working days
    If GetReportParameter(Report_Row_ID, "Recur every X days") <> vbNullString And _
        GetReportParameter(Report_Row_ID, "WD Country") <> vbNullString And _
        GetReportParameter(Report_Row_ID, "Only Working Days") = "Y" Then
        
        ResultingDateTime = GetClosestDateTimeRecurXWorkingDays( _
                                GetReportParameter(Report_Row_ID, "Start Date"), _
                                GetReportParameter(Report_Row_ID, "Recur every X days"), _
                                GetReportParameter(Report_Row_ID, "WD Country"), _
                                GetReportParameter(Report_Row_ID, "Execution Time"), _
                                GetReportParameter(Report_Row_ID, "Recur every X Minutes"), _
                                GetReportParameter(Report_Row_ID, "To Time"), _
                                Now(), _
                                Report_Row_ID)
    End If
    
    ' take earliest time
    If ResultingDateTime <> 0 Then
        If GetScheduledRunTime = 0 Then
            GetScheduledRunTime = ResultingDateTime
        Else
            If GetScheduledRunTime > ResultingDateTime Then
                GetScheduledRunTime = ResultingDateTime
            End If
        End If
    End If
    
    'Schedule by Month Days
    If GetReportParameter(Report_Row_ID, "Months") <> vbNullString And _
        GetReportParameter(Report_Row_ID, "Month Calendar Days") <> vbNullString Then
        
        ResultingDateTime = GetClosestScheduledMonthCalendarDateTime( _
                                GetReportParameter(Report_Row_ID, "Month Calendar Days"), _
                                GetReportParameter(Report_Row_ID, "Months"), _
                                GetReportParameter(Report_Row_ID, "Execution Time"), _
                                GetReportParameter(Report_Row_ID, "Recur every X Minutes"), _
                                GetReportParameter(Report_Row_ID, "To Time"), _
                                Now(), _
                                Report_Row_ID)
    End If
    
    ' take earliest time
    If ResultingDateTime <> 0 Then
        If GetScheduledRunTime = 0 Then
            GetScheduledRunTime = ResultingDateTime
        Else
            If GetScheduledRunTime > ResultingDateTime Then
                GetScheduledRunTime = ResultingDateTime
            End If
        End If
    End If
    
    ' Schedule by Month Working days
    If GetReportParameter(Report_Row_ID, "Months") <> vbNullString And _
        GetReportParameter(Report_Row_ID, "Month Working Days") <> vbNullString And _
        GetReportParameter(Report_Row_ID, "WD Country") <> vbNullString Then
        
        ResultingDateTime = GetClosestScheduledWorkingDayMonthDateTime( _
                                GetReportParameter(Report_Row_ID, "Month Working Days"), _
                                GetReportParameter(Report_Row_ID, "Months"), _
                                GetReportParameter(Report_Row_ID, "WD Country"), _
                                GetReportParameter(Report_Row_ID, "Execution Time"), _
                                GetReportParameter(Report_Row_ID, "Recur every X Minutes"), _
                                GetReportParameter(Report_Row_ID, "To Time"), _
                                Now(), _
                                Report_Row_ID)
    End If
    
    ' take earliest time
    If ResultingDateTime <> 0 Then
        If GetScheduledRunTime = 0 Then
            GetScheduledRunTime = ResultingDateTime
        Else
            If GetScheduledRunTime > ResultingDateTime Then
                GetScheduledRunTime = ResultingDateTime
            End If
        End If
    End If
    
    ' if all functions returned 0 -
    If GetScheduledRunTime = 0 Then
        GetScheduledRunTime = CDate("9999-12-31")
    End If
    
End Function
