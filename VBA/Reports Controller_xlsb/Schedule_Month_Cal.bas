Attribute VB_Name = "Schedule_Month_Cal"
Option Explicit
Option Compare Text

Function GetClosestScheduledMonthCalendarDateTime( _
                ScheduleString As String, _
                MonthsString As String, _
                ExecutionTime As Date, _
                Optional RecurXMinutes As Double, _
                Optional ToTime As Date, _
                Optional NotEarlierThanFixed As Date, _
                Optional Report_Row_ID As Long) As Date

' Logic
' TODO: describe new logic
' when offset reaches end of calendar - function returns 9999-12-31
    
    Dim tempDate As Date
    Dim bTodayIsFine As Boolean
    Dim NotEarlierThan As Date
    Dim ToTimeLimit As Date
    Dim sErrMessage As String
    Dim ExecTime As Date
    Dim MonthsStringConverted As String
    
    On Error GoTo ErrHandler
    
    If Control_Table Is Nothing Then
        Call Set_Global_Variables
    End If
            
    ' if one of parameters is empty - return error
    If (ScheduleString = vbNullString) Or (MonthsString = vbNullString) Then
        sErrMessage = "[Months] or [Month Calendar Days] is empty."
        GoTo ErrHandler
    End If
    
    With Calendar.ListObjects("Calendar")
        If .ListColumns("Date").DataBodyRange Is Nothing Then
            sErrMessage = "Calendar table is empty."
            GoTo ErrHandler
        End If
    End With
    
    NotEarlierThan = WorksheetFunction.Max(Now, NotEarlierThanFixed)
    
    MonthsStringConverted = PrepareMonthsString(MonthsString)
    
    ExecTime = TimeValue(Hour(ExecutionTime) & ":" & Minute(ExecutionTime) & ":" & Second(ExecutionTime))
    
    If RecurXMinutes <> 0 Then
    ' when scheduled with RecurXMinutes it may still be executed today
        
        ToTimeLimit = IIf(ToTime <> 0, _
                TimeValue(Hour(ToTime) & ":" & Minute(ToTime) & ":" & Second(ToTime)), _
                TimeValue("23:59:59"))
        
        bTodayIsFine = (Date + ToTimeLimit > NotEarlierThan)

        If bTodayIsFine Then
            ' find date
            tempDate = SeekClosestScheduledMonthCalendarDay( _
                    ScheduleString, MonthsStringConverted, NotEarlierThan, Report_Row_ID)
            ' no guarantee that returned Today !
        Else
            ' find date that is Not Earlier than Tomorrow (Today+1) (or NotEarlierThan, if given)
            tempDate = SeekClosestScheduledMonthCalendarDay(ScheduleString, MonthsStringConverted, _
                WorksheetFunction.Max(Date + 1, NotEarlierThan), _
                Report_Row_ID)
        End If
        ' found date
        
        If tempDate = Date Then
            ' if today - need to find closest time and compare with ToTime limit
            tempDate = GetClosestTime(tempDate + ExecTime, RecurXMinutes / 24 / 60, NotEarlierThan)    ' find time

            If tempDate > Date + ToTimeLimit Then
            ' if reached limit - search for next days
                tempDate = SeekClosestScheduledMonthCalendarDay( _
                        ScheduleString, MonthsStringConverted, WorksheetFunction.Max(Date + 1, NotEarlierThan), Report_Row_ID)
                ' left to add ExecutionTIme
                GetClosestScheduledMonthCalendarDateTime = DateValue(tempDate) + ExecTime
                Exit Function
            Else
            ' haven't reached limit yet
            ' no need to add ExecutionTime
                GetClosestScheduledMonthCalendarDateTime = tempDate
                Exit Function
            End If
        Else
        ' not today - but RecurXMinutes
        ' left to find ExecutionTime
            GetClosestScheduledMonthCalendarDateTime = GetClosestTime(tempDate + ExecTime, RecurXMinutes / 24 / 60, NotEarlierThan)    ' find time
            Exit Function
        End If
    Else
    ' if no schedule on minutes
        ' when no RecurXMinutes parameter - check if current time earlier then Today + Execution time
        
        If NotEarlierThan < Date + ExecTime Then
            tempDate = SeekClosestScheduledMonthCalendarDay( _
                            ScheduleString, _
                            MonthsStringConverted, _
                            WorksheetFunction.Max(Date, NotEarlierThan), _
                            Report_Row_ID)
        Else
        ' so search starting from next day
            tempDate = SeekClosestScheduledMonthCalendarDay( _
                            ScheduleString, _
                            MonthsStringConverted, _
                            WorksheetFunction.Max(Date + 1, NotEarlierThan), _
                            Report_Row_ID)
        End If
        
        GetClosestScheduledMonthCalendarDateTime = DateValue(tempDate) + ExecTime
        Exit Function
    
    End If ' If RecurXMinutes <> 0 Then
    
    
Exit_Function:

    Exit Function
    
ErrHandler:
    If Err.Number <> 0 Then
        Debug.Print Now, "GetClosestScheduledMonthCalendarDateTime", Err.Number & ": " & Err.Description _
            & IIf(Report_Row_ID <> 0, ". Row: " & Report_Row_ID & ", Report ID: " _
                & ControlPanel.Cells(Report_Row_ID, Control_Table.ListColumns("Report ID *").Range.Column).Value, "")
        Err.Clear
    Else
        Debug.Print Now, "GetClosestScheduledMonthCalendarDateTime", sErrMessage _
            & IIf(Report_Row_ID <> 0, " Row: " & Report_Row_ID & ", Report ID: " _
                & ControlPanel.Cells(Report_Row_ID, Control_Table.ListColumns("Report ID *").Range.Column).Value, "")
                
        If Report_Row_ID <> 0 Then
            If Not IsEditing Then
                If Application.CalculationState = xlDone Then
                    ControlPanel.Cells(Report_Row_ID, Control_Table.ListColumns("Schedule status").Range.Column).Value = _
                        sErrMessage
                End If
            End If
        End If
    End If
    
    GetClosestScheduledMonthCalendarDateTime = DateValue("9999-12-31")
    GoTo Exit_Function
    Resume ' for debugging
End Function

Function SeekClosestScheduledMonthCalendarDay( _
                ScheduleString As String, _
                MonthsStringConverted As String, _
                Optional NotEarlierThanFixed As Date, _
                Optional Report_Row_ID As Long) As Date
' Function returns Closest Scheduled Working Day for defined Month-string & [Month Calendar Days]-string
' may return Today
' Schedule String represents days of month
    ' Sample: 1, 3..10, last-5, last
' MonthsString
'   Represents list of month, when action should occur
' Next run can't be earlier than Now
' it means that if no 'NotEarlierThanFixed' provided function will seek for date
' that is not earlier than Today

    Dim ScheduleStringConverted As String
    Dim i As Long
    Dim arr
    Dim max_date As Date
    Dim sErrMessage As String
    Dim ClosestDayIndex As Integer
    Dim NotEarlierThan As Date
    
    On Error GoTo ErrHandler
    
    With Calendar.ListObjects("Calendar")
        max_date = WorksheetFunction.Max(.ListColumns("Date").DataBodyRange)
    End With
    
    If Date > max_date Then
        sErrMessage = "Calendar is obsolete. Update it with current year (next year recommended)."
        GoTo ErrHandler
    End If
    
    ' fix starting point
    ' so function may return Today
    NotEarlierThan = WorksheetFunction.Max(Date, DateValue(NotEarlierThanFixed))
    
    ' when 'NotEarlierThanFixed' provided, e.g. end of current month
    ' Next Run shouldn't happen before that date
    
    ' check if month of NotEarlierThan is in the list of allowed months
    If InStr(1, "," & MonthsStringConverted & ",", "," & Month(NotEarlierThan) & ",", vbTextCompare) > 0 Then
        ' if StartingDate-month is in the list of scheduled months
        ' OK - continue to seek WD in this month
        'Stop
    Else
        ' get offset to next scheduled month
        ' and call recursively to find WD within that month
        
        SeekClosestScheduledMonthCalendarDay = SeekClosestScheduledMonthCalendarDay( _
            ScheduleString, _
            MonthsStringConverted, _
            WorksheetFunction.EoMonth(NotEarlierThan, FindOffsetToNextScheduledMonth(NotEarlierThan, MonthsStringConverted) - 1) + 1, _
            Report_Row_ID)
        ' after a certain call we will find suitable month
        ' otherwise we reach max_date in calendar and goto ErrHandler, return 9999-12-31
        Exit Function
    End If
    
    'if we are here - month of NotEarlierThan is in the list of allowed months
    
    ' Convert string to get rid of keywords like "last", "all"
    ScheduleStringConverted = ScheduleStringToListOfMonthCalendarDays(NotEarlierThan, ScheduleString)
    
    ' String was converted for month of 'NotEarlierThan'
    ' "last"-keyword is month-specific
    arr = Split(ScheduleStringConverted, ",", , vbTextCompare) ' Split always returns array, even for 1-element
    
    ' now need to find min element which is greater than day( StartingDate )
    
    ClosestDayIndex = 999
    
    ' if not last day of month
    If Day(NotEarlierThan) < Day(WorksheetFunction.EoMonth(NotEarlierThan, 0)) Then
        ' find min and closest elements
        ' as array is unsorted - go through all elements
        For i = LBound(arr) To UBound(arr)
            ' consider "=" situation to support recursive calls with
            ' StartingDate = 1st day of next month when it is a working day - we have to take it
            ' therefore, initial call of this function has to be done from TOMORROW (desired date + 1)
            ' unless we want to return Today (when bTodayIsFine - for scenario with 'Recur Every X Minutes')
            ' TODO - consider 'Recur Every X Minutes'
            
            If Day(NotEarlierThan) <= CInt(arr(i)) Then
                ' overwrite ClosestDayIndex for each element of array when Day of StartingDate is less than day in schedule
                If ClosestDayIndex > CInt(arr(i)) Then
                    ClosestDayIndex = CInt(arr(i))
                End If
            End If
        Next i
        ' ClosestDayIndex now contains index of closest working day
        ' unless it wasn't found within month of StartingDate
    End If
    
    If ClosestDayIndex = 999 Then
        ' if couldn't find - Seek in next scheduled month
        SeekClosestScheduledMonthCalendarDay = SeekClosestScheduledMonthCalendarDay( _
            ScheduleString, _
            MonthsStringConverted, _
            WorksheetFunction.EoMonth(NotEarlierThan, FindOffsetToNextScheduledMonth(NotEarlierThan, MonthsStringConverted) - 1) + 1, _
            Report_Row_ID)
    Else
    ' when suitable day was found - resolve date and return it
        SeekClosestScheduledMonthCalendarDay = DateValue(Year(NotEarlierThan) & "-" _
                                                         & Month(NotEarlierThan) & "-" _
                                                         & ClosestDayIndex)
    End If

Exit_Function:

    Exit Function
    
ErrHandler:
    If Err.Number <> 0 Then
        Debug.Print Now, "SeekClosestScheduledMonthCalendarDay", Err.Number & ": " & Err.Description
        Err.Clear
    Else
        Debug.Print Now, "SeekClosestScheduledMonthCalendarDay", sErrMessage
        
        If Report_Row_ID <> 0 Then
            If Not IsEditing Then
                If Application.CalculationState = xlDone Then
                    ControlPanel.Cells(Report_Row_ID, Control_Table.ListColumns("Schedule status").Range.Column).Value = _
                        sErrMessage
                End If
            End If
        End If
    End If
    
    SeekClosestScheduledMonthCalendarDay = DateValue("9999-12-31")
    GoTo Exit_Function
    Resume ' for debugging
End Function

Function ScheduleStringToListOfMonthCalendarDays(StartingDate As Date, ScheduleString As String) As String
' shortly: returns plain string of days for schedule for Month
' replaces keywords: ALL, first, last
' convert ranges into lists of values

    Dim tmp_str As String
    
    On Error GoTo ErrHandler
    
    If ScheduleString = vbNullString Then Exit Function
    
    ' replace keywords ALL, first, remove spaces
    tmp_str = PrepareScheduleString(ScheduleString)
    
    ' replace last-keyword
    tmp_str = Replace(tmp_str, "last", _
                   Day(WorksheetFunction.EoMonth(StartingDate, 0)), , , vbTextCompare)
    
    ' at this point, we have no "last"-keyword in Schedule String
    ' and have no spaces,
    ' however, might have ranges, and former "last-X", which have to be calculated
    ' e.g. 1..5,10,15,20-3,20-2,20-1,20
    ' or 1..5,10,15,20,20-3..20
    
    ' convert intervals and evaluate former "Last-X"
    ScheduleStringToListOfMonthCalendarDays = OptimiseScheduleString(tmp_str)
    ' now sample string looks like string with comma separated integer values
    ' 1,2,3,4,5,10,15,17,18,19,20
    ' may have duplicates
    
    Exit Function

ErrHandler:
    ScheduleStringToListOfMonthCalendarDays = vbNullString
End Function
