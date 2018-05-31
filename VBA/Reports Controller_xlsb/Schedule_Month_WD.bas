Attribute VB_Name = "Schedule_Month_WD"
Option Explicit
Option Compare Text
'
'
' Module with functions which calculate next scheduled Working Day according to schudule strings provided in parameters
'     Month
'     Month Working Days
' Both parameters required for calculation.
' Calculation can't be performed when one of parameters is empty!
'
' PARAMETERS ==============================================================================================================================
' Month - string
    'Comma separated list of months (calendar months).
    ' Can contain ketwords.
    'If empty - all months are considered.
    'Integer values from 1 to 12 expected.
    'Affects only [Month Calendar Days] and [Month Working Days] schedule.
    'Samples:
    'ALL - keyword (same as empty cell)
    '1,12 - first and last month
    '3,6,9,12
    '1..12 - range, which same as all months
    '1..6,8,10,12
    '1..3,9..12
    ' Note: when parameter is empty it is not considered as schedule, even if parameters 'Month Working Days' and 'Only Working Days' are not empty.
'
' Month Working Days - string
'   Comma separated list of elements describing allowed indecies of working days.
'   String can contain intervals (..) and keywords (ALL, last)
    ' Key words:
    'ALL - each working day of month
    'Last - last working day
    'Samples:
    '1..3,10..13,last-3,last-2,last-1,last
    '1,2,last
    ' Index of last working day depends on month, so will be calculated accordingly.
    ' Note: when parameter is empty it is not considered as schedule, even if parameters 'Months' and 'Only Working Days' are not empty.
'
' FUNCTIONS ==============================================================================================================================
' PrepareMonthsString
'    Converts Month-string into comma separated list of months.
' GetClosestScheduledWorkingDayMonth
'    Function returns Closes Scheduled Working Day defined by Month-string & WorkingDayMonth-string.
'    This function resolved simple scenario
'
' SeekClosestScheduledWorkingDayMonth
'    This function contains main logic of searching. Read logic below.
'
' FindOffsetToNextScheduledMonth (StartingDate As Date, MonthsStringConverted As String) As Byte
'    Returns offset needed to reach next allowed by schedule month
'

' Main Function
Function GetClosestScheduledWorkingDayMonthDateTime( _
                        ScheduleString As String, _
                        MonthsString As String, _
                        country_code As String, _
                        ExecutionTime As Date, _
                        Optional RecurXMinutes As Double, _
                        Optional ToTime As Date, _
                        Optional NotEarlierThanFixed As Date, _
                        Optional Report_Row_ID As Long) As Date
' Function returns Closes Scheduled Working Day DateTime for defined Month-string & WorkingDayMonth-string
' Schedule String represents days of month
    ' Sample
    '       1, 3..10, last-5, last
    ' 3..10 is a range that means the same as string 3,4,5,6,7,8,9,10
    ' 'last' is a keyword, which means 'last working day', will be calculated accordingly to country
' MonthsString
'   Represents list of months, when action should occur
'   Sample
'   12 - in December
'   1..3 - from Jan to Mar
'   3,6,9,12 - Mar, Jun, Sep, Dec
' Country_Code - to find corresponding calendar
' Report_Row_ID - row of cell with report

    Dim sErrMessage As String
    Dim bTodayIsFine As Boolean
    Dim NotEarlierThan As Date
    Dim ToTimeLimit As Date
    Dim ExecTime As Date
    Dim MonthsStringConverted As String
    Dim tempDate As Date
    
    On Error GoTo ErrHandler
    
    If Control_Table Is Nothing Then
        Call Set_Global_Variables
    End If
    
    ' if one of parameters is empty - return error
    If (ScheduleString = vbNullString) Or (MonthsString = vbNullString) Then
        sErrMessage = "[Months] or [Month Working Days] is empty."
        GoTo ErrHandler
    End If
    
    If country_code = vbNullString Then
        sErrMessage = "[WD Country] is empty."
        GoTo ErrHandler
    End If
    
    With ThisWorkbook.Sheets("Calendar").ListObjects("Calendar")
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
            tempDate = SeekClosestScheduledWorkingDayMonth( _
                    ScheduleString, MonthsStringConverted, country_code, NotEarlierThan, Report_Row_ID)
            ' no guarantee that returned Today !
        Else
            ' find date that is Not Earlier than Tomorrow (Today+1) (or NotEarlierThan, if given)
            tempDate = SeekClosestScheduledWorkingDayMonth(ScheduleString, MonthsStringConverted, _
                country_code, _
                WorksheetFunction.Max(Date + 1, NotEarlierThan), _
                Report_Row_ID)
        End If
        ' found date
        
        If tempDate = Date Then
            ' if today - need to find closest time and compare with ToTime limit
            tempDate = GetClosestTime(tempDate + ExecTime, RecurXMinutes / 24 / 60, NotEarlierThan)    ' find time

            If tempDate > Date + ToTimeLimit Then
            ' if reached limit - search for next days
                tempDate = SeekClosestScheduledWorkingDayMonth( _
                                ScheduleString, MonthsStringConverted, _
                                country_code, _
                                WorksheetFunction.Max(Date + 1, NotEarlierThan), Report_Row_ID)
                ' left to add ExecutionTIme
                GetClosestScheduledWorkingDayMonthDateTime = DateValue(tempDate) + ExecTime
                Exit Function
            Else
            ' haven't reached limit yet
            ' no need to add ExecutionTime
                GetClosestScheduledWorkingDayMonthDateTime = tempDate
                Exit Function
            End If
        Else
        ' not today - but RecurXMinutes
        ' left to find ExecutionTime
            GetClosestScheduledWorkingDayMonthDateTime = GetClosestTime(tempDate + ExecTime, RecurXMinutes / 24 / 60, NotEarlierThan)    ' find time
            Exit Function
        End If
    Else
    ' if no schedule on minutes
        ' when no RecurXMinutes parameter - check if current time earlier then Today + Execution time
        If NotEarlierThan < Date + ExecTime Then
            tempDate = SeekClosestScheduledWorkingDayMonth( _
                            ScheduleString, _
                            MonthsStringConverted, _
                            country_code, _
                            WorksheetFunction.Max(Date, NotEarlierThan), _
                            Report_Row_ID)
        Else
            ' so search starting from next day
            tempDate = SeekClosestScheduledWorkingDayMonth( _
                            ScheduleString, MonthsStringConverted, _
                            country_code, _
                            WorksheetFunction.Max(Date + 1, NotEarlierThan), _
                            Report_Row_ID)
        End If
        
        GetClosestScheduledWorkingDayMonthDateTime = DateValue(tempDate) + ExecTime
        Exit Function
    
    End If ' If RecurXMinutes <> 0 Then
           
Exit_Function:
    
    Exit Function

ErrHandler:
    If Err.Number <> 0 Then
        Debug.Print Now, "GetClosestScheduledWorkingDayMonthDateTime", Err.Number & ": " & Err.Description _
            & IIf(Report_Row_ID <> 0, ". Row: " & Report_Row_ID & ", Report ID: " _
                & ControlPanel.Cells(Report_Row_ID, Control_Table.ListColumns("Report ID *").Range.Column).Value, "")
        Err.Clear
    Else
        Debug.Print Now, "GetClosestScheduledWorkingDayMonthDateTime", sErrMessage _
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
    
    GetClosestScheduledWorkingDayMonthDateTime = DateValue("9999-12-31")
    GoTo Exit_Function
    Resume ' for debugging
End Function

' Second Function
Private Function SeekClosestScheduledWorkingDayMonth( _
                ScheduleString As String, _
                MonthsStringConverted As String, _
                country_code As String, _
                Optional NotEarlierThanFixed As Date, _
                Optional Report_Row_ID As Long) As Date
' Function returns Closest Scheduled Working Day for defined Month-string & WorkingDayMonth-string
' may return Today
' Schedule String represents days of month
    ' Sample: 1, 3..10, last-5, last
' MonthsString
'   Represents list of month, when action should occur
' Country_Code - to find corresponding calendar

    Dim ScheduleStringConverted As String
    Dim i As Long
    Dim m As Long
    Dim arr
    Dim ClosestWDDayIndex As Integer
    Dim ddateWDIndex As Integer
    Dim max_date As Date
    Dim min_date As Date
    Dim sErrMessage As String
    Dim NotEarlierThan As Date
    
    On Error GoTo ErrHandler
    
    With ThisWorkbook.Sheets("Calendar").ListObjects("Calendar")
        max_date = WorksheetFunction.Max(.ListColumns("Date").DataBodyRange)
    End With
    
    If Date > max_date Then
        sErrMessage = "Calendar is obsolete. Update it with current year (next year recommended)."
        GoTo ErrHandler
    End If
    
    ' fix starting point
    ' so function may return Today
    NotEarlierThan = WorksheetFunction.Max(Date, DateValue(NotEarlierThanFixed))
    
    ' check if month of StartingDate is in the list of allowed months
    If InStr(1, "," & MonthsStringConverted & ",", "," & Month(NotEarlierThan) & ",", vbTextCompare) > 0 Then
        ' if StartingDate-month is in the list of scheduled months
        ' OK - continue to seek WD in this month
        'Stop
    Else
        ' get offset to next scheduled month
        ' and call recursively to find WD within that month
        
        SeekClosestScheduledWorkingDayMonth = SeekClosestScheduledWorkingDayMonth( _
            ScheduleString, _
            MonthsStringConverted, _
            country_code, _
            WorksheetFunction.EoMonth(NotEarlierThan, FindOffsetToNextScheduledMonth(NotEarlierThan, MonthsStringConverted) - 1) + 1, _
            Report_Row_ID)
        ' after certain call we will find suitable month
        ' otherwise we reach max_date in calendar and goto ErrHandler, return 9999-12-31
        Exit Function
    End If
    
    'if we here - month of 'NotEarlierThan' is in the list of allowed months
    ' get MonthWD-Index of NotEarlierThan
    ddateWDIndex = GetMWDbyDate(NotEarlierThan, country_code)
    
    ScheduleStringConverted = ScheduleStringToListOfDaysMWD(NotEarlierThan, ScheduleString, country_code)
    ' String was converted for month of NotEarlierThan
    ' "last"-keyword is month-specific, and it adds complexity...
    arr = Split(ScheduleStringConverted, ",", , vbTextCompare) ' Split always returns array, even for 1-element
    
    ClosestWDDayIndex = 999
    ' find min and closest elements
    ' as array is unsorted - go through all elements
    For i = LBound(arr) To UBound(arr)
        ' we seek for scheduled Working day within remaining working days in month (if exist)
        
        If NotEarlierThan <= GetDateByMWDnum(NotEarlierThan, CInt(arr(i)), country_code) Then
            ' overwrite ClosestWDDayIndex for each element of array when NotEarlierThan is less than resolved date
            If CInt(ClosestWDDayIndex) >= CInt(arr(i)) Then
                ClosestWDDayIndex = CInt(arr(i))
            End If
        End If
    Next i
    ' ClosestWDDayIndex now contains index of closest working day
    ' unless it wasn't found within month of StartingDate
    
    If ClosestWDDayIndex = 999 Then
        ' if couldn't find - Seek in next scheduled month
        SeekClosestScheduledWorkingDayMonth = SeekClosestScheduledWorkingDayMonth( _
            ScheduleString, _
            MonthsStringConverted, _
            country_code, _
            WorksheetFunction.EoMonth(NotEarlierThan, FindOffsetToNextScheduledMonth(NotEarlierThan, MonthsStringConverted) - 1) + 1, _
            Report_Row_ID)
    Else
    ' when suitable working day was found - resolve date and return it
        SeekClosestScheduledWorkingDayMonth = GetDateByMWDnum(NotEarlierThan, ClosestWDDayIndex, country_code)
    End If

Exit_Function:

    Exit Function
    
ErrHandler:
    If Err.Number <> 0 Then
        Debug.Print Now, "SeekClosestScheduledWorkingDayMonth", Err.Number & ": " & Err.Description
        Err.Clear
    Else
        Debug.Print Now, "SeekClosestScheduledWorkingDayMonth", sErrMessage
        
        If Report_Row_ID <> 0 Then
            If Not IsEditing Then
                If Application.CalculationState = xlDone Then
                    ControlPanel.Cells(Report_Row_ID, Control_Table.ListColumns("Schedule status").Range.Column).Value = _
                        sErrMessage
                End If
            End If
        End If
    End If
    
    SeekClosestScheduledWorkingDayMonth = DateValue("9999-12-31")
    GoTo Exit_Function
    Resume ' for debugging
End Function

' ================================================ SUPPORT FUNCTIONS ================================================
' ===================================================================================================================

Function PrepareMonthsString(MonthsString As String) As String
    Dim tmp_str As String
    Dim arrTmp
    Dim i As Byte
    
    tmp_str = Replace(MonthsString, " ", vbNullString) ' replace all spaces
    tmp_str = Replace(tmp_str, "ALL", "1,2,3,4,5,6,7,8,9,10,11,12", , , vbTextCompare)
    
    arrTmp = Split(tmp_str, ",", -1, vbTextCompare)
    tmp_str = vbNullString
    ' re-create tmp_str
    For i = LBound(arrTmp) To UBound(arrTmp)
        ' find intervals and convert them to lists
        If InStr(1, CStr(arrTmp(i)), "..", vbTextCompare) > 0 Then
            tmp_str = tmp_str & "," & NumberRangeToList(CStr(arrTmp(i)))
        Else
            tmp_str = tmp_str & "," & CStr(arrTmp(i))
        End If
    Next i ' arr element
    PrepareMonthsString = Mid(tmp_str, 2)
    
End Function

Function GetReallyLastWorkingDayOfMonth(ddate As Date, country_code As String) As Date
' starts search from last Calendar Day of the ddate-month
'
    Dim idate As Date
    Dim pos As Long
    
    On Error GoTo ErrHandler
    
    If ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("Date").DataBodyRange Is Nothing Then
        GetReallyLastWorkingDayOfMonth = 0
        Exit Function
    End If
    
    idate = WorksheetFunction.EoMonth(ddate, 0)
    With ThisWorkbook.Sheets("Calendar").ListObjects("Calendar")
        Do While Month(idate) = Month(ddate)
            pos = WorksheetFunction.Match(CDbl(idate), .ListColumns("Date").DataBodyRange, 0)
            If .ListColumns("WD " & country_code).DataBodyRange.Cells(pos, 1).Value = "Y" Then
                GetReallyLastWorkingDayOfMonth = idate
                Exit Function
            End If
            idate = idate - 1
        Loop
    End With
    
    ' if here - day wasn't found (all days are non-working)
    GetReallyLastWorkingDayOfMonth = 0
    Exit Function
    
ErrHandler:
    ' if here - something went wrong
    GetReallyLastWorkingDayOfMonth = 0
End Function

Function ScheduleStringToListOfDaysMWD(ddate As Date, ScheduleString As String, country_code As String) As String
' shortly: returns plain string of days for schedule for Month
' replace keywords: ALL, first, last
' convert ranges into lists of values

    Dim tmp_str As String
    
    On Error GoTo ErrHandler
    
    If ScheduleString = vbNullString Then Exit Function
    
    ' replace keywords ALL, first, remove spaces
    tmp_str = PrepareScheduleString(ScheduleString)
        
    If InStr(1, tmp_str, "last", vbTextCompare) > 0 Then
        tmp_str = ReplaceLastKeywordMWD(ddate, tmp_str, country_code)
    End If
         
    ' at this point, we have no "last"-keyword in Schedule String
    ' and have no spaces,
    ' however, might have ranges, and former "last-X", which have to be calculated
    ' e.g. 1..5,10,15,20-3,20-2,20-1,20
    ' or 1..5,10,15,20,20-3..20
    
    ' convert intervals and evaluate former "Last-X"
    ScheduleStringToListOfDaysMWD = OptimiseScheduleString(tmp_str)
    ' now sample string looks like string with comma separated integer values
    ' 1,2,3,4,5,10,15,17,18,19,20
    ' may have duplicates
    
    Exit Function

ErrHandler:
    ScheduleStringToListOfDaysMWD = vbNullString
End Function

Function ReplaceLastKeywordMWD(ddate As Date, ScheduleString As String, country_code As String) As String
' replaces "last"-keyword with its Working Day number
'
    Dim tmp_str As String
    Dim LastWdOfMonth As Date
    Dim IndexOfLastWorkingDayOfMonth As Integer
    Dim pos As Long
    
    On Error GoTo ErrHandler
    
    LastWdOfMonth = GetReallyLastWorkingDayOfMonth(ddate, country_code)   ' date
    
    With ThisWorkbook.Sheets("Calendar").ListObjects("Calendar")
        pos = WorksheetFunction.Match(CDbl(LastWdOfMonth), .ListColumns("Date").DataBodyRange, 0)
        
        IndexOfLastWorkingDayOfMonth = WorksheetFunction.Index( _
            .ListColumns("MWD " & country_code).DataBodyRange, _
            pos)
    End With
    
    ReplaceLastKeywordMWD = Replace(ScheduleString, "last", IndexOfLastWorkingDayOfMonth, , , vbTextCompare)
    Exit Function
    
ErrHandler:
    ReplaceLastKeywordMWD = vbNullString
End Function

Function GetMWDbyDate(ddate As Date, country_code As String) As Integer
' index of Working Day within month of ddate
    Dim pos As Long
    
    On Error GoTo ErrHandler
    
    With ThisWorkbook.Sheets("Calendar").ListObjects("Calendar")
    
        If .DataBodyRange Is Nothing Then
            GoTo ErrHandler
        End If
        
        pos = WorksheetFunction.Match(CDbl(ddate), .ListColumns("Date").DataBodyRange, 0)
        
        GetMWDbyDate = .ListColumns("MWD " & country_code).DataBodyRange.Cells(pos, 1).Value
        
    End With
        
    Exit Function
ErrHandler:
    GetMWDbyDate = -1
End Function

Function GetDateByMWDnum(ddate As Date, wd_num As Integer, country_code As String) As Date
' YWD num - index of working day in year
    Dim arrYear()
    Dim arrMonth()
    Dim arrMWD()
    Dim arrDates()
    Dim i As Long
    Dim dYear As Integer
    Dim dMonth As Byte
    
    On Error GoTo ErrHandler
    
    With ThisWorkbook.Sheets("Calendar").ListObjects("Calendar")
    
        If .DataBodyRange Is Nothing Then
            GoTo ErrHandler
        End If
        
        ' get values from sheet to array to speed up loop
        arrDates = .ListColumns("Date").DataBodyRange.Value
        arrYear = .ListColumns("Year").DataBodyRange.Value
        arrMonth = .ListColumns("Month").DataBodyRange.Value
        arrMWD = .ListColumns("MWD " & country_code).DataBodyRange.Value
        
    End With
    
    dMonth = Month(ddate)
    dYear = Year(ddate)
    
    For i = LBound(arrYear) To UBound(arrYear)
        If arrYear(i, 1) = dYear And arrMonth(i, 1) = dMonth And arrMWD(i, 1) = wd_num Then
            GetDateByMWDnum = CDate(arrDates(i, 1))
            Exit Function
        End If
    Next i
    
    Exit Function
ErrHandler:
    ' if here - something went wrong
    GetDateByMWDnum = 0
End Function


' ================================== OPTIONAL FUNCTIONS ==============
'  GetWDtoEOM(ddate As Date, wd_offset As String, country_code As String) As Date
'       returns data for WD-N for Month

Private Function GetFirstWorkingDayOfMonth(yr As Integer, mnth As Byte, country_code As String) As Date
' returns first working day of month
' function gets Year and Month as argument
' starts search from the first Calendar Day of the month
' goes forward till find working day or
' reach end of Calendar table

    Dim idate As Date
    Dim max_date As Date
    Dim pos As Long
    
    On Error GoTo ErrHandler
    
    With ThisWorkbook.Sheets("Calendar").ListObjects("Calendar")
        If .ListColumns("Date").DataBodyRange Is Nothing Then
            GetFirstWorkingDayOfMonth = 0
            Exit Function
        End If
        
        max_date = WorksheetFunction.Max(.ListColumns("Date").DataBodyRange)
        
        idate = CDate(yr & "-" & mnth & "-01")
        Do While idate <= max_date
            pos = WorksheetFunction.Match(CDbl(idate), .ListColumns("Date").DataBodyRange, 0)
            
            If .ListColumns("WD " & country_code).DataBodyRange.Cells(pos, 1).Value = "Y" Then
                GetFirstWorkingDayOfMonth = idate
                Exit Function
            End If
            idate = idate + 1
        Loop
    End With
    
ErrHandler:
    ' if here - something went wrong
    GetFirstWorkingDayOfMonth = 0
End Function

Private Function GetWDtoEOM(ddate As Date, wd_offset As String, country_code As String) As Date
' ddate can be any date
' usually it is a date of Next Run calculation
' can be last Calendar Day of the month.
' then function will calc result for next month.

' function Get Working Day to End Of Month (last WD in month)
' return Date that offset from Month End on specified number of days
' e.g. -3, => return 3rd working day to last WD of month
    Dim tmp_str As String
    Dim offset As Integer
    Dim i As Integer
    Dim idate As Date
    Dim max_date As Date
    
    On Error GoTo ErrHandler
    
    max_date = WorksheetFunction.Max(ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("Date").DataBodyRange)
    
    tmp_str = Replace(wd_offset, "last", vbNullString, , , vbTextCompare)
    
    offset = Val(Trim(tmp_str))
    ' e.g. -2
    ' seek for last working day of month of 'ddate' argument
    ' then 2 times search previous working day
    ' if it is below ddate - then same for next month
    ' keep in mind
    
    'idate = GetLastWorkingDayOfMonth(Year(ddate), Month(ddate), country_code)
    idate = GetReallyLastWorkingDayOfMonth(ddate, country_code)
    ' possibly can return 0 if end of calendar
    If CDbl(idate) <> 0 Then
        For i = -1 To offset Step -1
        ' if WD-1 - then loop will be skipped, because offset = -1
            idate = GetPreviousWorkingDay(idate, country_code)
            ' no need to check on max_date, as going back
        Next i
    End If
    
    GetWDtoEOM = idate
            
    If idate <= ddate And CDbl(idate) <> 0 Then
        ' call same function but for beginning of next month
        If WorksheetFunction.EoMonth(ddate, 0) + 1 < max_date Then
            GetWDtoEOM = GetWDtoEOM(WorksheetFunction.EoMonth(ddate, 0) + 1, wd_offset, country_code)
        Else
            ' if reached end of provided Calendar - error
            GoTo ErrHandler
        End If
    End If
    
    Exit Function
    
ErrHandler:
    GetWDtoEOM = 0
End Function

'Function GetLastWorkingDayOfMonth(yr As Integer, mnth As Byte, country_code As String) As Date
'' OBSOLETE FUNCTION - backward compatibility
'' function gets Year and Month as argument
'' starts search from last Calendar Day of the month
'' goes back till reach minimal date in calendar table
''
'    Dim idate As Date
'    Dim min_date As Date
'    Dim pos As Long
'
'    On Error GoTo ErrHandler
'
'    If ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("Date").DataBodyRange Is Nothing Then
'        GetLastWorkingDayOfMonth = 0
'        Exit Function
'    End If
'
'    min_date = WorksheetFunction.Min(ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("Date").DataBodyRange)
'    idate = WorksheetFunction.EoMonth(CDate(yr & "-" & mnth & "-01"), 0)
'    Do While idate >= min_date
'        pos = WorksheetFunction.Match(CDbl(idate), ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("Date").DataBodyRange, 0)
'        If ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("WD " & country_code).DataBodyRange.Cells(pos, 1).Value = "Y" Then
'            GetLastWorkingDayOfMonth = idate
'            Exit Function
'        End If
'        idate = idate - 1
'    Loop
'
'ErrHandler:
'    ' if here - something went wrong
'    GetLastWorkingDayOfMonth = 0
'End Function
