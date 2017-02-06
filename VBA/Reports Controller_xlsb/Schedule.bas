Attribute VB_Name = "Schedule"
Option Explicit

Function GetNextWorkingDay(ddate As Date, country_code As String) As Date
    Dim iDate As Date
    Dim max_date As Date
    Dim pos As Long
    
    If ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("Date").DataBodyRange Is Nothing Then
        GetNextWorkingDay = 0
        Exit Function
    End If
    
    max_date = WorksheetFunction.Max(ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("Date").DataBodyRange)
    iDate = ddate + 1
    Do While iDate <= max_date
        pos = WorksheetFunction.Match(CDbl(iDate), ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("Date").DataBodyRange, 0)
        If WorksheetFunction.Index(ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("WD " & country_code).DataBodyRange, _
            pos) = "Y" Then
            GetNextWorkingDay = iDate
            Exit Function
        End If
        
        iDate = iDate + 1
    Loop

ErrHandler:
    ' if here - something went wrong
    GetNextWorkingDay = 0
End Function

Function GetPreviousWorkingDay(ddate As Date, country_code As String) As Date
    Dim iDate As Date
    Dim min_date As Date
    Dim pos As Long
    
    If ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("Date").DataBodyRange Is Nothing Then
        GetPreviousWorkingDay = 0
        Exit Function
    End If
    
    min_date = WorksheetFunction.Min(ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("Date").DataBodyRange)
    iDate = ddate - 1
    Do While iDate >= min_date
        pos = WorksheetFunction.Match(CDbl(iDate), ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("Date").DataBodyRange, 0)
        If WorksheetFunction.Index(ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("WD " & country_code).DataBodyRange, _
            pos) = "Y" Then
            GetPreviousWorkingDay = iDate
            Exit Function
        End If
        
        iDate = iDate - 1
    Loop

ErrHandler:
    ' if here - something went wrong
    GetPreviousWorkingDay = 0
End Function

' to get first working day of current week - just call with TODAY()-7
Function GetNextWeekFirstWorkingDay(ddate As Date, country_code As String)
    ' if entire week is a holiday - check next week after
    Dim iDate As Date
    Dim max_date As Date
    Dim pos As Long
    Dim NextWeekStart As Date
    Dim CurrentWeekNum As Byte
    
    If ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("Date").DataBodyRange Is Nothing Then
        GetNextWeekFirstWorkingDay = 0
        Exit Function
    End If
    
    max_date = WorksheetFunction.Max(ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("Date").DataBodyRange)
    
    ' position of Today
    pos = WorksheetFunction.Match(CDbl(ddate), ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("Date").DataBodyRange, 0)
    CurrentWeekNum = WorksheetFunction.Index(ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns( _
                        "WeekNum " & country_code).DataBodyRange, pos)
                        
    iDate = ddate + 1
    Do While iDate <= max_date
        ' next year can start - thus search for change of week num
        pos = WorksheetFunction.Match(CDbl(iDate), ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("Date").DataBodyRange, 0)
        If WorksheetFunction.Index(ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns( _
                "WeekNum " & country_code).DataBodyRange, pos) <> CurrentWeekNum Then
              ' found start of next week
              ' check if it is a working day
            If WorksheetFunction.Index(ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns( _
                    "WD " & country_code).DataBodyRange, pos) = "Y" Then
                GetNextWeekFirstWorkingDay = iDate
                Exit Function
            End If
        End If
        iDate = iDate + 1
    Loop

ErrHandler:
    ' if here - something went wrong
    GetNextWeekFirstWorkingDay = 0
End Function

' to get last working day of current week - just call with TODAY()-7
' or ( [desired date] - 7 )
Function GetNextWeekLastWorkingDay(ddate As Date, country_code As String)
    ' if entire week is a holiday - check next week after
    Dim iDate As Date
    Dim max_date As Date
    Dim pos As Long
    Dim NextWeekStart As Date
    Dim CurrentWeekNum As Byte
    Dim TargetWeekNum As Byte
    
    If ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("Date").DataBodyRange Is Nothing Then
        GetNextWeekLastWorkingDay = 0
        Exit Function
    End If
    
    max_date = WorksheetFunction.Max(ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("Date").DataBodyRange)
    
    ' position of Today
    pos = WorksheetFunction.Match(CDbl(ddate), ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("Date").DataBodyRange, 0)
    CurrentWeekNum = WorksheetFunction.Index(ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns( _
                        "WeekNum " & country_code).DataBodyRange, pos)
                        
    iDate = ddate + 1
    Do While iDate <= max_date
        ' next year can start - thus search for change of week num
        pos = WorksheetFunction.Match(CDbl(iDate), ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("Date").DataBodyRange, 0)
        If WorksheetFunction.Index(ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns( _
                "WeekNum " & country_code).DataBodyRange, pos) <> CurrentWeekNum Then
            
            ' found start of next week
            ' remember its ID
            If TargetWeekNum = 0 Then
                ' do this only when TargetWeekNum has no value yet
                TargetWeekNum = WorksheetFunction.Index(ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns( _
                    "WeekNum " & country_code).DataBodyRange, pos)
            Else
                ' check if program reached next week after TargetWeekNum
                If TargetWeekNum <> WorksheetFunction.Index(ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns( _
                    "WeekNum " & country_code).DataBodyRange, pos) Then
                    ' if something has been found before - take it, it is last working day of Target Week
                    If GetNextWeekLastWorkingDay <> 0 Then
                        Exit Function
                    Else
                        ' probably entire week had no working days
                        ' in such case - re-define TargetWeekNum
                        TargetWeekNum = WorksheetFunction.Index(ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns( _
                                            "WeekNum " & country_code).DataBodyRange, pos)
                    End If
                End If
            End If
            
            ' check if it is a working day
            If WorksheetFunction.Index(ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns( _
                    "WD " & country_code).DataBodyRange, pos) = "Y" Then
                GetNextWeekLastWorkingDay = iDate
            End If
        End If
        iDate = iDate + 1
    Loop

ErrHandler:
    ' if here - something went wrong
    GetNextWeekLastWorkingDay = 0
End Function

Function GetLastWorkingDayOfMonth(yr As Integer, mnth As Byte, country_code As String)
' TODO - function definition
    Dim iDate As Date
    Dim min_date As Date
    Dim pos As Long
    
    If ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("Date").DataBodyRange Is Nothing Then
        GetLastWorkingDayOfMonth = 0
        Exit Function
    End If
    
    min_date = WorksheetFunction.Min(ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("Date").DataBodyRange)
    iDate = WorksheetFunction.EoMonth(CDate(yr & "-" & mnth & "-01"), 0)
    Do While iDate >= min_date
        pos = WorksheetFunction.Match(CDbl(iDate), ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("Date").DataBodyRange, 0)
        If WorksheetFunction.Index(ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("WD " & country_code).DataBodyRange, _
            pos) = "Y" Then
            GetLastWorkingDayOfMonth = iDate
            Exit Function
        End If
        iDate = iDate - 1
    Loop

ErrHandler:
    ' if here - something went wrong
    GetLastWorkingDayOfMonth = 0
End Function

Function GetLastWorkingDayOfYear(yr As Integer, country_code As String)
' TODO - function definition
    Dim iDate As Date
    Dim min_date As Date
    Dim pos As Long
    
    If ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("Date").DataBodyRange Is Nothing Then
        GetLastWorkingDayOfYear = 0
        Exit Function
    End If
    
    min_date = WorksheetFunction.Min(ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("Date").DataBodyRange)
    
    iDate = CDate(yr & "-12-31")
    Do While iDate >= min_date
        pos = WorksheetFunction.Match(CDbl(iDate), ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("Date").DataBodyRange, 0)
        If WorksheetFunction.Index(ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("WD " & country_code).DataBodyRange, _
            pos) = "Y" Then
            GetLastWorkingDayOfYear = iDate
            Exit Function
        End If
        iDate = iDate - 1
    Loop

ErrHandler:
    ' if here - something went wrong
    GetLastWorkingDayOfYear = 0
End Function

Function GetFirstWorkingDayOfMonth(yr As Integer, mnth As Byte, country_code As String)
' TODO - function definition
    Dim iDate As Date
    Dim max_date As Date
    Dim pos As Long
    
    On Error GoTo ErrHandler
    
    If ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("Date").DataBodyRange Is Nothing Then
        GetFirstWorkingDayOfMonth = 0
        Exit Function
    End If
    
    max_date = WorksheetFunction.Max(ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("Date").DataBodyRange)
    
    iDate = WorksheetFunction.EoMonth(CDate(yr & "-" & mnth & "-01"), -1)
    Do While iDate <= max_date
        pos = WorksheetFunction.Match(CDbl(iDate), ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("Date").DataBodyRange, 0)
        If WorksheetFunction.Index(ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("WD " & country_code).DataBodyRange, _
            pos) = "Y" Then
            GetFirstWorkingDayOfMonth = iDate
            Exit Function
        End If
        iDate = iDate + 1
    Loop

ErrHandler:
    ' if here - something went wrong
    GetFirstWorkingDayOfMonth = 0
End Function

Function GetFirstWorkingDayOfYear(yr As Integer, country_code As String)
' TODO - function definition
    Dim iDate As Date
    Dim max_date As Date
    Dim pos As Long
    
    If ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("Date").DataBodyRange Is Nothing Then
        GetFirstWorkingDayOfYear = 0
        Exit Function
    End If
    
    max_date = WorksheetFunction.Max(ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("Date").DataBodyRange)
    
    iDate = CDate(yr & "-01-01")
    Do While iDate <= max_date
        pos = WorksheetFunction.Match(CDbl(iDate), ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("Date").DataBodyRange, 0)
        If WorksheetFunction.Index(ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("WD " & country_code).DataBodyRange, _
            pos) = "Y" Then
            GetFirstWorkingDayOfYear = iDate
            Exit Function
        End If
        iDate = iDate + 1
    Loop

ErrHandler:
    ' if here - something went wrong
    GetFirstWorkingDayOfYear = 0
End Function

Function GetClosestScheduledWorkingDayYear(ddate As Date, Schedule As String, country_code As String)
' TODO - function definition
    Dim dDateWdNum As Integer
    Dim arrSchedule
    Dim LastWdOfCurrentYear As Date
    Dim NumOrLastWdOfCurrentYear As Integer
    Dim i As Integer
    Dim tmp_str As String
    Dim pos As Long
    Dim min_scheduled_day As Integer
    Dim min_next_day As Integer
    Dim arrTmp
    Dim wd_to_yec As Date
    
    If Schedule = vbNullString Then Exit Function
    
    On Error GoTo ErrHandler
    ' get row of requested date
    pos = WorksheetFunction.Match(CDbl(ddate), ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("Date").DataBodyRange, 0)
    
    ' check if it is a working day
    dDateWdNum = WorksheetFunction.Index(ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("YWD " & country_code).DataBodyRange, pos)
    If dDateWdNum = 0 Then
        ' working day number of requested date
        ' if it is holiday - then number of last working day before it - see table Calendar
        ' GetNextWorkingDay(dDate As Date, country_code As String) as date
        pos = WorksheetFunction.Match(CDbl(GetNextWorkingDay(ddate, country_code)), _
            ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("Date").DataBodyRange, 0)
        dDateWdNum = WorksheetFunction.Index(ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("YWD " & country_code).DataBodyRange, pos)
    End If
    
    ' parse Schedule
    tmp_str = Replace(Schedule, " ", vbNullString) ' replace all spaces
    tmp_str = Replace(tmp_str, "first", "1", , , vbTextCompare)
    
    ' replace keyword 'last' with its number
    If InStr(1, tmp_str, "last", vbTextCompare) > 0 Then
        
        LastWdOfCurrentYear = GetLastWorkingDayOfYear(Year(ddate), country_code)
        If ddate < LastWdOfCurrentYear Then
            NumOrLastWdOfCurrentYear = WorksheetFunction.Index(ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("YWD " & country_code).DataBodyRange, _
                WorksheetFunction.Match(CDbl(LastWdOfCurrentYear), ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("Date").DataBodyRange, 0))
            tmp_str = Replace(tmp_str, "last", NumOrLastWdOfCurrentYear, , , vbTextCompare)
        Else
            ' last WD of next year
            tmp_str = Replace(tmp_str, "last", _
                WorksheetFunction.Index(ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("YWD " & country_code).DataBodyRange, _
                    WorksheetFunction.Match(CDbl(GetLastWorkingDayOfYear(Year(ddate) + 1, country_code)), ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("Date").DataBodyRange, 0)), _
                    , , vbTextCompare)
        End If
    End If
    
    ' convert all ranges to array elements
    ' check if it contains ranges, e.g. 4-7
    ' can be following values WD-2, WD-1, - which are not ranges
    If InStr(1, Replace(tmp_str, "WD-", "#", , , vbTextCompare), "-", vbTextCompare) > 0 Then
        ' num range(s) provided
        arrTmp = Split(tmp_str, ",", -1, vbTextCompare)
        tmp_str = vbNullString
        
        ' re-create tmp_str
        For i = LBound(arrTmp) To UBound(arrTmp)
            ' if it is not "WD*" and contains "-"
            If UCase(Left(arrTmp(i), 2)) <> "WD" And InStr(1, arrTmp(i), "-", vbTextCompare) > 0 Then
                tmp_str = tmp_str & "," & NumberRangeToList(CStr(arrTmp(i)))
            Else
                tmp_str = tmp_str & "," & arrTmp(i)
            End If
        Next i
        tmp_str = Mid(tmp_str, 2)
    End If
    
    ' prepared schedule to new array
    arrSchedule = Split(tmp_str, ",", -1, vbTextCompare)
    
    min_scheduled_day = 1000
    min_next_day = 1000
    ' loop through schedule elements - numbers of work days
    For i = LBound(arrSchedule) To UBound(arrSchedule)
        If UCase(Left(arrSchedule(i), 2)) <> "WD" Then
            arrSchedule(i) = CInt(arrSchedule(i))
            ' get min_day - in case of will have to use next year
            If min_scheduled_day > arrSchedule(i) Then
                min_scheduled_day = arrSchedule(i)
            End If
    
            ' check if within year of requested date
            ' when found first date that is greater than WdNum of requested date - take it
            If arrSchedule(i) > dDateWdNum Then
                If min_next_day > arrSchedule(i) Then
                    min_next_day = arrSchedule(i)
                End If
            End If
        Else
            ' handler for WD-1, WD-2 etc.
            ' GetWDtoYEC(ddate As Date, wd_offset As String, country_code As String) As Date
            wd_to_yec = GetWDtoMEC(ddate, CStr(arrSchedule(i)), country_code)
            If min_scheduled_day > wd_to_yec Then
                min_scheduled_day = wd_to_yec
            End If
        End If
    Next i
    
    If min_next_day = 1000 Then
        GetClosestScheduledWorkingDayYear = GetDateByYWDnum(Year(ddate) + 1, min_scheduled_day, country_code)
    Else
        GetClosestScheduledWorkingDayYear = GetDateByYWDnum(Year(ddate), min_next_day, country_code)
    End If
    
    Exit Function
    
ErrHandler:
    ' if here - something went wrong
    GetClosestScheduledWorkingDayYear = 0
    ' Resume
End Function

Function GetDateByYWDnum(yr As Integer, wd_num As Integer, country_code As String)
    Dim arrYear()
    Dim arrYWD()
    Dim arrDates()
    Dim i As Long
    
    If ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").DataBodyRange Is Nothing Then
        Exit Function
    End If
    
    arrDates = ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("Date").DataBodyRange.Value
    arrYear = ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("Year").DataBodyRange.Value
    arrYWD = ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("YWD " & country_code).DataBodyRange.Value
    
    For i = LBound(arrYear) To UBound(arrYear)
        If arrYear(i, 1) = yr And arrYWD(i, 1) = wd_num Then
            GetDateByYWDnum = CDate(arrDates(i, 1))
            Exit Function
        End If
    Next i
        
End Function

Function NumberRangeToList(str As String) As String
    Dim arrValues
    Dim i As Long
    
    arrValues = Split(str, "-")
    arrValues(0) = CInt(arrValues(0))
    arrValues(1) = CInt(arrValues(1))
    
    For i = arrValues(0) To arrValues(1)
        NumberRangeToList = NumberRangeToList & "," & i
    Next i
    NumberRangeToList = Mid(NumberRangeToList, 2)
End Function

Function GetWDtoMEC(ddate As Date, wd_offset As String, country_code As String) As Date
' function Get Working Day to Month End Closing
' return Date that offset from Month End on specified number of days
' e.g. -3, => return 3rd working day till end of month
    Dim tmp_str As String
    Dim offset As Integer
    Dim i As Integer
    Dim iDate As Date
        
    tmp_str = Replace(wd_offset, "WD", vbNullString, , , vbTextCompare)
    
    offset = Val(tmp_str)
    ' e.g. -2
    ' seek for last working day of month of 'ddate' argument
    ' then 2 time search previous working day
    ' if it is below ddate - then same for next month
    
    iDate = GetLastWorkingDayOfMonth(Year(ddate), Month(ddate), country_code)
    If CDbl(iDate) <> 0 Then
        For i = -2 To offset Step -1
        ' if WD-1 - then loop will be skipped
            iDate = GetPreviousWorkingDay(iDate, country_code)
        Next i
    End If
    
    GetWDtoMEC = iDate
    If iDate <= ddate And CDbl(iDate) <> 0 Then
        ' call same function but for beginning of next month
        GetWDtoMEC = GetWDtoMEC(WorksheetFunction.EoMonth(ddate, 0) + 1, wd_offset, country_code)
    End If
    
    Exit Function
    
ErrHandler:
    GetWDtoMEC = 0
End Function

Function GetWDtoYEC(ddate As Date, wd_offset As String, country_code As String) As Date
' function Get Working Day to Year End Closing
' return Date that offset from Year End on specified number of days
' e.g. -3, => return 3rd working day till end of year
    Dim tmp_str As String
    Dim offset As Integer
    Dim i As Integer
    Dim iDate As Date
    
    tmp_str = Replace(wd_offset, "WD", vbNullString, , , vbTextCompare)
    
    offset = Val(tmp_str)
    ' e.g. -2
    ' seek for last working day of month of 'ddate' argument
    ' then 2 time search previous working day
    ' if it is below ddate - then same for next month
    
    iDate = GetLastWorkingDayOfYear(Year(ddate), country_code)
    If CDbl(iDate) <> 0 Then
        For i = -2 To offset Step -1
        ' if WD-1 - then loop will be skipped
            iDate = GetPreviousWorkingDay(iDate, country_code)
        Next i
    End If
    
    GetWDtoYEC = iDate
    If iDate <= ddate And CDbl(iDate) <> 0 Then
        ' call same function but for beginning of next month
        GetWDtoYEC = GetWDtoYEC(WorksheetFunction.EoMonth(ddate, 11) + 1, wd_offset, country_code)
    End If
    
    Exit Function
    
ErrHandler:
    GetWDtoYEC = 0
End Function

