Attribute VB_Name = "Schedule_YearWD"
Option Explicit
Option Compare Text

'  GetFirstWorkingDayOfYear(yr As Integer, country_code As String) as date
'       returns first working day of year

'  GetLastWorkingDayOfYear(yr As Integer, country_code As String) As Date
'       returns last working day of year

Function GetClosestScheduledWorkingDayYear(ddate As Date, Schedule As String, country_code As String) As Date
' Function gets ddate - starting point of calculation
' Schedule: string with possible patterns
' first,3-10,WD-5, last
' first is a keyword, will be replaced with 1
' last is a keyword, which means 'last working day'
' WD-1 is a last working day
' WD-2 is [Working Day-2] prior working day to [WD-1]
' 3-10 will be converted into string 3,4,5,6,7,8,9,10
' note: WD-0 doesn't exist, as WD-X is a "number of workdays till Financial Month End Closing"
' which usually starts from the beginning of new month (WD1)
 
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
    Dim date_of_wd_to_yec As Date
    Dim wd_to_yec As Integer
    Dim max_date As Date
    
    
    If Schedule = vbNullString Then Exit Function
    
    On Error GoTo ErrHandler
    ' get row of requested date
    
    With ThisWorkbook.Sheets("Calendar").ListObjects("Calendar")
        
        ' last available date in calendar
        max_date = WorksheetFunction.Max(.ListColumns("Date").DataBodyRange)
        
        pos = WorksheetFunction.Match(CDbl(ddate), .ListColumns("Date").DataBodyRange, 0)
        ' found position (row) or ddate in Calendar table
        
        ' check if it is a working day
        dDateWdNum = .ListColumns("YWD " & country_code).DataBodyRange.Cells(pos, 1).Value
        ' if beginning of the year, ddate may have 0-index of YWD (index of consequent WD in a year)
        If dDateWdNum = 0 Then
            ' working day number of requested date
            ' if it is holiday - then search for the closest (next) working date
            
            pos = WorksheetFunction.Match(CDbl(GetNextWorkingDay(ddate, country_code)), _
                .ListColumns("Date").DataBodyRange, 0)
            ' got position of WD
            dDateWdNum = .ListColumns("YWD " & country_code).DataBodyRange.Cells(pos, 1).Value
            ' got index of working day in the year
        End If
        
        ' parse Schedule
        tmp_str = Replace(Schedule, " ", vbNullString) ' replace all spaces
        tmp_str = Replace(tmp_str, "first", "1", , , vbTextCompare)
        
        ' replace keyword 'last' with its number
        If InStr(1, tmp_str, "last", vbTextCompare) > 0 Then
            
            LastWdOfCurrentYear = GetLastWorkingDayOfYear(Year(ddate), country_code)
            ' found last WD of year
            ' if NextRun calculation happens during last days of year, which are already holidays
            ' we have to check if ddate is less than Last YWD
            If ddate < LastWdOfCurrentYear Then
                NumOrLastWdOfCurrentYear = WorksheetFunction.Index( _
                    .ListColumns("YWD " & country_code).DataBodyRange, _
                    WorksheetFunction.Match(CDbl(LastWdOfCurrentYear), .ListColumns("Date").DataBodyRange, 0))
                tmp_str = Replace(tmp_str, "last", NumOrLastWdOfCurrentYear, , , vbTextCompare)
            Else
                ' last WD of next year
                ' if next year is not in Calendar table, here might be an error
                ' check if enough dates in Calendar table
                If CDate(Year(ddate) + 1 & "-12-31") > max_date Then
                    ' no end of next year, just remove "last"
                    tmp_str = Replace(tmp_str, "last", vbNullString, , , vbTextCompare)
                    ' if it was the only word in Schedule string, function returns 0-date
                Else
                    tmp_str = Replace(tmp_str, "last", _
                        WorksheetFunction.Index(.ListColumns("YWD " & country_code).DataBodyRange, _
                            WorksheetFunction.Match(CDbl(GetLastWorkingDayOfYear(Year(ddate) + 1, country_code)), _
                                .ListColumns("Date").DataBodyRange, 0)), _
                            , , vbTextCompare)
                End If
            End If
        End If
    
    End With
    
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
                ' transform range to list of values
                tmp_str = tmp_str & "," & NumberRangeToList(Trim(CStr(arrTmp(i))))
            Else
            ' otherwise - take single value
                tmp_str = tmp_str & "," & Trim(arrTmp(i))
            End If
        Next i
        tmp_str = Mid(tmp_str, 2)
    End If
    
    ' prepared schedule to new array
    arrSchedule = Split(tmp_str, ",", -1, vbTextCompare)
    
    ' at this moment arrSchedule is possibly "not sorted" array of values
    ' e.g. 1,2,3,218,WD-3,WD-2,WD-1
    ' for instance, YWD of ddate is 215
    ' we have to find minimal YWD which is greater 215
    ' in theory it can be one of those WD-3,WD-2,WD1
    ' when number of working days in Year is less than 220
    
    ' dummy values
    min_scheduled_day = 1000 ' store value from Schedule array, as it is unsorted
    ' after loop var will store min scheduled day, which will be used for next year
    
    min_next_day = 1000
    ' if success, we will find day in year of ddate,
    ' var will store closest day
    
    ' loop through schedule elements - numbers of working days
    ' as array in theory is unsorted - have to go through all elements
    For i = LBound(arrSchedule) To UBound(arrSchedule)
        If Trim(arrSchedule(i)) <> vbNullString Then
            If UCase(Left(arrSchedule(i), 2)) <> "WD" Then
                ' if ordinary value
                arrSchedule(i) = CInt(arrSchedule(i)) ' trying to get integer from it
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
                ' GetWDtoEOY(ddate As Date, wd_offset As String, country_code As String) As Date
                date_of_wd_to_yec = GetWDtoEOY(ddate, CStr(arrSchedule(i)), country_code)
                
                ' now have to find YWD index of found date
                pos = WorksheetFunction.Match(CDbl(date_of_wd_to_yec), _
                    ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("Date").DataBodyRange, 0)
                wd_to_yec = ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("YWD " & country_code).DataBodyRange.Cells(pos, 1).Value
                
                If min_scheduled_day > wd_to_yec Then
                    min_scheduled_day = wd_to_yec
                End If
            End If ' If UCase(Left(arrSchedule(i), 2)) <> "WD" Then
        End If ' If Trim(arrSchedule(i)) <> vbNullString Then
    Next i
    
    ' if after loop we couldn't find min_next_day (closest day in ddate's Year
    ' we will have to schedule report on next year using min_scheduled_day found in Schedule array
    If min_next_day = 1000 Then
    '   schedule on next year
        If min_scheduled_day <> 1000 Then ' was found
            GetClosestScheduledWorkingDayYear = GetDateByYWDnum(Year(ddate) + 1, min_scheduled_day, country_code)
        Else
            GoTo ErrHandler
        End If
    Else
    ' schedule on closest day
        GetClosestScheduledWorkingDayYear = GetDateByYWDnum(Year(ddate), min_next_day, country_code)
    End If
    
    Exit Function
    
ErrHandler:
    ' if here - something went wrong
    GetClosestScheduledWorkingDayYear = 0
    ' Resume
End Function

Function GetFirstWorkingDayOfYear(yr As Integer, country_code As String) As Date
' gets Year as argument
' starts search from 1st Calendar Day of Year
'
    Dim idate As Date
    Dim max_date As Date
    Dim pos As Long
    
    On Error GoTo ErrHandler
    
    With ThisWorkbook.Sheets("Calendar").ListObjects("Calendar")
        If .ListColumns("Date").DataBodyRange Is Nothing Then
            GetFirstWorkingDayOfYear = 0
            Exit Function
        End If
        
        max_date = WorksheetFunction.Max(.ListColumns("Date").DataBodyRange)
        
        idate = CDate(yr & "-01-01")
        Do While idate <= max_date
            ' will be an error if date not found
            pos = WorksheetFunction.Match(CDbl(idate), .ListColumns("Date").DataBodyRange, 0)
            
            
            If .ListColumns("WD " & country_code).DataBodyRange.Cells(pos, 1).Value = "Y" Then
                GetFirstWorkingDayOfYear = idate
                Exit Function
            End If
            idate = idate + 1
        Loop
    End With
    
ErrHandler:
    ' if here - something went wrong
    GetFirstWorkingDayOfYear = 0
End Function

Function GetLastWorkingDayOfYear(yr As Integer, country_code As String) As Date
' function gets Year as argument
' starts search from last Calendar Day of the year
' goes back till find working day or
' reach minimal date in calendar table

    Dim idate As Date
    Dim min_date As Date
    Dim pos As Long
    
    On Error GoTo ErrHandler
    
    With ThisWorkbook.Sheets("Calendar").ListObjects("Calendar")
    
        If .ListColumns("Date").DataBodyRange Is Nothing Then
            GetLastWorkingDayOfYear = 0
            Exit Function
        End If
        
        min_date = WorksheetFunction.Min(.ListColumns("Date").DataBodyRange)
        
        idate = CDate(yr & "-12-31")
        Do While idate >= min_date
            pos = WorksheetFunction.Match(CDbl(idate), .ListColumns("Date").DataBodyRange, 0)
            
            If .ListColumns("WD " & country_code).DataBodyRange.Cells(pos, 1).Value = "Y" Then
                GetLastWorkingDayOfYear = idate
                Exit Function
            End If
            idate = idate - 1
        Loop
        
    End With
ErrHandler:
    ' if here - something went wrong
    GetLastWorkingDayOfYear = 0
End Function

Function GetYWDbyDate(ddate As Date, country_code As String) As Integer
' index of working day withing year of ddate
    Dim pos As Long
    
    On Error GoTo ErrHandler
    
    With ThisWorkbook.Sheets("Calendar").ListObjects("Calendar")
    
        If .DataBodyRange Is Nothing Then
            Exit Function
        End If
        
        pos = WorksheetFunction.Match(CDbl(ddate), .ListColumns("Date").DataBodyRange, 0)
        
        GetYWDbyDate = .ListColumns("YWD " & country_code).DataBodyRange.Cells(pos, 1).Value
        
    End With
    
    Exit Function
ErrHandler:
    GetYWDbyDate = -1
End Function

Function GetDateByYWDnum(yr As Integer, wd_num As Integer, country_code As String) As Date
' YWD num - index of working day in year
    Dim arrYear()
    Dim arrYWD()
    Dim arrDates()
    Dim i As Long
    
    On Error GoTo ErrHandler
    
    With ThisWorkbook.Sheets("Calendar").ListObjects("Calendar")
    
        If .DataBodyRange Is Nothing Then
            Exit Function
        End If
        
        ' get values from sheet to array to speed up loop
        arrDates = .ListColumns("Date").DataBodyRange.Value
        arrYear = .ListColumns("Year").DataBodyRange.Value
        arrYWD = .ListColumns("YWD " & country_code).DataBodyRange.Value
        
    End With
    
    For i = LBound(arrYear) To UBound(arrYear)
        If arrYear(i, 1) = yr And arrYWD(i, 1) = wd_num Then
            GetDateByYWDnum = CDate(arrDates(i, 1))
            Exit Function
        End If
    Next i
    
ErrHandler:
    ' if here - something went wrong
    GetDateByYWDnum = 0
End Function

Function GetWDtoEOY(ddate As Date, wd_offset As String, country_code As String) As Date
' function Get Working Day to Year End Closing
' return Date that offset from Year End on specified number of days
' ddate is a starting point
' wd_offset is expected in formats
' WD-1, WD-2 or -1, -2
' e.g. -3, => return 3rd working day till end of year
' note: WD-0 doesn't exist, as WD-X is a "number of workdays till Financial Month End Closing"
' which usually starts from the beginning of new month (WD1)

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
    ' then 2 time search previous working day
    ' if it is below ddate - then same for next month
    
    idate = GetLastWorkingDayOfYear(Year(ddate), country_code)
    If CDbl(idate) <> 0 Then
        For i = -2 To offset Step -1
        ' if WD-1 - then loop will be skipped
            idate = GetPreviousWorkingDay(idate, country_code)
        Next i
    End If
    
    GetWDtoEOY = idate
    If idate <= ddate And CDbl(idate) <> 0 Then
        ' call same function but for month in next year (if Calendar is provided)
        If WorksheetFunction.EoMonth(ddate, 11) + 1 < max_date Then
            GetWDtoEOY = GetWDtoEOY(WorksheetFunction.EoMonth(ddate, 11) + 1, wd_offset, country_code)
        Else
        '   no dates in Calendar table
            GoTo ErrHandler
        End If
    End If
    
    Exit Function
    
ErrHandler:
    GetWDtoEOY = 0
End Function

Function ReplaceLastKeywordYWD(ddate As Date, ScheduleString As String, country_code As String) As String
' replaces "last"-keyword with its Working Day number
'
    Dim tmp_str As String
    Dim LastWdOfYear As Date
    Dim IndexOfLastWorkingDayOfYear As Integer
    Dim max_date As Date
    
    On Error GoTo ErrHandler
    
    tmp_str = ScheduleString
    With ThisWorkbook.Sheets("Calendar").ListObjects("Calendar")
        ' last available date in calendar
        max_date = WorksheetFunction.Max(.ListColumns("Date").DataBodyRange)
        
        LastWdOfYear = GetLastWorkingDayOfYear(Year(ddate), country_code) ' date
        ' found last WD of year
        ' if NextRun calculation happens during last days of year, which are already holidays
        ' we have to check if ddate (starting date) is less than Last YWD
        ' because GetLastWorkingDayOfYear returns next year when current is finished (no more working days)
        If ddate < LastWdOfYear Then
            IndexOfLastWorkingDayOfYear = WorksheetFunction.Index( _
                .ListColumns("YWD " & country_code).DataBodyRange, _
                WorksheetFunction.Match(CDbl(LastWdOfYear), .ListColumns("Date").DataBodyRange, 0))
            tmp_str = Replace(tmp_str, "last", IndexOfLastWorkingDayOfYear, , , vbTextCompare)
        Else
            ' will use last WD of next year (to cover scenario when only elements with "last"-keyword used)
            ' check if enough dates in Calendar table
            ' max_date - last date in calendar, if less then end of next year - cannot rely on such calendar
            If max_date < CDate(Year(ddate) + 1 & "-12-31") Then
                ' not an end of next year, just remove "last"
                tmp_str = Replace(tmp_str, "last", vbNullString, , , vbTextCompare)
                ' "last-2', "last-1" etc. will be negative values after such replacement
                ' and will be ignored
            Else
                ' enough dates in calendar
                ' replace last with Index of last working day of next year
                tmp_str = Replace(tmp_str, "last", _
                    WorksheetFunction.Index(.ListColumns("YWD " & country_code).DataBodyRange, _
                        WorksheetFunction.Match(CDbl(GetLastWorkingDayOfYear(Year(ddate) + 1, country_code)), _
                            .ListColumns("Date").DataBodyRange, 0)), _
                        , , vbTextCompare)
            End If
        End If
    End With
    
    ReplaceLastKeywordYWD = tmp_str
    Exit Function
    
ErrHandler:
    ReplaceLastKeywordYWD = vbNullString
End Function

Function ScheduleStringToListOfDaysYWD(ddate As Date, ScheduleString As String, country_code As String) As String
' shortly: return plain string of days for schedule
' replace keywords: ALL, first, last
' convert ranges into lists of values

    Dim tmp_str As String
    
    On Error GoTo ErrHandler
    
    If ScheduleString = vbNullString Then Exit Function
    
    ' replace keyworkds ALL, first, remove spaces
    tmp_str = PrepareScheduleString(ScheduleString)
        
    If InStr(1, tmp_str, "last", vbTextCompare) > 0 Then
        tmp_str = ReplaceLastKeywordYWD(ddate, tmp_str, country_code)
    End If
        
    ' at this point, we have no "last"-keyword in Schedule String
    ' and have no spaces,
    ' however, might have ranges, and former "last-X", which have to be calculated
    ' e.g. 1..5,10,15,20,240-4,240-3,240-2,240-1,240
    ' or 1..5,10,15,20,240-4..240
    
    ScheduleStringToListOfDaysYWD = OptimiseScheduleString(tmp_str)
    ' now sample string looks like string with comma separated integer values
    ' 1,3,4,5,6,7,8,9,10,242,243,244,246,247,250,251,252
        
    Exit Function

ErrHandler:
    ScheduleStringToListOfDaysYWD = vbNullString
End Function

