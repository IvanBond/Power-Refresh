Attribute VB_Name = "Schedule_WeekWD"
Option Explicit
Option Compare Text

'  GetNextWeekLastWorkingDay(ddate As Date, country_code As String) As Date
'       returns last working day of next week (relative to ddate, first argument)
'           to get last working day of ddate's week, call with ddate-7
'
'  GetNextWeekFirstWorkingDay(ddate As Date, country_code As String) As Date
'       returns first working day of next week (relative to ddate)
'

Function ReplaceLastKeywordWWD(ddate As Date, ScheduleString As String, country_code As String) As String
' replaces "last"-keyword with its Working Day number
' Considers Week that contains ddate

    Dim tmp_str As String
    Dim LastWdOfWeek As Date
    Dim IndexOfLastWorkingDayOfWeek As Integer
    Dim max_date As Date
    Dim ddateWeekNum As Integer
    Dim pos As Long
    Dim LastWdOfWeek_WeekNum As Integer
    
    On Error GoTo ErrHandler
    
    tmp_str = ScheduleString
    With ThisWorkbook.Sheets("Calendar").ListObjects("Calendar")
        
        LastWdOfWeek = GetLastWorkingDayOfWeek(ddate, country_code)   ' date
        ' GetLastWorkingDayOfWeek takes next week automatically
        ' so, if ddate is last working day in week, LastWdOfWeek will be in next week
        
        pos = WorksheetFunction.Match(CDbl(ddate), .ListColumns("Date").DataBodyRange, 0)
        ddateWeekNum = .ListColumns("WeekNum " & country_code).DataBodyRange.Cells(pos, 1).Value
               
        pos = WorksheetFunction.Match(CDbl(LastWdOfWeek), .ListColumns("Date").DataBodyRange, 0)
        LastWdOfWeek_WeekNum = .ListColumns("WeekNum " & country_code).DataBodyRange.Cells(pos, 1).Value
                       
        ' found last WD of Week
        ' if NextRun calculation happens at last day of Week (next run will be in next week)
        ' we have to check if ddate (starting date) is less than Last MWD
        ' because GetLastWorkingDayOfMonth returns next month when current is finished (no more working days)
        If ddate < LastWdOfWeek Then
            IndexOfLastWorkingDayOfWeek = WorksheetFunction.Index( _
                .ListColumns("WWD " & country_code).DataBodyRange, _
                WorksheetFunction.Match(CDbl(LastWdOfWeek), .ListColumns("Date").DataBodyRange, 0))
            tmp_str = Replace(tmp_str, "last", IndexOfLastWorkingDayOfWeek, , , vbTextCompare)
        Else
        End If
    End With
    
    ReplaceLastKeywordWWD = tmp_str
    Exit Function
    
ErrHandler:
    ReplaceLastKeywordWWD = vbNullString
End Function


Function GetWWDbyDate(ddate As Date, country_code As String) As Integer
' index of Working Day within Week of ddate
    Dim pos As Long
    
    On Error GoTo ErrHandler
    
    With ThisWorkbook.Sheets("Calendar").ListObjects("Calendar")
    
        If .DataBodyRange Is Nothing Then
            Exit Function
        End If
        
        pos = WorksheetFunction.Match(CDbl(ddate), .ListColumns("Date").DataBodyRange, 0)
        
        GetWWDbyDate = .ListColumns("WWD " & country_code).DataBodyRange.Cells(pos, 1).Value
        
    End With
        
    Exit Function
ErrHandler:
    GetWWDbyDate = -1
End Function

Function GetFirstWorkingDayOfWeek(ddate As Date, country_code As String) As Date
    ' if entire week is a holiday - return first WD for next week
    Dim idate As Date
    Dim max_date As Date
    Dim pos As Long
    Dim NextWeekStart As Date
    Dim CurrentWeekNum As Byte
    
    On Error GoTo ErrHandler
    
    If ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("Date").DataBodyRange Is Nothing Then
        GetFirstWorkingDayOfWeek = 0
        Exit Function
    End If
    
    max_date = WorksheetFunction.Max(ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("Date").DataBodyRange)
    
    ' position of ddate-7 (previous week)
    pos = WorksheetFunction.Match(CDbl(ddate - 7), ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("Date").DataBodyRange, 0)
    CurrentWeekNum = ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns( _
                        "WeekNum " & country_code).DataBodyRange.Cells(pos, 1).Value
    
    
    idate = ddate - 7 + 1
    Do While idate <= max_date
        ' next year might start - thus search for change of week num
        pos = WorksheetFunction.Match(CDbl(idate), ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("Date").DataBodyRange, 0)
                
        If ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns( _
                "WeekNum " & country_code).DataBodyRange.Cells(pos, 1).Value <> CurrentWeekNum Then
              ' found next week
              ' check if it is a working day
            If ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns( _
                    "WD " & country_code).DataBodyRange.Cells(pos, 1).Value = "Y" Then
                GetFirstWorkingDayOfWeek = idate
                Exit Function
            End If
        End If
        idate = idate + 1
    Loop

    Exit Function
ErrHandler:
    ' if here - something went wrong
    GetFirstWorkingDayOfWeek = 0
End Function

Function GetLastWorkingDayOfWeek(ddate As Date, country_code As String) As Date
' if entire week is a holiday - takes next week and so on until find working days

    Dim idate As Date
    Dim max_date As Date
    Dim pos As Long
    Dim NextWeekStart As Date
    Dim CurrentWeekNum As Byte
    Dim TargetWeekNum As Byte
    
    On Error GoTo ErrHandler
    
    If ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("Date").DataBodyRange Is Nothing Then
        GetLastWorkingDayOfWeek = 0
        Exit Function
    End If
    
    max_date = WorksheetFunction.Max(ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("Date").DataBodyRange)
    
    ' position of ddate
    pos = WorksheetFunction.Match(CDbl(ddate - 7), ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("Date").DataBodyRange, 0)
    CurrentWeekNum = ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns( _
                        "WeekNum " & country_code).DataBodyRange.Cells(pos, 1).Value
                        
    idate = ddate - 7 + 1
    Do While idate <= max_date
        ' next year can start - thus search for change of week num
        pos = WorksheetFunction.Match(CDbl(idate), ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("Date").DataBodyRange, 0)
        If ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns( _
                "WeekNum " & country_code).DataBodyRange.Cells(pos, 1).Value <> CurrentWeekNum Then
            
            ' found start of next week
            ' remember its ID
            If TargetWeekNum = 0 Then
                ' do this only when TargetWeekNum has no value yet
                TargetWeekNum = ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns( _
                    "WeekNum " & country_code).DataBodyRange.Cells(pos, 1).Value
            Else
                ' check if program reached next week following after TargetWeekNum
                If TargetWeekNum <> ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns( _
                    "WeekNum " & country_code).DataBodyRange.Cells(pos, 1).Value Then
                    ' if something has been found before - take it, it is the last working day of Target Week
                    If GetLastWorkingDayOfWeek <> 0 Then
                        Exit Function
                    Else
                        ' probably entire week had no working days
                        ' in such case - re-define TargetWeekNum
                        TargetWeekNum = ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns( _
                                            "WeekNum " & country_code).DataBodyRange.Cells(pos, 1).Value
                    End If
                End If
            End If
            
            ' check if it is a working day
            If ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns( _
                    "WD " & country_code).DataBodyRange.Cells(pos, 1).Value = "Y" Then
                GetLastWorkingDayOfWeek = idate
                'Exit Function
            End If
        End If ' pos <> CurrentWeekNum
        
        idate = idate + 1
    Loop

    Exit Function
ErrHandler:
    ' if here - something went wrong
    GetLastWorkingDayOfWeek = 0
End Function

Function ScheduleStringToListOfDaysWWD(ddate As Date, ScheduleString As String, country_code As String) As String
' shortly: return plain string of days for schedule for Week
' replace keywords: ALL, first, last
' convert ranges into lists of values

    Dim tmp_str As String
    
    On Error GoTo ErrHandler
    
    If ScheduleString = vbNullString Then Exit Function
    
    ' replace keyworkds ALL, first, remove spaces
    tmp_str = PrepareScheduleString(ScheduleString)
        
    If InStr(1, tmp_str, "last", vbTextCompare) > 0 Then
        tmp_str = ReplaceLastKeywordWWD(ddate, tmp_str, country_code)
    End If
         
    ' at this point, we have no "last"-keyword in Schedule String
    ' and have no spaces,
    ' however, might have ranges, and former "last-X", which have to be calculated
    ' e.g. 1..5,10,15,20-3,20-2,20-1,20
    ' or 1..5,10,15,20,20-3..20
    
    ScheduleStringToListOfDaysWWD = OptimiseScheduleString(tmp_str)
    ' now sample string looks like string with comma separated integer values
    ' 1,2,3,4,5,10,15,17,18,19,20
    
    Exit Function

ErrHandler:
    ScheduleStringToListOfDaysWWD = vbNullString
End Function




' old version of function with auto-increment of Week
'Function ReplaceLastKeywordWWD(ddate As Date, ScheduleString As String, country_code As String) As String
'' replaces "last"-keyword with its Working Day number
''
'    Dim tmp_str As String
'    Dim LastWdOfWeek As Date
'    Dim IndexOfLastWorkingDayOfWeek As Integer
'    Dim max_date As Date
'
'    On Error GoTo ErrHandler
'
'    tmp_str = ScheduleString
'    With ThisWorkbook.Sheets("Calendar").ListObjects("Calendar")
'        ' last available date in calendar
'        max_date = WorksheetFunction.Max(.ListColumns("Date").DataBodyRange)
'
'        LastWdOfWeek = GetLastWorkingDayOfWeek(ddate, country_code)   ' date
'        ' found last WD of Week
'        ' if NextRun calculation happens at last day of Week (next run will be in next week)
'        ' we have to check if ddate (starting date) is less than Last MWD
'        ' because GetLastWorkingDayOfMonth returns next month when current is finished (no more working days)
'        If ddate < LastWdOfWeek Then
'            IndexOfLastWorkingDayOfWeek = WorksheetFunction.Index( _
'                .ListColumns("WWD " & country_code).DataBodyRange, _
'                WorksheetFunction.Match(CDbl(LastWdOfWeek), .ListColumns("Date").DataBodyRange, 0))
'            tmp_str = Replace(tmp_str, "last", IndexOfLastWorkingDayOfWeek, , , vbTextCompare)
'        Else
'            ' will use last WD of next month
'            ' check if enough dates in Calendar table
'            ' max_date - last date in calendar, if less then end of next year - cannot rely on such calendar
'            If max_date < CDate(WorksheetFunction.EoMonth(ddate, 1)) Then
'                ' have no enough dates in calendar
'                ' cannot calculate date of next run in such case
'                tmp_str = Replace(tmp_str, "last", vbNullString, , , vbTextCompare)
'                ' "last-2', "last-1" etc. will be negative values after such replacement
'                ' and will be ignored
'            Else
'                ' enough dates in calendar
'                ' replace last with Index of last working day of next Week
'                tmp_str = Replace(tmp_str, "last", _
'                    WorksheetFunction.Index(.ListColumns("WWD " & country_code).DataBodyRange, _
'                        WorksheetFunction.Match( _
'                            CDbl(GetLastWorkingDayOfWeek(ddate, _
'                                country_code)), _
'                            .ListColumns("Date").DataBodyRange, 0)), _
'                        , , vbTextCompare)
'            End If
'        End If
'    End With
'
'    ReplaceLastKeywordWWD = tmp_str
'    Exit Function
'
'ErrHandler:
'    ReplaceLastKeywordWWD = vbNullString
'End Function




