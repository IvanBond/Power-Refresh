Attribute VB_Name = "Schedule_A_BasicFunctions"
Option Explicit
Option Compare Text

'  GetClosestTime(start As Date, step As Double, Optional NotEarlierThanFixed As Date) As Date
'       returns Closest time in future, not earlier than Now or 'NotEarlierThanFixed' parameter.
'
'  GetNextWorkingDay(ddate As Date, country_code As String) As Date
'       returns next working date
'
'  GetPreviousWorkingDay(ddate As Date, country_code As String) As Date
'       returns previous working date

Function GetClosestTime(start As Date, _
                        step As Double, _
                        Optional NotEarlierThanFixed As Date) As Date

' Function returns closest datetime that is greater or equal to system datetime. Or 'NotEarlierThanFixed'.
' Arguments
' start: starting datetime
' step: can be a minute, hour or day - in Double type
' NotEarlierThanFixed: specify datetime if you want to get closest datetime that is greater than or equal to 'NotEarlierThanFixed'.
' By default, function compares with system datetime.

    Dim NotEarlierThan As Date
    
    If (NotEarlierThanFixed = 0) Then
        NotEarlierThan = Now
    Else
        NotEarlierThan = NotEarlierThanFixed
    End If
    
    GetClosestTime = start
    Do While GetClosestTime < NotEarlierThan
        GetClosestTime = GetClosestTime + step
    Loop
    
End Function

Function GetNextWorkingDay(ddate As Date, country_code As String) As Date
    Dim idate As Date
    Dim max_date As Date
    Dim min_date As Date
    Dim pos As Long
    
    On Error GoTo ErrHandler
    
    ' ddate is a starting date
    If ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("Date").DataBodyRange Is Nothing Then
        Debug.Print Now, "GetNextWorkingDay", "Table 'Calendar' is empty"
        GetNextWorkingDay = 0
        Exit Function
    End If
    
    If country_code = vbNullString Then
        Debug.Print Now, "GetNextWorkingDay", "'Country code' is not provided."
        GetNextWorkingDay = 0
        Exit Function
    End If
    
    min_date = WorksheetFunction.Min(ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("Date").DataBodyRange)
    max_date = WorksheetFunction.Max(ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("Date").DataBodyRange)
    
    If ddate < min_date Or ddate > max_date Then
        Debug.Print Now, "GetNextWorkingDay", "'" & Format(ddate, "yyyy-mm-dd") & "' is out of Calendar range."
        ddate = min_date - 1
    End If
    
    idate = ddate + 1
    Do While idate <= max_date
        On Error Resume Next
        pos = WorksheetFunction.Match(CDbl(DateValue(idate)), ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("Date").DataBodyRange, 0)
        If Err.Number <> 0 Then
            Debug.Print Now, "GetNextWorkingDay", "Date " & Format(CDbl(DateValue(idate)), "yyyy-mm-dd"); " hasn't been found in calendar"
            Err.Clear
            GoTo ErrHandler
        End If
        On Error GoTo ErrHandler
        
        If ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("WD " & country_code).DataBodyRange.Cells(pos, 1).Value = "Y" Then
            GetNextWorkingDay = idate
            Exit Function
        End If
        
        idate = idate + 1
    Loop

Exit_Function:
    Exit Function

ErrHandler:
    If Err.Number <> 0 Then
        Debug.Print Now, "GetNextWorkingDay", Err.Number & ": " & Err.Description
        Err.Clear
    End If
    ' if here - something went wrong
    GetNextWorkingDay = 0
    GoTo Exit_Function
    Resume
End Function

Function GetPreviousWorkingDay(ddate As Date, country_code As String) As Date
    Dim idate As Date
    Dim min_date As Date
    Dim pos As Long
    
    On Error GoTo ErrHandler
    
    If ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("Date").DataBodyRange Is Nothing Then
        GetPreviousWorkingDay = 0
        Exit Function
    End If
    
    min_date = WorksheetFunction.Min(ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("Date").DataBodyRange)
    idate = ddate - 1
    Do While idate >= min_date
        pos = WorksheetFunction.Match(CDbl(idate), ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("Date").DataBodyRange, 0)
        
        If ThisWorkbook.Sheets("Calendar").ListObjects("Calendar").ListColumns("WD " & country_code).DataBodyRange.Cells(pos, 1).Value = "Y" Then
            GetPreviousWorkingDay = idate
            Exit Function
        End If
        
        idate = idate - 1
    Loop
    
Exit_Function:
    Exit Function

ErrHandler:
    If Err.Number <> 0 Then
        Debug.Print Now, "GetNextWorkingDay", Err.Number & ": " & Err.Description
        Err.Clear
    End If
    ' if here - something went wrong
    GetPreviousWorkingDay = 0
    GoTo Exit_Function
    Resume
End Function

