Attribute VB_Name = "Schedule_B_Every_X_Days_WD"
Option Explicit
Option Compare Text

' ========================================================= WORKING DAYS =========================================================
' ================================================================================================================================

'
' GetDateRecurXWorkingDays(StartingDate As Date, _
                            RecurXDays As Integer, _
                            WDCountry As String, _
                            Optional NotEarlierThanFixed As Date, _
                            Optional Report_Row_ID As Long) As Date
'
' GetClosestDateTimeRecurXWorkingDays(StartingDate As Date, _
                            RecurXDays As Integer, _
                            WDCountry As String, _
                            ExecutionTime As Double, _
                            Optional RecurXMinutes As Double, _
                            Optional ToTime As Date, _
                            Optional NotEarlierThanFixed As Date, _
                            Optional Report_Row_ID As Long) as Date
'
'
'

' Main function
Function GetClosestDateTimeRecurXWorkingDays(StartingDate As Date, _
                            RecurXDays As Integer, _
                            WDCountry As String, _
                            ExecutionTime As Double, _
                            Optional RecurXMinutes As Double, _
                            Optional ToTime As Date, _
                            Optional NotEarlierThanFixed As Date, _
                            Optional Report_Row_ID As Long) As Date
                            
    Dim tempDate As Date
    Dim bTodayIsFine As Boolean
    Dim NotEarlierThan As Date
    Dim ToTimeLimit As Date
        
    NotEarlierThan = WorksheetFunction.Max(Now, NotEarlierThanFixed)
        
    If RecurXMinutes <> 0 Then
    ' when scheduled with RecurXMinutes it may still be executed today
    
        ToTimeLimit = IIf(ToTime <> 0, TimeValue(Hour(ToTime) & ":" & Minute(ToTime) & ":" & Second(ToTime)), TimeValue("23:59:59"))
                
        bTodayIsFine = (Date + ToTimeLimit > NotEarlierThan)
        ' as this function also checks Time - have to check 'ToTime' limit
        
        tempDate = DateValue(StartingDate)
        If bTodayIsFine Then
            tempDate = GetDateRecurXWorkingDays(tempDate, RecurXDays, WDCountry, NotEarlierThan)    ' find date
        Else
            tempDate = GetDateRecurXWorkingDays(tempDate, RecurXDays, WDCountry, _
                WorksheetFunction.Max(Date + 1, NotEarlierThan)) ' find date
        End If
        
        ' when working day falls on today
        If tempDate = Date Then
        ' if today - need to find closest time and compare with ToTime limit
            tempDate = GetClosestTime(tempDate + ExecutionTime, RecurXMinutes / 24 / 60, NotEarlierThan)  ' find time
            
            If tempDate > Date + ToTimeLimit Then
            ' if reached limit - find next working day
                tempDate = GetDateRecurXWorkingDays(Date, RecurXDays, WDCountry, Date + 1) ' find next date
                ' left to add ExecutionTIme
                GetClosestDateTimeRecurXWorkingDays = DateValue(tempDate) + ExecutionTime
                Exit Function
            Else
            ' before ToTime limit
            ' no need to add ExecutionTime
                GetClosestDateTimeRecurXWorkingDays = tempDate
                Exit Function
            End If
        Else
        ' not today - left to find ExecutionTime
            GetClosestDateTimeRecurXWorkingDays = GetClosestTime(tempDate + ExecutionTime, RecurXMinutes / 24 / 60, NotEarlierThan) ' find time
            Exit Function
        End If
    Else
    ' if no schedule on minutes
    ' when no RecurXMinutes parameter - Today can't be OK
        ' so search starting from next day
        tempDate = GetDateRecurXWorkingDays(DateValue(StartingDate), RecurXDays, WDCountry, _
                WorksheetFunction.Max(Date + 1, NotEarlierThan))
        GetClosestDateTimeRecurXWorkingDays = DateValue(tempDate) + ExecutionTime
        Exit Function
    End If

Exit_Function:
    
    Exit Function
    
ErrHandler:
    If Err.Number <> 0 Then
        Debug.Print Now, "GetClosestDateTimeRecurXWorkingDays", Err.Number & ": " & Err.Description
        Err.Clear
    End If
    GetClosestDateTimeRecurXWorkingDays = CDate("9999-12-31")
    GoTo Exit_Function
    Resume
End Function

' Supporting Function
Function GetDateRecurXWorkingDays(StartingDate As Date, _
                            RecurXDays As Integer, _
                            WDCountry As String, _
                            Optional NotEarlierThanFixed As Date, _
                            Optional Report_Row_ID As Long) As Date
' return a date that suits criteria for scenario: recur X working days
' may return Today

    Dim i As Integer
    Dim NotEarlierThan As Date
    
    NotEarlierThan = WorksheetFunction.Max(Date, DateValue(NotEarlierThanFixed))
    
    On Error GoTo ErrHandler
    
    GetDateRecurXWorkingDays = StartingDate
        
    Do While GetDateRecurXWorkingDays < NotEarlierThan
        For i = 1 To RecurXDays
            GetDateRecurXWorkingDays = GetNextWorkingDay(GetDateRecurXWorkingDays, WDCountry)
            
            If GetDateRecurXWorkingDays = 0 Then
                Debug.Print Now, "GetDateRecurXWorkingDays", "Couldn't calculate Next Working Day for Country '" _
                    & WDCountry & "' and date '" & Format(GetDateRecurXWorkingDays, "yyyy-mm-dd") & "'"
                GoTo ErrHandler
            End If
        Next i
    Loop
    
Exit_Function:
    
    Exit Function
    
ErrHandler:
    If Err.Number <> 0 Then
        Debug.Print Now, "GetDateRecurXWorkingDays", Err.Number & ": " & Err.Description
        Err.Clear
    End If
    GetDateRecurXWorkingDays = CDate("9999-12-31")
    GoTo Exit_Function
    Resume
End Function
