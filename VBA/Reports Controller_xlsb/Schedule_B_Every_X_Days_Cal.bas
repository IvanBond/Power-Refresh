Attribute VB_Name = "Schedule_B_Every_X_Days_Cal"
Option Explicit
Option Compare Text

' ========================================================= NON WORKING DAYS =========================================================
' ====================================================================================================================================

' GetClosestDateTimeRecurXDays(StartingDate As Date, _
                            RecurXDays As Integer, _
                            ExecutionTime As Double, _
                            Optional RecurXMinutes As Double, _
                            Optional ToTime As Date, _
                            Optional NotEarlierThanFixed As Date, _
                            Optional Report_Row_ID As Long) As Date
'
'
' GetDateRecurXDays(StartingDate As Date, _
                           RecurXDays As Integer, _
                           Optional NotEarlierThanFixed As Date, _
                           Optional Report_Row_ID As Long) As Date
'


' Main function
Function GetClosestDateTimeRecurXDays(StartingDate As Date, _
                            RecurXDays As Integer, _
                            ExecutionTime As Double, _
                            Optional RecurXMinutes As Double, _
                            Optional ToTime As Date, _
                            Optional NotEarlierThanFixed As Date, _
                            Optional Report_Row_ID As Long) As Date
' returns a DateTime that suits scenarios:
'     1) Recur X Days
'     2) Recur X Minutes (Optional)

    Dim tempDate As Date
    Dim bTodayIsFine As Boolean
    Dim NotEarlierThan As Date
    Dim ToTimeLimit As Date
        
    On Error GoTo ErrHandler
    
    NotEarlierThan = WorksheetFunction.Max(Now, NotEarlierThanFixed)
    
    If RecurXMinutes <> 0 Then
    ' when scheduled with RecurXMinutes it may still be executed today
        
        ToTimeLimit = IIf(ToTime <> 0, TimeValue(Hour(ToTime) & ":" & Minute(ToTime) & ":" & Second(ToTime)), TimeValue("23:59:59"))
        
        bTodayIsFine = (Date + ToTimeLimit > NotEarlierThan)
        If bTodayIsFine Then
            tempDate = GetDateRecurXDays(StartingDate, RecurXDays, NotEarlierThan)   ' find date
        Else
            tempDate = GetDateRecurXDays(StartingDate, RecurXDays, _
                WorksheetFunction.Max(Date + 1, NotEarlierThan)) ' find date
        End If
        
        If tempDate = Date Then
        ' if today - need to find closest time and compare with ToTime limit
            tempDate = GetClosestTime(tempDate + ExecutionTime, RecurXMinutes / 24 / 60, NotEarlierThan)  ' find time
            If tempDate > Date + ToTimeLimit Then
            ' if reached limit - move to nextXdays
                tempDate = GetDateRecurXDays(Date, RecurXDays, WorksheetFunction.Max(Date + 1, NotEarlierThan))  ' find next date
                ' left to add ExecutionTIme
                GetClosestDateTimeRecurXDays = DateValue(tempDate) + ExecutionTime
                Exit Function
            Else
            ' haven't reached limit yet
            ' no need to add ExecutionTime
                GetClosestDateTimeRecurXDays = tempDate
                Exit Function
            End If
        Else
        ' not today - left to find ExecutionTime
            GetClosestDateTimeRecurXDays = GetClosestTime(tempDate + ExecutionTime, RecurXMinutes / 24 / 60, NotEarlierThan)  ' find time
            Exit Function
        End If
    Else
    ' if no schedule on minutes
        tempDate = GetDateRecurXDays(StartingDate, RecurXDays, WorksheetFunction.Max(Date + 1, NotEarlierThan), Report_Row_ID)
        GetClosestDateTimeRecurXDays = DateValue(tempDate) + ExecutionTime
        Exit Function
    End If
    
Exit_Function:
    
    Exit Function
    
ErrHandler:
    If Err.Number <> 0 Then
        Debug.Print Now, "GetClosestDateTimeRecurXDays", Err.Number & ": " & Err.Description
        Err.Clear
    End If
    GetClosestDateTimeRecurXDays = CDate("9999-12-31")
    GoTo Exit_Function
    Resume
End Function

' Supporting function
Function GetDateRecurXDays(StartingDate As Date, _
                           RecurXDays As Integer, _
                           Optional NotEarlierThanFixed As Date, _
                           Optional Report_Row_ID As Long) As Date
                           
' returns date for scenario: Recur X Days (for non-working days)
' may return today

    Dim NotEarlierThan As Date
    
    ' cut Time off from NotEarlierThanFixed
    NotEarlierThan = WorksheetFunction.Max(Date, DateValue(NotEarlierThanFixed))
    
    GetDateRecurXDays = StartingDate
    
    Do While GetDateRecurXDays < NotEarlierThan
        GetDateRecurXDays = GetDateRecurXDays + RecurXDays
    Loop
    
End Function
