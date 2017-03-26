Attribute VB_Name = "NextRunCalculation"
Option Explicit
'
'

Sub TestNextRun()
    ' Debug.Print GetClosestTime(Now() - 30, CDbl(CDate("00:00:10")))
    Set_Global_Variables
    Dim a
    If Hour(Now()) < Hour(Control_Table.Parent.Cells(19, Control_Table.ListColumns("Execution Time").Range.Column).Value) Then
        a = Date + Control_Table.Parent.Cells(19, Control_Table.ListColumns("Execution Time").Range.Column).Value
    Else
        a = Date + 1 + Control_Table.Parent.Cells(19, Control_Table.ListColumns("Execution Time").Range.Column).Value
    End If
    
    Set_Global_Variables
    Debug.Print Get_Next_Run_DateTime(5)
End Sub

Function Get_Next_Run_DateTime(report_row_id As Long) As Date
    Dim i As Long
    ' simplest logic - next day in the same time (set in column 'Execution Time')
    
    If Control_Table.Parent.Cells(report_row_id, Control_Table.ListColumns("Execution Time").Range.Column).Value = vbNullString Then
        Get_Next_Run_DateTime = CDate("31.12.9999")
        Exit Function
    End If
    
    ' Dummy - next day, same time
    Get_Next_Run_DateTime = Round(Date + 1, 0) + Control_Table. _
        Parent.Cells(report_row_id, Control_Table.ListColumns("Execution Time").Range.Column).Value
    
    ' Default - Recur every X days
    If Control_Table.Parent.Cells(report_row_id, Control_Table.ListColumns("Recur every X days").Range.Column).Value <> vbNullString Then
        ' if only on working days
        If Control_Table.Parent.Cells(report_row_id, Control_Table.ListColumns("Only Working Days").Range.Column).Value = "Y" Then
            ' begin with Start Date
            Get_Next_Run_DateTime = Control_Table.Parent.Cells(report_row_id, Control_Table.ListColumns("Start Date").Range.Column).Value
            
            ' add X working days till reach closest datetime in future
            Do While Get_Next_Run_DateTime < Now()
                For i = 1 To Control_Table.Parent.Cells(report_row_id, Control_Table.ListColumns("Recur every X days").Range.Column).Value
                    Get_Next_Run_DateTime = GetNextWorkingDay(Get_Next_Run_DateTime, _
                        Control_Table.Parent.Cells(report_row_id, Control_Table.ListColumns("WD Country").Range.Column).Value) _
                        
                Next i
            Loop
            
            ' add time
            Get_Next_Run_DateTime = Get_Next_Run_DateTime + _
                Control_Table.Parent.Cells(report_row_id, Control_Table.ListColumns("Execution Time").Range.Column).Value
        Else
        ' just add X days
            Get_Next_Run_DateTime = GetClosestTime(Control_Table.Parent.Cells(report_row_id, Control_Table.ListColumns("Start Date").Range.Column).Value _
                + Control_Table.Parent.Cells(report_row_id, Control_Table.ListColumns("Execution Time").Range.Column).Value, _
                Control_Table.Parent.Cells(report_row_id, Control_Table.ListColumns("Recur every X days").Range.Column).Value)
        End If
    End If
    
    ' recur every X minutes
    If Control_Table.Parent.Cells(report_row_id, Control_Table.ListColumns("Recur every X minutes").Range.Column).Value <> vbNullString Then
        Get_Next_Run_DateTime = GetClosestTime(Control_Table.Parent.Cells(report_row_id, Control_Table.ListColumns("Start Date").Range.Column).Value _
            + Control_Table.Parent.Cells(report_row_id, Control_Table.ListColumns("Execution Time").Range.Column).Value, _
                Control_Table.Parent.Cells(report_row_id, Control_Table.ListColumns("Recur every X minutes").Range.Column).Value / 60 / 24)
        
        ' if calcualted Next Run is later 'To Time' restriction
        If Get_Next_Run_DateTime > Date + _
                Control_Table.Parent.Cells(report_row_id, Control_Table.ListColumns("To Time").Range.Column).Value Then
            
            ' if only on Working days
            If Control_Table.Parent.Cells(report_row_id, Control_Table.ListColumns("Only Working Days").Range.Column).Value = "Y" Then
                If Hour(Now()) < Hour(Control_Table.Parent.Cells(report_row_id, Control_Table.ListColumns("Execution Time").Range.Column).Value) Then
                    ' we should check if today is a working day for particular country and schedule for today execution
                    ' or just get next working day starting from yesterday
                    ' if today is a working day - it will be returned
                    Get_Next_Run_DateTime = GetNextWorkingDay(Date - 1, _
                               Control_Table.Parent.Cells(report_row_id, Control_Table.ListColumns("WD Country").Range.Column).Value) + _
                        Control_Table.Parent.Cells(report_row_id, Control_Table.ListColumns("Execution Time").Range.Column).Value
                Else
                    Get_Next_Run_DateTime = GetNextWorkingDay(Date, _
                               Control_Table.Parent.Cells(report_row_id, Control_Table.ListColumns("WD Country").Range.Column).Value) + _
                        Control_Table.Parent.Cells(report_row_id, Control_Table.ListColumns("Execution Time").Range.Column).Value
                End If
            Else
                ' if calc is done in early AM - can be below 'Execution Time' (maybe was a delay in Scheduler)
                If Hour(Now()) < Hour(Control_Table.Parent.Cells(report_row_id, Control_Table.ListColumns("Execution Time").Range.Column).Value) Then
                    ' then just take Date of code execution and 'Execution Time'
                    Get_Next_Run_DateTime = Date + Control_Table.Parent.Cells(report_row_id, Control_Table.ListColumns("Execution Time").Range.Column).Value
                Else
                    ' otherwise - next Date and 'Execution Time'
                    Get_Next_Run_DateTime = Date + 1 + Control_Table.Parent.Cells(report_row_id, Control_Table.ListColumns("Execution Time").Range.Column).Value
                End If
            End If
        End If
        
    End If
    
End Function

Function GetClosestTime(start As Date, step As Double) As Date
' function returns closest time for Next Run
' resulting time should be in future
' argument Step can be a minute, hour or day - in Double type
    GetClosestTime = start
    Do While GetClosestTime < Now()
        GetClosestTime = GetClosestTime + step
    Loop
End Function

Function GetNextRunWorkingDays(rng As Range)
' TODO - function definition and purpose
    Dim arrWD
    Dim minimal
    Dim i As Byte
    Dim bFound As Boolean
    On Error Resume Next
    ' rng with list of working days
    arrWD = Split(Replace(rng.Value, " ", ""), ",")
        
    For i = LBound(arrWD) To UBound(arrWD)
        ' what if next month?
        If CInt(arrWD(i)) >= ThisWorkbook.Sheets("Calendar").Cells(2, 5) Then
            bFound = True
            ' "=MIN(IF(Calendar[Date]>TODAY(),1,1000) * IF(Calendar[WD RU]=""Y"", 1, 1000) * Calendar[Date] )"
            
        End If
    Next i
End Function

Function GetNextRunMonthly()
    
End Function

Function GetNextRunWeekly()
    
End Function

Function GetNextRunDaily()
    
End Function

Function GetClosestWorkingDay(ddate As Date, country As String)
        ' use GetNextWorkingDay for yesterday ?
End Function

'
' Logic draft
' if Frequency = "D"
' If only on working days
'   if Country is not null
'       if column with country found in Calendar table
'           Get Date of next country-specific WD
'           how: find TODAY in list of dates, then loop through days until Y in WD [Country]
'
'       if Month or Year = ALL
'           just use found Date (+Execution time)
'       if Year WD not empty
'           ? get next country-specific WD key, e.g. 2016178
'           get Number of next WD - 178
'           get arrOfYearWorkingDays
'           e.g. 1-5, 21, 73-79 -> 1,2,3,4,5,21,73,74,75,76,77,78,79
'           + word: last - Get last WD number of cur year
'           if only last WD -> bOnlyLastWD = True
'           ' 1: next WD <= max( element of arrOfYearWorkingDays)
            ' find MIN WD that is >= next WD
            ' 2: next WD > max( element of arrOfYearWorkingDays)
            ' next run will be in next Year
            ' if only last WD - get it for next Year
            ' [Next Year] & MIN arrOfYearWorkingDays
'       if Month WD is not empty
'       get Date of next WD - can be next month / next Year
'       get Number of next WD
'       get arrMonthWD + Last word
'
'       if not ALL
'           arrDays = Split
'               for each element -
'                   if el contains "-" - then it is range
' build arrFullDays
'  find el that is
' 20160112
' 20170115 - key of working day, 15th working day of Jan 2017
' 201773 - key of working day, 73th working day of 2017 year

