Attribute VB_Name = "Schedule_C_StringConversion"
Option Explicit
Option Compare Text

' Module with basic functions used for conversion of strings

Function PrepareScheduleString(ScheduleString As String) As String
    Dim tmp_str As String
    tmp_str = Replace(ScheduleString, " ", vbNullString) ' replace all spaces
    tmp_str = Replace(tmp_str, "first", "1", , , vbTextCompare)
    PrepareScheduleString = Replace(tmp_str, "ALL", "1..last", , , vbTextCompare)
End Function

Function OptimiseScheduleString(ScheduleString As String) As String
' polishing of string: evaluate former "last-X" and expand intervals
' function that optimises strings like
' 1..5,10,15,20,240-4,240-3,240-2,240-1,240
' or 1..5,10,15,20,240-4..240
' by converting ranges to list of values

    Dim arrTmp
    Dim tmp_str As String
    Dim i As Long
    
    On Error GoTo ErrHandler
    
    arrTmp = Split(ScheduleString, ",", -1, vbTextCompare)
    ' re-create tmp_str
    For i = LBound(arrTmp) To UBound(arrTmp)
    ' check if conatins "-" - then have to calculate value
        If InStr(1, CStr(arrTmp(i)), "-", vbTextCompare) > 0 Then
            ' check if it contains ".."
            If InStr(1, CStr(arrTmp(i)), "..", vbTextCompare) > 0 Then
                ' have to split values and calculate each component separately
                ' and have to convert range to comma separated list of values
                tmp_str = tmp_str & "," & NumberRangeToList( _
                        CStr(Application.Evaluate( _
                            Left(CStr(arrTmp(i)), _
                                  InStr(1, CStr(arrTmp(i)), "..", vbTextCompare) - 1)) & ".." & _
                            Application.Evaluate( _
                                Mid(CStr(arrTmp(i)), _
                                      InStr(1, CStr(arrTmp(i)), "..", vbTextCompare) + 2))))
            Else
                ' just calculate value
                ' Application.Evaluate("240-2")
                tmp_str = tmp_str & "," & Application.Evaluate(CStr(arrTmp(i)))
            End If
            
        Else
        ' otherwise - take single value
            If InStr(1, CStr(arrTmp(i)), "..", vbTextCompare) > 0 Then
                tmp_str = tmp_str & "," & NumberRangeToList(CStr(arrTmp(i)))
            Else
                tmp_str = tmp_str & "," & CStr(arrTmp(i))
            End If
        End If
    Next i
    
    OptimiseScheduleString = Mid(tmp_str, 2)
    
    ' TODO:
    ' what else can be optimised:
    ' remove duplicates and sort array
    ' remove negative values
    ' however, nice to have, not must
Exit_Function:
    Exit Function
ErrHandler:
    OptimiseScheduleString = vbNullString
    GoTo Exit_Function
    Resume
End Function

Function NumberRangeToList(str As String) As String
' function accepts string with pattern [d..d]
' e.g. 1..5
' function will return: 1,2,3,4,5

    Dim arrValues
    Dim i As Long
    
    On Error GoTo ErrHandler
    
    arrValues = Split(str, "..")
    arrValues(0) = CInt(Trim(arrValues(0)))
    arrValues(1) = CInt(Trim(arrValues(1)))
    
    For i = arrValues(0) To arrValues(1)
        NumberRangeToList = NumberRangeToList & "," & i
    Next i
    NumberRangeToList = Mid(NumberRangeToList, 2)
    
    Exit Function
    
ErrHandler:
    ' if here - something went wrong
    NumberRangeToList = vbNullString
End Function

Function FindOffsetToNextScheduledMonth(StartingDate As Date, MonthsStringConverted As String) As Byte
' Function returns number of months that we have to skip to get next scheduled month.
' StartingDate - base for calculation
' MonthsStringConverted
'   Original string converted with PrepareMonthsString(). Represents comma separated list of allowed months.
' Function works for both working and non-working days scenarios.

    Dim arr
    Dim i As Byte
    Dim dMonth As Byte
    Dim minOffset As Byte
    Dim offset As Byte
    Dim minArrElement As Byte
    
    On Error GoTo ErrHandler
    
    minOffset = 25
    minArrElement = 25
    dMonth = Month(StartingDate)
    
    arr = Split(MonthsStringConverted, ",", , vbTextCompare)
    For i = LBound(arr) To UBound(arr)
        If minArrElement > CByte(arr(i)) Then
            minArrElement = CByte(arr(i))
        End If
        
        ' seek for month in future
        If CByte(arr(i)) > dMonth Then
            ' if have month scheduled within year - calc offset
            If minOffset > CByte(arr(i)) - dMonth Then
                minOffset = CByte(arr(i)) - dMonth
            End If
        End If
    Next i
    
    ' if couldn't find month within year of StartingDate
    If minOffset = 25 Then
        ' calc offset to min month in next year
        ' e.g. 3 in next year, currently 8
        ' 3 + (12-8)
        minOffset = minArrElement + 12 - dMonth
    End If
    
    FindOffsetToNextScheduledMonth = minOffset
    
Exit_Function:
    Exit Function
ErrHandler:
    ' if here - something went wrong
    FindOffsetToNextScheduledMonth = 0
End Function
