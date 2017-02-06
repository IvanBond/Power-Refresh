Attribute VB_Name = "Support_Log"
Option Explicit

Sub Write_Log(str As String, Optional Mandatory_Record As Boolean, Optional Child_ProcessID As Long)
    Dim LogFile As String
    On Error Resume Next
    Const ForAppending = 8
    'LogFile = LogsFolder & "Log_" & ReportName & ".txt"
    
    ' (1) option - one file per ReportID.
    ' row format
    ' now # Process ID # message [# execution time]
    
    ' (2) option - one file per Report ID / Process ID - generates huge amount of small files
    ' row format
    ' now # message [# execution time]
    ' Logging version depends on Emailing about error as log file is attached to email
    
    If Log_Enabled Or Mandatory_Record Then
        ' possible extension - additional param log_on_reportid_level
        ' if log on ReportID level - then just
         LogFile = LogsFolderPath & ReportID & ".log" ' (1) option
        '
        ' otherwise - on ReportID + ProcessID - for easier tracking result of execution in Reports Controller
        ' LogFile = LogsFolderPath & ReportID & "_" & ProcessID & ".log" ' (2) option
        
        If Not objFSO.FileExists(LogFile) Then objFSO.CreateTextFile (LogFile)
        Set objLog = objFSO.OpenTextFile(LogFile, ForAppending)
        
        If Child_ProcessID = 0 Then
            objLog.WriteLine (Format(Now(), "YYYY-MM-dd hh:mm:ss") & "# " & ProcessID & " # " & str) ' (1) option
        Else
            ' as this process write log instead of child - probably child is hanging
            objLog.WriteLine (Format(Now(), "YYYY-MM-dd hh:mm:ss") & "# " & Child_ProcessID & " # " & str)
        End If
        
        'objLog.WriteLine (Format(Now(), "YYYY-MM-dd hh:mm:ss") & "# " & str) ' (2) option
        objLog.Close
        Set objLog = Nothing
    End If
End Sub

Function Get_Log_Records_For_Process(strProcessID As String) As String
    Dim MyData
    Dim arrData()
    Dim arrTmp()
    Dim i As Long
    
    Open LogsFolderPath & ReportID & ".log" For Binary As #1
    MyData = Space$(LOF(1))
    Get #1, , MyData
    Close #1
    arrData() = Split(MyData, vbCrLf)
    
    For i = LBound(arrData) To UBound(arrData)
        If Len(Trim(arrData(i))) <> 0 Then
            arrTmp = Split(arrData(i), "#")
            If Trim(arrTmp(1)) = strProcessID Then
            ' maybe worth to check also date of record
                Get_Log_Records_For_Process = Get_Log_Records_For_Process & arrTmp(2) & vbCrLf
            End If
        End If
    Next i
End Function

Function Get_Log_Last_Record(strProcessID As String) As String
    Dim MyData
    Dim arrData()
    Dim arrTmp()
    Dim i As Long
    
    Open LogsFolderPath & ReportID & ".log" For Binary As #1
    MyData = Space$(LOF(1))
    Get #1, , MyData
    Close #1
    arrData() = Split(MyData, vbCrLf)
    
    For i = LBound(arrData) To UBound(arrData)
        If Len(Trim(arrData(i))) <> 0 Then
            arrTmp = Split(arrData(i), "#")
            If Trim(arrTmp(1)) = strProcessID Then
                Get_Log_Last_Record = arrData(i)
            End If
        End If
    Next i
End Function

