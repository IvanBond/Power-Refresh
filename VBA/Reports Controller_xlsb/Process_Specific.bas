Attribute VB_Name = "Process_Specific"
Option Explicit
Option Compare Text
'
' Module contains subroutines to work with WMI processes
' Scenarios
' 1. Controller checks if the process is still running to update status
'    also checks Execution Time Limit: if exceeded - kill process with dependent processes
'    need to find all processes with certain "Report ID *" in command line
' 2. User set flag to Terminate process: controller checks flag - if non empty - kill process with dependent processes
'    need to find all processes with certain "Report ID *" in command line
' 3. before start new Excel process need to check how many processes are already running
'    by checking process count limit in settings SETTINGS_PROCESS_COUNT_LIMIT
'

Sub CheckAndTerminateProcessesByReportID( _
            Report_Row_ID As Long, _
            Optional Timelimit As Double, _
            Optional Send_Email As Boolean)
' Function executes CheckAndTerminateRunningProcesses
' for the report provided. Reference to report is set by Row id on Control Panel sheet
' Used for Scenario 1 - main function
            
    ' check processes with Report ID in command_line
    Call CheckAndTerminateProcessesByCommandLineContains( _
        BuildReportIDstring(Report_Row_ID), _
        Timelimit, _
        Send_Email)
        
End Sub

Private Function BuildReportIDstring(Report_Row_ID As Long) As String
    ' check process is running
    If Control_Table Is Nothing Then
        Call Set_Global_Variables
    End If
    
    BuildReportIDstring = " /x /e/report_id:" & ControlPanel.Cells(Report_Row_ID, _
            Control_Table.ListColumns("Report ID *").Range.Column).Value
            
    ' Controller calls Refresher with encoded command line
    If Val(Application.Version) >= 15 Then
        BuildReportIDstring = WorksheetFunction.EncodeURL(BuildReportIDstring)
    Else
        ' EncodeURL is not available in prev versions
        BuildReportIDstring = Support_Functions.URLEncode(BuildReportIDstring)
    End If
End Function

Sub CheckAndTerminateProcessesByCommandLineContains( _
                    substr As String, _
                    Optional Timelimit As Double, _
                    Optional Send_Email As Boolean, _
                    Optional Report_Row_ID As Long)
' Function checks uptime of process
'
    Dim objWMIService As Object
    Dim colItems As Object
    Dim objItem As Object
    Dim StartTime As Date
    
    On Error GoTo ErrHandler
    
    Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
        
    Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_Process" & _
               " WHERE CommandLine like '%" & substr & "%'", , 48)
        
    For Each objItem In colItems
                        
        If Timelimit = 0 Then
        ' function called to kill processes
            Call KillProcessWithDependents(objItem.ProcessID)
        Else
        ' function called to check TimeLimit
            
            StartTime = DateValue(WMIDateStringToDateTime(objItem.CreationDate)) + _
                TimeValue(WMIDateStringToDateTime(objItem.CreationDate))

            If Timelimit <= Round((Now - StartTime) * 60 * 24, 1) Then
                ' when exceeded time limit - terminate process
                Call KillProcessWithDependents(objItem.ProcessID)
                
                If Not Report_Row_ID = 0 Then
                    Call SetReportParameter(Report_Row_ID, "Status", "TERMINATED")
                    
                    ' when called to check timelimit - have to send notification
                    Call SendNotification( _
                        "Power Refresh: Report '" & GetReportParameter(Report_Row_ID, "Report ID *") & "' - TIME EXCEEDED", _
                        "Power Refresh Failure Message", _
                        Report_Row_ID)
                    
                End If ' If Not Report_Row_ID = 0 Then
            Else
                ' still running
            End If ' If TimeLimit <= Round((Now - StartTime) * 60 * 24, 1) Then
        End If ' Timelimit = 0
    Next objItem

Exit_Sub:
    On Error Resume Next
    Set objWMIService = Nothing
    Set colItems = Nothing
    Set objItem = Nothing
    Err.Clear
    Exit Sub
    
ErrHandler:
    Debug.Print Now, "CheckAndTerminateProcessesByCommandLineContains", Err.Number & ": " & Err.Description
    Err.Clear
    GoTo Exit_Sub
    Resume
End Sub

Function GetProcessesIDByReportID(ReportID As String) As String
' Function returns comma separated list with Process IDs
' Converts ReportID provided into expected substring of CommandLine
' according to logic used in original function

    Dim objWMIService As Object
    Dim colItems As Object
    Dim objItem As Object
    Dim tmpstr As String
    
    On Error GoTo ErrHandler
    
    Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
    
    ' CommandLine was created by
    'Set objProc = objShell.Exec(Excel_Path & " /x " & _
                            "/e" & WorksheetFunction.encodeURL(Collect_Parameters(cell.Row)) & _
                            " /r """ & Refresher_Path & """")
    
    ' where first parameter is /report_id
    '    Collect_Parameters = "/report_id:" & Control_Table.Parent.Cells(Report_Row_ID, _
        Control_Table.ListColumns(Field_Name).Range.Column).Value

    ' so we look for process with command line that contains
    '   /x /e/report_id:[MYREPORTID]
    ' e.g. /x /e/report_id:Query Xrates
    ' however, we need to use EncodeURL before
    
    tmpstr = " /x /e/report_id:" & ReportID
    If Val(Application.Version) >= 15 Then
        tmpstr = WorksheetFunction.EncodeURL(tmpstr)
    Else
        ' EncodeURL is not available in prev versions
        tmpstr = Support_Functions.URLEncode(tmpstr)
    End If
        
    Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_Process" & _
               " WHERE CommandLine like '%" & tmpstr & "%'", , 48)
    
    tmpstr = vbNullString
    For Each objItem In colItems
        tmpstr = tmpstr & "," & objItem.ProcessID
    Next objItem
    
    If tmpstr <> vbNullString Then
        GetProcessesIDByReportID = Mid(tmpstr, 2) ' skip first comma
    End If
    
Exit_Function:
    On Error Resume Next
    Set objWMIService = Nothing
    Set colItems = Nothing
    Set objItem = Nothing
    Err.Clear
    Exit Function
    
ErrHandler:
    Debug.Print Now, "GetProcessIDsByReportID", Err.Number & ": " & Err.Description
    Err.Clear
    GetProcessesIDByReportID = vbNullString
    GoTo Exit_Function
    Resume ' for debug
End Function

Function GetOldestStartTime(strProcesses As String) As Double
' default return = DateValue("9999-12-31")
    Dim arr
    Dim i As Long
    Dim tmpDate As Date
    Dim ProcessStartTime As Date
    
    On Error GoTo ErrHandler
    
    tmpDate = DateValue("9999-12-31")
    arr = Split(Replace(strProcesses, " ", vbNullString), ",", , vbTextCompare)
    For i = LBound(arr) To UBound(arr)
        ProcessStartTime = GetProcessStartTime(CStr(arr(i)))
        If tmpDate > ProcessStartTime Then
            tmpDate = ProcessStartTime
        End If
    Next i
    GetOldestStartTime = tmpDate

Exit_Function:
    Exit Function
    
ErrHandler:
    Debug.Print Now, "GetOldestStartTime", Err.Number & ": " & Err.Description
    Err.Clear
    GoTo Exit_Function
    Resume ' for debug
End Function
