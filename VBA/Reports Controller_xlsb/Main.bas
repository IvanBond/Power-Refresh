Attribute VB_Name = "Main"
Option Explicit
Option Compare Text

'
' procedure is operated by Timer, but can be executed manually as well
Sub Check_And_Run()

    Dim cell As Range
    Dim Field_Name As String
    Dim objShell, objProc As Object
    Dim log_row As Long
    Dim StartTime As Double
    Dim arrScopes
    Dim i As Long
    
    Dim sh As Worksheet
    
    On Error GoTo ErrHandler
    
    ' skip cycle step if Application is in edit mode
    If IsEditing Then
        Debug.Print Now, "Check_And_Run", "Edit Mode detected, Skip cycle step"
        GoTo Exit_sub
    End If
    
    ' re-define each time, just in case
    Call Set_Global_Variables
    
    Set objShell = CreateObject("WScript.Shell")
    
    Call UpdateActivityTrackingFile ' update file which can be used to check that Report Controller is online from external app
        
    ' loop through rows
    For Each cell In Control_Table.ListColumns("Report ID *").DataBodyRange
        
        ' Terminate if flag "Terminate" is set
        Call IfTerminateProcess(cell.Row)
        
        ' Check reports that are 'In Process'
        Call IfInProcess(cell.Row)
        
        ' check all conditions for run
        ' Enabled or Next Run time is passed
        ' or if Manual Trigger
        If (GetReportParameter(cell.Row, "Enabled") = "Y" And _
            GetReportParameter(cell.Row, "Next Run") < Now()) Or _
                GetReportParameter(cell.Row, "Force Start") <> vbNullString Then
            
            ' just in case check row validity
            If Is_Row_Valid(cell.Row) Then
                
                ' if reached Excel processes limit - remember time when we reached the limit
                If IfReachedLimitOfExcelProcesses(cell.Row) Then
                    ' don't clear Manual Trigger / don't calc next run time...
                    ' just check other rows and wait...
                    GoTo Next_Task
                End If
                
                ' if reached Limit of Workstation Resources - remember time when we reached the limit
                If IfReachedLimitOfWorkstationResources(cell.Row) Then
                    ' don't clear Manual Trigger / don't calc next run time...
                    ' just check other rows and wait...
                    GoTo Next_Task
                End If

                Call SetReportParameter(cell.Row, "Last Run", Now)
                ' if we run task via putting Manual Trigger - we do not need to re-calc Next Run.
                ' in other words, we re-calc next run only then Next Run < Now()
                
                If GetReportParameter(cell.Row, "Next Run") < Now Then
                    Call SetReportParameter(cell.Row, "Next Run", GetScheduledRunTime(cell.Row))
                End If
                
                ' always clear Manual Trigger
                Call SetReportParameter(cell.Row, "Force Start", vbNullString)
                Call SetReportParameter(cell.Row, "Status", "In Process: 0:00")
                
                Set objProc = objShell.Exec(Excel_Path & " /x " & _
                            "/e" & Support_Functions.URLEncodeString(Collect_Parameters(cell.Row)) & _
                            " /r """ & Refresher_Path & """")
                
                ' URL encoding is important because otherwise no chance to pass values with spaces and special chars
                ' Excel execution through shell will trigger every value separated by space as new file to be opened
                
                ' can run without wait if no workbooks in
                ' C:\Users\<username>\AppData\Roaming\Microsoft\Excel\XLSTART
                
                Sleep 5000
                
                ' write to internal LOG table
                If Not IsEditing Then
                    If Application.CalculationState = xlDone Then
                        Call Write_Log(cell.Row, objProc.ProcessID)
                    End If
                End If
                
                Set objProc = Nothing
            Else
                ' row is not valid - function Is_Row_Valid puts necessary comment to Status field
            End If ' if row is valid
                        
        End If ' check if enabled

Next_Task:
    Next cell

Exit_sub:
    On Error Resume Next
    ThisWorkbook.Save
    Set objShell = Nothing
    Application.Interactive = True
    Application.ScreenUpdating = True
    Exit Sub
    
ErrHandler:
    Debug.Print Now, "Check_And_Run", Err.Number & ": " & Err.Description
    Err.Clear
    GoTo Exit_sub
    Resume
End Sub

Function GetReportParameter(Report_Row_ID As Long, strParameter As String)
    On Error Resume Next
    GetReportParameter = Control_Table.Parent.Cells(Report_Row_ID, _
        Control_Table.ListColumns(strParameter).Range.Column).Value
End Function

Function GetReportParameterByReportID(Report_ID As String, strParameter As String)
    On Error Resume Next
    Dim cell As Range
    
    For Each cell In Control_Table.ListColumns("Report ID *").DataBodyRange
        If cell.Value = Report_ID Then
            GetReportParameterByReportID = Control_Table.Parent.Cells(cell.Row, _
                Control_Table.ListColumns(strParameter).Range.Column).Value
            Exit For
        End If
    Next cell
    Set cell = Nothing
End Function


Sub SetReportParameter(Report_Row_ID As Long, strParameter As String, svalue)
    On Error Resume Next
    If Not IsEditing Then
        If Application.CalculationState = xlDone Then
            Control_Table.Parent.Cells(Report_Row_ID, _
                Control_Table.ListColumns(strParameter).Range.Column).Value = svalue
        End If
    End If
End Sub

Private Function IfReachedLimitOfExcelProcesses(Report_Row_ID As Long)
    On Error GoTo ErrHandler
    'Check number of running Excel processes
    If ReachedLimitOfExcelProcesses(Report_Row_ID) Then
        
        IfReachedLimitOfExcelProcesses = True
        
        ' count time of inability to start new process
        ' send notification after certain time
        If ReachedExcelProcessesLimitsTime = 0 Then
            ' if it is first time - just remember time
            ReachedExcelProcessesLimitsTime = Now
        Else
            ' check if more than limit time
            If Now - ReachedExcelProcessesLimitsTime > _
                    (Val(ThisWorkbook.Names("SETTINGS_MINUTES_CANT_START_EXCEL").RefersToRange.Value) / 24 / 60) Then
                ' if cannot start Excel for 15 min, for example - send notification
                Call SendNotification( _
                    "Power Refresh: Warning! Number of Excel processes exceeded limit", _
                    "Power Refresh Warning Message")
            End If
        End If
    Else
    ' reset counter
        ReachedExcelProcessesLimitsTime = 0
    End If

Exit_sub:
    Exit Function
    
ErrHandler:
    Debug.Print Now, "IfReachedLimitOfExcelProcesses", Err.Number & ": " & Err.Description
    Err.Clear
    GoTo Exit_sub
    Resume
End Function

Private Sub IfTerminateProcess(Report_Row_ID As Long)
    
    If GetReportParameter(Report_Row_ID, "Terminate Process") <> vbNullString Then
        ' call function that checks all Excel processes with "Report ID *" in command line
        ' and terminate them
        Call CheckAndTerminateProcessesByReportID(Report_Row_ID)
        ' no check of Timelimit, no Email
        
        ' Clear flag
        Call SetReportParameter(Report_Row_ID, "Terminate Process", vbNullString)
        ' Set Status
        Call SetReportParameter(Report_Row_ID, "Status", "TERMINATED by User")
    End If
    
End Sub

Sub tttt()
    Call IfInProcess(29)
End Sub

Sub IfInProcess(Report_Row_ID As Long)
    Dim strProcesses As String
    Dim StartTime As Date
    
    On Error GoTo ErrHandler
    ' if Status contains In Process
    If Left(GetReportParameter(Report_Row_ID, "Status"), 10) = "In Process" Then
        ' Update status if still in process
        
        strProcesses = GetProcessesIDByReportID( _
                            GetReportParameter(Report_Row_ID, _
                                                "Report ID *"))
        
        If strProcesses <> vbNullString Then
            StartTime = GetOldestStartTime(strProcesses)
            
            If StartTime <> DateValue("9999-12-31") Then
                If GetReportParameter(Report_Row_ID, "Execution Time Limit") <> vbNullString Then
                
                    Call CheckAndTerminateProcessesByReportID(Report_Row_ID, _
                            GetReportParameter(Report_Row_ID, "Execution Time Limit"), _
                            True)
                Else
                    ' if no Execution Time Limit
                    Call SetReportParameter(Report_Row_ID, "Status", "In Process: " & Format(Now() - StartTime, "hh:mm:ss"))
                End If
            End If ' If StartTime <> DateValue("9999-12-31") Then
        Else
        ' if process doesn't exist anymore
            Call SetReportParameter(Report_Row_ID, "Status", _
                    Replace(GetReportParameter(Report_Row_ID, "Status"), _
                            "In Process", "Completed") & "+")
        End If ' If strProcesses <> vbNullString Then
    End If ' if Status contains In Process

Exit_sub:
    Exit Sub
    
ErrHandler:
    Debug.Print Now, "IfInProcess", Err.Number & ": " & Err.Description
    Err.Clear
    GoTo Exit_sub
    Resume ' for debug
End Sub

Function Is_Row_Valid(Report_Row_ID As Long) As Boolean
    Dim Field_Name  As String
    
    ' ******************** check mandatory fields ********************
    
    Field_Name = "File or Folder Path *"
    If GetReportParameter(Report_Row_ID, Field_Name) = vbNullString Then
        Call SetReportParameter(Report_Row_ID, "Status", "Provide valid report path")
        Exit Function
    End If
    
    ' ******************** end of check mandatory fields ********************
    
    ' ******************** fields validation ********************
    Field_Name = "Save Inplace"
    If GetReportParameter(Report_Row_ID, Field_Name) = "Y" Then
        ' Save Inplace can't be used with some other parameters
        Field_Name = "Save Sheets"
        If GetReportParameter(Report_Row_ID, Field_Name) <> vbNullString Then
            Call SetReportParameter(Report_Row_ID, "Status", "'Save Inplace' can't be used with 'Save Sheets' as it will break your initial file. Consider saving a copy.")
            Exit Function
        End If
        
        Field_Name = "Delete Sheets"
        If GetReportParameter(Report_Row_ID, Field_Name) <> vbNullString Then
            Call SetReportParameter(Report_Row_ID, "Status", "'Save Inplace' can't be used with 'Delete Sheets' as it will break your initial file. Consider saving a copy.")
            Exit Function
        End If
        
        Field_Name = "Formulas to Values"
        If GetReportParameter(Report_Row_ID, Field_Name) <> vbNullString Then
            Call SetReportParameter(Report_Row_ID, "Status", "'Save Inplace' can't be used with 'Formulas to Values' as it will break your initial file. Consider saving a copy.")
            Exit Function
        End If
        
        Field_Name = "Delete WB Queries"
        If GetReportParameter(Report_Row_ID, Field_Name) <> vbNullString Then
            Call SetReportParameter(Report_Row_ID, "Status", "'Save Inplace' can't be used with 'Delete WB queries' as it will break your initial file. Consider saving a copy.")
            Exit Function
        End If
        
        Field_Name = "Do Not Save"
        If GetReportParameter(Report_Row_ID, Field_Name) <> vbNullString Then
            Call SetReportParameter(Report_Row_ID, "Status", "'Save Inplace' can't be used with 'Do Not Save'.")
            Exit Function
        End If
        
        Field_Name = "Add DateTime"
        If GetReportParameter(Report_Row_ID, Field_Name) <> vbNullString Then
            Call SetReportParameter(Report_Row_ID, "Status", "'Save Inplace' can't be used with 'Add DateTime'.")
            Exit Function
        End If
        
        Field_Name = "Parallel Refresh of Scopes"
        If GetReportParameter(Report_Row_ID, Field_Name) <> vbNullString Then
            Call SetReportParameter(Report_Row_ID, "Status", "'Save Inplace' can't be used with 'Parallel Refresh of Scopes'.")
            Exit Function
        End If
        
    End If
    ' ******************** end of fields validation ********************
    
    ' if passed
    Is_Row_Valid = True
End Function

Function Collect_Parameters(Report_Row_ID As Long, Optional Scope As String) As String
    Dim str As String
    Dim Field_Name As String
    
    ' Sample
    ' Collect_Parameters = "/debug_mode/log_enabled/file_path:C:\Temp\Test.xlsx/" & _
        "result_folder_path:\\server_name\ssis in\/macro_before:test/macro_after:after" & _
        "/error_email_to:test@domain.com/success_email_to:test@domain.com/result_filename:reportX/extension:xlsb" & _
        "/add_datetime/scopes:KZ HR,UA AMS/parameters:FROM=2016-01-01,TO=2016-10-31"
    
    Field_Name = "Report ID *"
    Collect_Parameters = "/report_id:" & Control_Table.Parent.Cells(Report_Row_ID, _
        Control_Table.ListColumns(Field_Name).Range.Column).Value
    
    Field_Name = "File or Folder Path *"
    Collect_Parameters = Collect_Parameters & "/target_path:" & _
        Replace(CheckPath(GetReportParameter(Report_Row_ID, Field_Name)), "/", "|")
    ' cheap solution - replace '/' before transfer parameters
                ' solves problem of web folder
    ' '|' cannot be used in path, so it will be replaced back in Refresher
    
'    Field_Name = "Type"
'        Collect_Parameters = Collect_Parameters & "/type:" & _
'            IIf(Control_Table.Parent.Cells(report_row_id, _
'                    Control_Table.ListColumns(Field_Name).Range.Column).Value = vbNullString, _
'                "R", _
'                Control_Table.Parent.Cells(report_row_id, _
'                    Control_Table.ListColumns(Field_Name).Range.Column).Value)
    
    Field_Name = "Macro Before"
    If Control_Table.Parent.Cells(Report_Row_ID, _
        Control_Table.ListColumns(Field_Name).Range.Column).Value <> vbNullString Then
        Collect_Parameters = Collect_Parameters & "/macro_before:" & Control_Table.Parent.Cells(Report_Row_ID, _
            Control_Table.ListColumns(Field_Name).Range.Column).Value
    End If
    
    Field_Name = "Skip RefreshAll"
    If Control_Table.Parent.Cells(Report_Row_ID, _
        Control_Table.ListColumns(Field_Name).Range.Column).Value <> vbNullString Then
        Collect_Parameters = Collect_Parameters & "/skip_refresh_all"
    End If
    
    Field_Name = "Macro After"
    If Control_Table.Parent.Cells(Report_Row_ID, _
        Control_Table.ListColumns(Field_Name).Range.Column).Value <> vbNullString Then
        Collect_Parameters = Collect_Parameters & "/macro_after:" & Control_Table.Parent.Cells(Report_Row_ID, _
            Control_Table.ListColumns(Field_Name).Range.Column).Value
    End If
    
    Field_Name = "Error Email To"
    If Control_Table.Parent.Cells(Report_Row_ID, _
        Control_Table.ListColumns(Field_Name).Range.Column).Value <> vbNullString Then
        Collect_Parameters = Collect_Parameters & "/error_email_to:" & Control_Table.Parent.Cells(Report_Row_ID, _
            Control_Table.ListColumns(Field_Name).Range.Column).Value
    End If
    
    Field_Name = "Success Email To"
    If Control_Table.Parent.Cells(Report_Row_ID, _
        Control_Table.ListColumns(Field_Name).Range.Column).Value <> vbNullString Then
        Collect_Parameters = Collect_Parameters & "/success_email_to:" & Control_Table.Parent.Cells(Report_Row_ID, _
            Control_Table.ListColumns(Field_Name).Range.Column).Value
    End If
        
    Field_Name = "Debug Mode"
    If Control_Table.Parent.Cells(Report_Row_ID, _
        Control_Table.ListColumns(Field_Name).Range.Column).Value = "Y" Then
        Collect_Parameters = Collect_Parameters & "/debug_mode"
    End If
    
    Field_Name = "Log Enabled"
    If Control_Table.Parent.Cells(Report_Row_ID, _
        Control_Table.ListColumns(Field_Name).Range.Column).Value = "Y" Then
        Collect_Parameters = Collect_Parameters & "/log_enabled"
    End If
    
    Field_Name = "Save Inplace"
    If Control_Table.Parent.Cells(Report_Row_ID, _
        Control_Table.ListColumns(Field_Name).Range.Column).Value = "Y" Then
        Collect_Parameters = Collect_Parameters & "/save_inplace"
    End If
    
    Field_Name = "Do Not Save"
    If Control_Table.Parent.Cells(Report_Row_ID, _
        Control_Table.ListColumns(Field_Name).Range.Column).Value = "Y" Then
        Collect_Parameters = Collect_Parameters & "/do_not_save"
    End If
    
    Field_Name = "Result Folder Path"
    If Control_Table.Parent.Cells(Report_Row_ID, _
        Control_Table.ListColumns(Field_Name).Range.Column).Value <> vbNullString Then
        Collect_Parameters = Collect_Parameters & "/result_folder_path:" & _
         Replace(Control_Table.Parent.Cells(Report_Row_ID, Control_Table.ListColumns(Field_Name).Range.Column).Value, _
                "/", "|") ' cheap solution - replace '/' before transfer parameters
                ' solves problem of web folder
                ' '|' cannot be used in path, so it will be replaced back in Refresher
    End If
    
    Field_Name = "Result FileName"
    If Control_Table.Parent.Cells(Report_Row_ID, _
        Control_Table.ListColumns(Field_Name).Range.Column).Value <> vbNullString Then
        Collect_Parameters = Collect_Parameters & "/result_filename:" & Control_Table.Parent.Cells(Report_Row_ID, _
            Control_Table.ListColumns(Field_Name).Range.Column).Value
    End If
    
    Field_Name = "Add Datetime"
    If Control_Table.Parent.Cells(Report_Row_ID, _
        Control_Table.ListColumns(Field_Name).Range.Column).Value = "Y" Then
        Collect_Parameters = Collect_Parameters & "/add_datetime"
    End If
    
    Field_Name = "Extension"
    If Control_Table.Parent.Cells(Report_Row_ID, _
        Control_Table.ListColumns(Field_Name).Range.Column).Value <> vbNullString Then
        Collect_Parameters = Collect_Parameters & "/extension:" & Control_Table.Parent.Cells(Report_Row_ID, _
            Control_Table.ListColumns(Field_Name).Range.Column).Value
    End If
    
    Field_Name = "Execution Time Limit"
    If Control_Table.Parent.Cells(Report_Row_ID, _
        Control_Table.ListColumns(Field_Name).Range.Column).Value <> vbNullString Then
        Collect_Parameters = Collect_Parameters & "/time_limit:" & Control_Table.Parent.Cells(Report_Row_ID, _
            Control_Table.ListColumns(Field_Name).Range.Column).Value
    End If
    
    Field_Name = "Parallel Refresh of Scopes"
    If Control_Table.Parent.Cells(Report_Row_ID, _
        Control_Table.ListColumns(Field_Name).Range.Column).Value <> vbNullString Then
        Collect_Parameters = Collect_Parameters & "/scopes_in_parallel"
    End If
    
    Field_Name = "Parallel Refresh of files in folder"
    If Control_Table.Parent.Cells(Report_Row_ID, _
        Control_Table.ListColumns(Field_Name).Range.Column).Value <> vbNullString Then
        Collect_Parameters = Collect_Parameters & "/files_in_parallel"
    End If
        
    ' will ignore if Save Inplace
    ' if save only one sheet
    Field_Name = "Save Sheet"
    If Control_Table.Parent.Cells(Report_Row_ID, _
        Control_Table.ListColumns(Field_Name).Range.Column).Value <> vbNullString Then
        Collect_Parameters = Collect_Parameters & "/save_sheet:" & Control_Table.Parent.Cells(Report_Row_ID, _
            Control_Table.ListColumns(Field_Name).Range.Column).Value
    End If

    ' will ignore if "Save Inplace"
    Field_Name = "Save Sheets"
    If Control_Table.Parent.Cells(Report_Row_ID, _
        Control_Table.ListColumns(Field_Name).Range.Column).Value <> vbNullString Then
        Collect_Parameters = Collect_Parameters & _
            "/save_sheets:" & Replace(Control_Table.Parent.Cells(Report_Row_ID, _
                Control_Table.ListColumns(Field_Name).Range.Column).Value, ";", ",")
    End If
    
    ' won't delete sheets if "Save Inplace"
    Field_Name = "Delete Sheets"
    If Control_Table.Parent.Cells(Report_Row_ID, _
        Control_Table.ListColumns(Field_Name).Range.Column).Value <> vbNullString Then
        Collect_Parameters = Collect_Parameters & _
            "/delete_sheets:" & Replace(Control_Table.Parent.Cells(Report_Row_ID, _
                Control_Table.ListColumns(Field_Name).Range.Column).Value, ";", ",")
    End If
    
    ' will ignore if "Save Inplace"
    Field_Name = "Formulas to Values"
    If Control_Table.Parent.Cells(Report_Row_ID, _
        Control_Table.ListColumns(Field_Name).Range.Column).Value <> vbNullString Then
        Collect_Parameters = Collect_Parameters & _
            "/formulas_to_values:" & Replace(Control_Table.Parent.Cells(Report_Row_ID, _
                Control_Table.ListColumns(Field_Name).Range.Column).Value, ";", ",")
    End If
    
    ' will ignore if "Save Inplace"
    Field_Name = "Delete WB queries"
    If Control_Table.Parent.Cells(Report_Row_ID, _
        Control_Table.ListColumns(Field_Name).Range.Column).Value <> vbNullString Then
        Collect_Parameters = Collect_Parameters & _
            "/delete_wb_queries:" & Replace(Control_Table.Parent.Cells(Report_Row_ID, _
                Control_Table.ListColumns(Field_Name).Range.Column).Value, ";", ",")
    End If
    
    Field_Name = "Parameters"
    If Control_Table.Parent.Cells(Report_Row_ID, _
        Control_Table.ListColumns(Field_Name).Range.Column).Value <> vbNullString Then
        Collect_Parameters = Collect_Parameters & "/parameters:" & _
            Replace(Control_Table.Parent.Cells(Report_Row_ID, Control_Table.ListColumns(Field_Name).Range.Column).Value, _
                    "/", "{|}")
            ' such replace is needed due to
            ' potential risk of appearing '/'
            ' this may lead to issue with resolving file/folder path
            ' changed back on Refresher side during params parsing
    End If
        
    Field_Name = "Error Email To"
    If Control_Table.Parent.Cells(Report_Row_ID, _
        Control_Table.ListColumns(Field_Name).Range.Column).Value <> vbNullString Then
        Collect_Parameters = Collect_Parameters & "/err_email_to:" & _
            Control_Table.Parent.Cells(Report_Row_ID, Control_Table.ListColumns(Field_Name).Range.Column).Value
    End If
    
    Field_Name = "Success Email To"
    If Control_Table.Parent.Cells(Report_Row_ID, _
        Control_Table.ListColumns(Field_Name).Range.Column).Value <> vbNullString Then
        Collect_Parameters = Collect_Parameters & "/succ_email_to:" & _
            Control_Table.Parent.Cells(Report_Row_ID, Control_Table.ListColumns(Field_Name).Range.Column).Value
    End If
    
    Field_Name = "Stop At Start"
    If Control_Table.Parent.Cells(Report_Row_ID, _
        Control_Table.ListColumns(Field_Name).Range.Column).Value <> vbNullString Then
        Collect_Parameters = Collect_Parameters & "/stop_at_start:" & _
            Control_Table.Parent.Cells(Report_Row_ID, Control_Table.ListColumns(Field_Name).Range.Column).Value
    End If
    
    Field_Name = "Stop on Macro Before"
    If Control_Table.Parent.Cells(Report_Row_ID, _
        Control_Table.ListColumns(Field_Name).Range.Column).Value <> vbNullString Then
        Collect_Parameters = Collect_Parameters & "/stop_on_macro_before:" & _
            Control_Table.Parent.Cells(Report_Row_ID, Control_Table.ListColumns(Field_Name).Range.Column).Value
    End If
    
    Field_Name = "Stop Before RefreshAll"
    If Control_Table.Parent.Cells(Report_Row_ID, _
        Control_Table.ListColumns(Field_Name).Range.Column).Value <> vbNullString Then
        Collect_Parameters = Collect_Parameters & "/stop_before_refresh_all:" & _
            Control_Table.Parent.Cells(Report_Row_ID, Control_Table.ListColumns(Field_Name).Range.Column).Value
    End If
    
    Field_Name = "Stop After RefreshAll"
    If Control_Table.Parent.Cells(Report_Row_ID, _
        Control_Table.ListColumns(Field_Name).Range.Column).Value <> vbNullString Then
        Collect_Parameters = Collect_Parameters & "/stop_after_refresh_all:" & _
            Control_Table.Parent.Cells(Report_Row_ID, Control_Table.ListColumns(Field_Name).Range.Column).Value
    End If
    
    Field_Name = "Stop on Macro After"
    If Control_Table.Parent.Cells(Report_Row_ID, _
        Control_Table.ListColumns(Field_Name).Range.Column).Value <> vbNullString Then
        Collect_Parameters = Collect_Parameters & "/stop_on_macro_after:" & _
            Control_Table.Parent.Cells(Report_Row_ID, Control_Table.ListColumns(Field_Name).Range.Column).Value
    End If
    
    Field_Name = "Stop Before Open WB"
    If Control_Table.Parent.Cells(Report_Row_ID, _
        Control_Table.ListColumns(Field_Name).Range.Column).Value <> vbNullString Then
        Collect_Parameters = Collect_Parameters & "/stop_before_open_wb:" & _
            Control_Table.Parent.Cells(Report_Row_ID, Control_Table.ListColumns(Field_Name).Range.Column).Value
    End If
    
    ' Scope can be passed as argument of current function
    If Scope = vbNullString Then
        Field_Name = "Scopes"
        If Control_Table.Parent.Cells(Report_Row_ID, _
            Control_Table.ListColumns(Field_Name).Range.Column).Value <> vbNullString Then
            
            Collect_Parameters = Collect_Parameters & "/scopes:" & _
            Replace(Control_Table.Parent.Cells(Report_Row_ID, Control_Table.ListColumns(Field_Name).Range.Column).Value, _
                "/", "{|}")
        End If
    Else
        Collect_Parameters = Collect_Parameters & "/scopes:" & Replace(Scope, "/", "{|}")
    End If
    
    ' Debug.Print Collect_Parameters
End Function

Function CheckSMTPSettings() As Boolean
    On Error Resume Next
    If [SETTINGS_SMTP_SERVER].Value = vbNullString Then: Exit Function
    If [SETTINGS_SMTP_FROM].Value = vbNullString Then: Exit Function
    
    If [SETTINGS_SMTP_AUTHENTICATION].Value = "Basic" Then
        If [SETTINGS_SMTP_USERNAME].Value = vbNullString Or _
            [SETTINGS_SMTP_PASSWORD].Value = vbNullString Then
            Exit Function
        End If
    End If
    
    If Err.Number = 0 Then
        CheckSMTPSettings = True
    End If
    Err.Clear
End Function

Function CheckPath(path As String) As String
' Function check if file exists
' then if folder exists
' adds "\" if there were no such char

    Dim objFSO As Object
    Dim sPath As String
    
    On Error Resume Next
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    sPath = path
    ' remove leading and trailing quotes
    If Left$(sPath, 1) = """" Then
        sPath = Mid(path, 2)
    End If
    If Right$(sPath, 1) = """" Then
        sPath = Left(sPath, Len(sPath) - 1)
    End If
    
    ' can't check existence of SharePoint file
    ' current verison supports only files on SharePoint
    If Left(sPath, 4) <> "http" Then
        If Not objFSO.FileExists(sPath) Then
        ' file doesn't exists
        ' check if it is a folder
            If Not objFSO.folderexists(sPath & IIf(Right$(sPath, 1) = "\", "", "\")) Then
            ' folder doesn't exists
                CheckPath = vbNullString
                GoTo Exit_Function
            Else
            ' folder exists
                CheckPath = sPath & IIf(Right$(sPath, 1) = "\", "", "\")
                GoTo Exit_Function
            End If
        Else
        ' file exists
            CheckPath = sPath
        End If
    Else
    ' SharePoint file / folder
    ' TODO - find a way to check SharePoint files and folders
    ' by default - consider it is a path to file
        If InStr(sPath, "?") > 0 Then
        ' we don't need parameters
            CheckPath = Left(sPath, InStr(sPath, "?") - 1)
        Else
            CheckPath = sPath
        End If
    End If

Exit_Function:
    Set objFSO = Nothing
    Err.Clear
End Function

Function ReachedLimitOfExcelProcesses(Report_Row_ID As Long) As Boolean
    Dim ScopesCount As Long
    
    On Error GoTo ErrHandler
    
    ' if no parameter SETTINGS_PROCESS_COUNT_LIMIT provided - no limits
    If ThisWorkbook.Names("SETTINGS_PROCESS_COUNT_LIMIT").RefersToRange.Value <> vbNullString Then
    ' otherwise - estimate number of Excel copies required
                
        If GetReportParameter(Report_Row_ID, "Parallel Refresh of Scopes") = "Y" Then
            
            ScopesCount = UBound(Split(GetReportParameter(Report_Row_ID, "Scopes"), ",", , vbTextCompare)) + 1
            
            ReachedLimitOfExcelProcesses = (GetRunningProcessesCountByName("excel.exe") + ScopesCount > _
                                            Val(ThisWorkbook.Names("SETTINGS_PROCESS_COUNT_LIMIT").RefersToRange.Value))
        Else

            ReachedLimitOfExcelProcesses = (GetRunningProcessesCountByName("excel.exe") + 1 > _
                                            Val(ThisWorkbook.Names("SETTINGS_PROCESS_COUNT_LIMIT").RefersToRange.Value))
                    
        End If ' If GetReportParameter(Report_Row_ID, "Parallel Refresh of Scopes") = "Y" Then
    End If ' ThisWorkbook.Names("SETTINGS_PROCESS_COUNT_LIMIT").RefersToRange.Value <> vbNullString Then
    
ErrHandler:
    Err.Clear
End Function

Private Function IfReachedLimitOfWorkstationResources(Report_Row_ID As Long)
    On Error GoTo ErrHandler
    'Check number of running Excel processes
    If ReachedLimitOfWorkstationResources(Report_Row_ID) Then
        
        IfReachedLimitOfWorkstationResources = True
        
        ' count time of inability to start new process
        ' send notification after certain time
        If ReachedExcelProcessesLimitsTime = 0 Then
            ' if it is first time - just remember time
            ReachedExcelProcessesLimitsTime = Now
        Else
            ' check if more than limit time
            If Now - ReachedExcelProcessesLimitsTime > _
                    (Val(ThisWorkbook.Names("SETTINGS_MINUTES_CANT_START_EXCEL").RefersToRange.Value) / 24 / 60) Then
                ' if cannot start Excel for [%parameter%] minutes - send notification
                Call SendNotification( _
                    "Power Refresh: Warning! Workstation Resources Limits has been reached", _
                    "Power Refresh Warning Message")
            End If
        End If
    Else
    ' reset counter
        ReachedExcelProcessesLimitsTime = 0
    End If

Exit_sub:
    Exit Function
    
ErrHandler:
    Debug.Print Now, "IfReachedLimitOfWorkstationResources", Err.Number & ": " & Err.Description
    Err.Clear
    GoTo Exit_sub
    Resume
End Function

Function ReachedLimitOfWorkstationResources(Report_Row_ID As Long) As Boolean
    ' we need to check all running Excel proceses
    ' take those that contain specific CommandLine generated by Reports Controller
    ' get ReportID from there
    ' for each ReportID - get parameter 'Required Resources (points)'
    Dim tmpstr As String
    Dim arr
    Dim i As Integer
    Dim UsedResources As Double
    Dim ReportRequiredResources As Double
    Dim WorkstationResourcesLimit As Double
    
    On Error GoTo ErrHandler
    
    tmpstr = GetListOfRunningReports
    arr = Split(tmpstr, ",", , vbTextCompare)
    
    ' loop through the list of Report IDs
    For i = LBound(arr) To UBound(arr)
        ReportRequiredResources = Val(GetReportParameterByReportID(CStr(arr(i)), "Required Resources (points)"))
        UsedResources = UsedResources + ReportRequiredResources
    Next i
    
    ' Current Report Required resources
    ReportRequiredResources = GetReportParameter(Report_Row_ID, "Required Resources (points)")
    
    ' check Settings - if no values - means no limit
    If ThisWorkbook.Names("SETTINGS_WORKSTATION_RESOURCES_LIMIT").RefersToRange.Value <> vbNullString Then
        WorkstationResourcesLimit = Val(ThisWorkbook.Names("SETTINGS_WORKSTATION_RESOURCES_LIMIT").RefersToRange.Value)
        
        ReachedLimitOfWorkstationResources = (UsedResources + ReportRequiredResources > WorkstationResourcesLimit)
        ' if more - then can't start Report_Row_ID-report
    End If

ErrHandler:
    Err.Clear
End Function


' OBSOLETE

'Sub Check_And_Run()
'
'    Dim cell As Range
'    Dim Field_Name As String
'    Dim objShell, objProc As Object
'    Dim log_row As Long
'    Dim StartTime As Double
'    Dim arrScopes
'    Dim i As Long
'
'    Dim sh As Worksheet
'
'    ' skip cycle step if Application is in edit mode
'    If IsEditing Then
'        Debug.Print Now() & ": " & "Edit Mode detected, Skip cycle step"
'        GoTo Exit_Sub
'    End If
'
'    Call Set_Global_Variables
'    Set objShell = CreateObject("WScript.Shell")
'
'    Call UpdateActivityTrackingFile ' update file which can be used to check that Report Controller is online from external app
'
'    ' build parameters string
'    ' loop through rows
'    For Each cell In Control_Table.ListColumns("Report ID *").DataBodyRange
'        ' Terminate is requested
'        If Control_Table.Parent.Cells(cell.Row, Control_Table.ListColumns("Terminate Process").Range.Column).Value <> vbNullString Then
'            If CheckProcessExist(Get_Last_Log_Record(cell.Row, "Process ID")) Then
'                ' TODO: find all processes with that Report ID, not only last one
'                ' this may happen if Scopes work in parallel
'                Call KillProcessWithDependents(Get_Last_Log_Record(cell.Row, "Process ID"))
'
'                If Not IsEditing Then
'                    If Application.CalculationState = xlDone Then
'                        Control_Table.Parent.Cells(cell.Row, _
'                            Control_Table.ListColumns("Status").Range.Column).Value = "TERMINATED"
'                    End If
'                End If
'                ' no need to send email when user requested to terminate process
'
'            End If
'
'            ' clear flag
'            If Not IsEditing Then
'                If Application.CalculationState = xlDone Then
'                    Control_Table.Parent.Cells(cell.Row, _
'                        Control_Table.ListColumns("Terminate Process").Range.Column).Value = vbNullString
'                End If
'            End If
'        End If
'        ' end of termination check
'
'        ' if Status contains In Process
'        If Left(Control_Table.Parent.Cells(cell.Row, Control_Table.ListColumns("Status").Range.Column).Value, 10) = "In Process" Then
'            ' Update status if still in process
'            ' check if process exists
'            ' TODO replace with another function
'            ' check parameters of that process (command line)
'
'            If CheckProcessExist(Get_Last_Log_Record(cell.Row, "Process ID")) Then
'                StartTime = Get_Last_Log_Record(cell.Row, "Start Time") ' get process start time
'                If StartTime <> -1 Then
'                    ' check if execution takes longer time than expected
'                    If Control_Table.Parent.Cells(cell.Row, _
'                                Control_Table.ListColumns("Execution Time Limit").Range.Column).Value <> vbNullString Then
'
'                        If (Now() - StartTime) * 24 * 60 >= Control_Table.Parent.Cells(cell.Row, _
'                                Control_Table.ListColumns("Execution Time Limit").Range.Column).Value Then
'                            ' kill process with its children
'                            ' User has to provide 'Time Limit' including total time of execution of all possible dependent tasks
'                            Call KillProcessWithDependents(Get_Last_Log_Record(cell.Row, "Process ID"))
'
'                            If Not IsEditing Then
'                                If Application.CalculationState = xlDone Then
'                                    Control_Table.Parent.Cells(cell.Row, _
'                                        Control_Table.ListColumns("Status").Range.Column).Value = "TERMINATED"
'                                End If
'                            End If
'
'                            If [SETTINGS_EMAIL_ERRORS_TO].Value <> vbNullString Then
'
'                                If [SETTINGS_EMAIL_METHOD].Value = "Outlook" Then
'
'                                    Call Send_Email_Outlook([SETTINGS_EMAIL_ERRORS_TO].Value, _
'                                        "Report '" & cell.Value & "' - TIME EXCEEDED", _
'                                        "Failure Message", _
'                                        IIf([SETTINGS_EMAIL_ATTACH_LOGFILE].Value = "Y", ThisWorkbook.path & "\Logs\" & cell.Value & ".log", vbNullString), _
'                                        IIf([SETTINGS_EMAIL_IMPORTANCE].Value <> vbNullString, [SETTINGS_EMAIL_IMPORTANCE].Value, "Normal"))
'
'                                ElseIf [SETTINGS_EMAIL_METHOD].Value = "SMTP" Then
'
'                                    If CheckSMTPSettings Then
'                                        Call Send_EMail_CDO([SETTINGS_SMTP_FROM].Value, _
'                                            [SETTINGS_EMAIL_ERRORS_TO].Value, _
'                                            "Report '" & cell.Value & "' - TIME EXCEEDED", _
'                                            "Failure Message", _
'                                            IIf([SETTINGS_EMAIL_ATTACH_LOGFILE].Value = "Y", ThisWorkbook.path & "\Logs\" & cell.Value & ".log", vbNullString), _
'                                            IIf([SETTINGS_EMAIL_IMPORTANCE].Value <> vbNullString, [SETTINGS_EMAIL_IMPORTANCE].Value, "Normal"))
'                                    End If
'                                End If
'
'                            End If ' SETTINGS_EMAIL_ERRORS_TO
'
'                        End If
'                    Else
'
'                        If Not IsEditing Then
'                            If Application.CalculationState = xlDone Then
'                                Control_Table.Parent.Cells(cell.Row, _
'                                    Control_Table.ListColumns("Status").Range.Column).Value = "In Process: " & Format(Now() - StartTime, "hh:mm:ss")
'                            End If
'                        End If
'                    End If
'                End If
'            Else
'            ' if process doesn't exist anymore
'                If Not IsEditing Then
'                    If Application.CalculationState = xlDone Then
'                        Control_Table.Parent.Cells(cell.Row, Control_Table.ListColumns("Status").Range.Column).Value = Replace( _
'                            Control_Table.Parent.Cells(cell.Row, Control_Table.ListColumns("Status").Range.Column).Value, _
'                                "In Process", "Completed", compare:=vbTextCompare) & "+"
'                    End If
'                End If
'            End If
'        End If ' if Status contains In Process
'
'        ' check all conditions for run
'        ' Enabled or Next Run time is passed
'        ' or if Manual Trigger
'        If (Control_Table.Parent.Cells(cell.Row, Control_Table.ListColumns("Enabled").Range.Column).Value = "Y" And _
'            Control_Table.Parent.Cells(cell.Row, Control_Table.ListColumns("Next Run").Range.Column).Value < Now()) Or _
'                Control_Table.Parent.Cells(cell.Row, Control_Table.ListColumns("Manual Trigger").Range.Column).Value <> vbNullString Then
'
'            ' just in case check row validity
'            If Is_Row_Valid(cell.Row) Then
'                If Not IsEditing Then
'                    If Application.CalculationState = xlDone Then
'                        Control_Table.Parent.Cells(cell.Row, Control_Table.ListColumns("Last Run").Range.Column).Value = Now()
'                    End If
'                End If
'                ' if we run task that is planned on e.g. next week via putting Manual Trigger - we do not need to re-calc Next Run
'                ' in other words, we re-calc next run only then Next Run < Now()
'
'                If Control_Table.Parent.Cells(cell.Row, Control_Table.ListColumns("Next Run").Range.Column).Value < Now() Then
'                    If Not IsEditing Then
'                        If Application.CalculationState = xlDone Then
'                            ' Calculate Next Run
'                             Control_Table.Parent.Cells(cell.Row, _
'                                Control_Table.ListColumns("Next Run").Range.Column).Value = Get_Next_Run_DateTime(cell.Row)
'                        End If
'                    End If
'                     ' Debug.Print Get_Next_Run_DateTime(cell.Row)
'
'                End If
'
'                If Not IsEditing Then
'                    If Application.CalculationState = xlDone Then
'                        ' remove manual trigger - every time
'                        Control_Table.Parent.Cells(cell.Row, _
'                            Control_Table.ListColumns("Manual Trigger").Range.Column).Value = vbNullString
'
'                        ' clear Status - as this code is executed only when we start new instance
'                        Control_Table.Parent.Cells(cell.Row, _
'                            Control_Table.ListColumns("Status").Range.Column).Value = "In Process: 0:00"
'                    End If
'                End If
'
'                ' therefore - /r is last parameter
'                ' order is important for Parsing macro in Refresher.xlsb !!!
'
'                If Val(Application.Version) >= 15 Then
'                    Set objProc = objShell.Exec(Excel_Path & " /x " & _
'                            "/e" & WorksheetFunction.EncodeURL(Collect_Parameters(cell.Row)) & _
'                            " /r """ & Refresher_Path & """")
'                Else
'                    ' EncodeURL is not available in prev versions
'                    Set objProc = objShell.Exec(Excel_Path & " /x " & _
'                            "/e" & Support_Functions.URLEncode(Collect_Parameters(cell.Row)) & _
'                            " /r """ & Refresher_Path & """")
'                End If
'
'                ' URL encoding is important because otherwise no chance to pass values with spaces and special chars
'                ' Excel execution through shell will trigger every value separated by space as new file to be opened
'
'                ' can run without wait if no workbooks in
'                ' C:\Users\<username>\AppData\Roaming\Microsoft\Excel\XLSTART
'
'                Sleep 5000
'
'                ' write to LOG table
'                If Not IsEditing Then
'                    If Application.CalculationState = xlDone Then
'                        Call Write_Log(cell.Row, objProc.ProcessID)
'                    End If
'                End If
'
'                Set objProc = Nothing
'            Else
'                ' row is not valid - function Is_Row_Valid put necessary comment to Status field
'            End If ' if row is valid
'
'            'Debug.Print cell.Row; Control_Table.Parent.Cells(cell.Row, Control_Table.ListColumns("Enabled").Range.Column).Value
'        End If ' check if enabled
'
'Next_Cell:
'    Next cell
'
'Exit_Sub:
'    On Error Resume Next
'    ThisWorkbook.Save
'    Set objShell = Nothing
'    Application.Interactive = True
'    Application.ScreenUpdating = True
'    Exit Sub
'
'ErrHandler:
'    Debug.Print Now() & ": " & "Check And Run: " & Err.Number & ": " & Err.Description
'    Err.Clear
'    GoTo Exit_Sub
'    Resume
'End Sub

