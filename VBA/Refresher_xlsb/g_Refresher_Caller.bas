Attribute VB_Name = "g_Refresher_Caller"
Option Explicit
Option Compare Text

Sub Refresher_Caller(Optional Scope As String, Optional Target_File As String)
    ' sub is called when we need to refresh multiple scopes or files
    ' scopes one-by-one
    ' files one-by-one
    ' in case of error during execution of this sub
    ' ...
    Dim objShell, objProc As Object
    Dim waitOnReturn As Boolean: waitOnReturn = True
    Dim windowStyle As Integer: windowStyle = 1
    Dim tmp_process_id As Long
    Dim BeforeAction
    
    ' about WshShell Run method
    'https://msdn.microsoft.com/en-us/library/d5fk67ky(v=vs.85).aspx
    
    On Error GoTo ErrHandler ' in case of error with creation of objects
    Set objShell = CreateObject("WScript.Shell")
        
    ' for subsequent execution use Run command as it can wait until target application is closed
    ' http://stackoverflow.com/questions/15951837/wait-for-shell-command-to-complete
    'objShell.Run """" & Excel_Path & """ /x " & _
                        "/e" & objScriptEngine.Run("encodeURIComponent", Collect_Parameters(Scope, Target_File)) & _
                        " /r """ & Refresher_Path & """", windowStyle, waitOnReturn
    
    
    ' ALTERNATIVE OPTION - preferrable
    ' alternative option - use .Exec and then wait in loop and check process Status
    ' http://stackoverflow.com/questions/32127687/excel-vba-wait-for-shell-to-finish-before-continuing-with-script
    ' basics of exec method https://msdn.microsoft.com/en-us/library/ateytk4a(v=vs.85).aspx
    
    BeforeAction = Now()
    
    Call Write_Log("Creating new instance of Refresher... " & Target_File & Scope)
    
    On Error Resume Next ' catch error if cannot create object
    
    If Val(Application.Version) >= 15 Then
        Set objProc = objShell.Exec(Excel_Path & " /x " & _
                        "/e" & WorksheetFunction.EncodeURL(Collect_Parameters(Scope, Target_File)) & _
                        " /r """ & Refresher_Path & """")
    Else
        ' Excel 2013 and above have function WorksheetFunction.EncodeURL
        Set objProc = objShell.Exec(Excel_Path & " /x " & _
                        "/e" & Support_Functions.URLEncode(Collect_Parameters(Scope, Target_File)) & _
                        " /r """ & Refresher_Path & """")
    End If
    
    tmp_process_id = objProc.ProcessID
    If Err.Number <> 0 Then
        Call Write_Log("Couldn't create process", bMandatoryLogRecord)
        ' try again after some time?
        GoTo ErrHandler
    End If
    On Error GoTo 0
    
    Call Write_Log("New instance of Refresher has been created, process id: " & tmp_process_id & " : " & Scope & Target_File)
    
    If (Scope <> vbNullString And ThisWorkbook.Names("SETTINGS_SCOPES_IN_PARALLEL").RefersToRange.Value = "Y") Or _
        (Target_File <> vbNullString And ThisWorkbook.Names("SETTINGS_FILES_IN_PARALLEL").RefersToRange.Value = "Y") Then

        ' if parallel execution
        ' if Debug_Mode
        If ThisWorkbook.Names("SETTINGS_DEBUG_MODE").RefersToRange.Value = "Y" Then
            ' add child process to table
            Call Add_Child_Process(tmp_process_id)
        End If
        
        ' add new process object to array
        Child_Counter = Child_Counter + 1
        ReDim Preserve arrChild_Processes(3, Child_Counter)
        arrChild_Processes(0, Child_Counter) = tmp_process_id
        arrChild_Processes(1, Child_Counter) = Scope & Target_File ' always only one is not empty
        arrChild_Processes(2, Child_Counter) = BeforeAction
        Set arrChild_Processes(3, Child_Counter) = objProc
    Else
        ' if not in parallel XOR file or scope
        ' wait until process is finished

        On Error GoTo ErrHandler
        Do While objProc.Status = 0 ' Running
            Application.Wait (Now() + TimeValue("00:00:10")) ' 10 seconds
            ' Support_Functions.WaitSeconds ?
            
            ' Process handles errors during execution itself
            
            ' in addition - we can check it from outside to prevent it running infinitely
            If Round(Round((Now() - BeforeAction) * 86400, 0) / 60, 0) > _
                IIf(Scope <> vbNullString, Time_Limit_Per_Scope, Time_Limit_Per_File) Then
                ' TOThnk: check log of target child process
                ' if process had record 'waiting for new try' (since BeforeAction time)
                ' extent wait limit on [count of 'waits of new try'] * [waiting time]
                ' low prio - just use initially large Time Limit
                If InStr(Get_Log_Last_Record(CStr(ProcessID)), "Waiting") > 0 Then
                    If Scope <> vbNullString Then
                        Time_Limit_Per_Scope = Time_Limit_Per_Scope + Delay_Between_Tries
                    Else
                        Time_Limit_Per_File = Time_Limit_Per_File + Delay_Between_Tries
                    End If
                
                Else
                    
                    Call Write_Log("Time limit exceeded for process " & tmp_process_id, bMandatoryLogRecord)
                    
                    ' Send Email
                    Call Send_Mail(False, "Time limit exceeded for process " & tmp_process_id _
                        & vbCrLf & Scope & Target_File & vbCrLf & vbCrLf & Collect_Parameters, Scope)
                    
                    ' terminate process
                    ' do not raise error - continue with rest of files / scopes
                    ' if it was file with set of scopes - we also should terminate its children processes
                    If Target_File <> vbNullString Then
                        Stop
                        ' File non-responsive itself or is waiting for child processes
                        'TODO: procedure that kills process with its children
                        
                    Else
                        ' if child process is just a scope - simply kill it
                        objProc.Terminate
                    End If
    
                    ' if time limit is much larger than needed time - it is valid to say that process died for some reason
                    ' probably worth to try it again... it depends
                    ' self-restarter could fail to restart
                End If
            End If
        Loop
        
    End If
                                                                  
                                  
    Call Write_Log("Process " & tmp_process_id & " is completed # " & Round((Now() - BeforeAction) * 86400, 0) & "s")
    
Exit_Sub:
    On Error Resume Next
    If (Scope <> vbNullString And ThisWorkbook.Names("SETTINGS_SCOPES_IN_PARALLEL").RefersToRange.Value = "Y") Or _
        (Target_File <> vbNullString And ThisWorkbook.Names("SETTINGS_FILES_IN_PARALLEL").RefersToRange.Value = "Y") Then
        ' parallel - just skip
    Else
        ' sequential execution - if process still exists - destroy it
        If Not objProc Is Nothing Then
            objProc.Terminate
        End If
    End If
    
    Set objProc = Nothing
    Set objShell = Nothing
    
    Exit Sub

ErrHandler:
    If Err.Number <> 0 Then
        Call Write_Log("Refresher_Caller: " & Err.Number & ": " & Err.Description, bMandatoryLogRecord)
    End If
    GoTo Exit_Sub
    Resume ' for test purpose
End Sub

Function Collect_Parameters(Optional Scope As String, Optional Target_File As String) As String
    ' Function is relevant only in case when we need to create new instance of Refresher
    ' this happens when
    ' (1) parallel / subsequent refresh for list of scopes - then we call this Function with particular scope
    ' (2) refresh folder - each file - new instance of refresher (same list of scopes came from Reports Controller)
    '    for this option we should copy list of Scopes from range Settings_Scopes
    
    Dim str As String
    Dim Field_Name As String
    
    ' Collect_Parameters = "/debug_mode/log_enabled/file_path:C:\Temp\Test.xlsx/" & _
        "result_folder_path:\\server_name\ssis in\/macro_before:test/macro_after:after" & _
        "/error_email_to:test@domain.com/success_email_to:test@domain.com/result_filename:reportX/extension:xlsb" & _
        "/add_datetime/scopes:KZ HR,UA AMS/parameters:FROM=2016-01-01,TO=2016-10-31"
        
    Collect_Parameters = "/report_id:" & ThisWorkbook.Names("SETTINGS_REPORT_ID").RefersToRange.Value
    
    If ThisWorkbook.Names("SETTINGS_DEBUG_MODE").RefersToRange.Value = "Y" Then
        Collect_Parameters = Collect_Parameters & "/debug_mode"
    End If
    
    If ThisWorkbook.Names("SETTINGS_LOG_ENABLED").RefersToRange.Value = "Y" Then
        Collect_Parameters = Collect_Parameters & "/log_enabled"
    End If
    
    'If ThisWorkbook.Names("SETTINGS_SESSION_TYPE").RefersToRange.Value <> vbNullString Then
    '    Collect_Parameters = Collect_Parameters & "/type:" & ThisWorkbook.Names("SETTINGS_SESSION_TYPE").RefersToRange.Value
    'Else
    '    Collect_Parameters = Collect_Parameters & "/type:R"
    'End If
        
    ' target path is mandatory parameter
    If Target_File = vbNullString Then
        ' scenario of file refresh - just pass the same file
        Collect_Parameters = Collect_Parameters & "/target_path:" & Replace(ThisWorkbook.Names("SETTINGS_TARGET_PATH").RefersToRange.Value, "/", "|")
    Else
        ' scenario of folder refresh - call this function to create new instance of Refresher to refresh particular file
        Collect_Parameters = Collect_Parameters & "/target_path:" & Replace(Target_File, "/", "|")
    End If
    
    If ThisWorkbook.Names("SETTINGS_MACRO_BEFORE").RefersToRange.Value <> vbNullString Then
        Collect_Parameters = Collect_Parameters & "/macro_before:" & ThisWorkbook.Names("SETTINGS_MACRO_BEFORE").RefersToRange.Value
    End If
    
    If ThisWorkbook.Names("SETTINGS_SKIP_REFRESH_ALL").RefersToRange.Value <> vbNullString Then
        Collect_Parameters = Collect_Parameters & "/skip_refresh_all"
    End If
    
    If ThisWorkbook.Names("SETTINGS_MACRO_AFTER").RefersToRange.Value <> vbNullString Then
        Collect_Parameters = Collect_Parameters & "/macro_after:" & ThisWorkbook.Names("SETTINGS_MACRO_AFTER").RefersToRange.Value
    End If
    
    If ThisWorkbook.Names("SETTINGS_ERROR_EMAIL_TO").RefersToRange.Value <> vbNullString Then
        Collect_Parameters = Collect_Parameters & "/error_email_to:" & ThisWorkbook.Names("SETTINGS_ERROR_EMAIL_TO").RefersToRange.Value
    End If
    
    If ThisWorkbook.Names("SETTINGS_SUCCESS_EMAIL_TO").RefersToRange.Value <> vbNullString Then
        Collect_Parameters = Collect_Parameters & "/success_email_to:" & ThisWorkbook.Names("SETTINGS_SUCCESS_EMAIL_TO").RefersToRange.Value
    End If
    
    If ThisWorkbook.Names("SETTINGS_DO_NOT_SAVE").RefersToRange.Value <> vbNullString Then
        Collect_Parameters = Collect_Parameters & "/do_not_save"
    End If
    
    If ThisWorkbook.Names("SETTINGS_SAVE_INPLACE").RefersToRange.Value <> vbNullString Then
        Collect_Parameters = Collect_Parameters & "/save_inplace"
    End If
    
    If ThisWorkbook.Names("SETTINGS_RESULT_FOLDER_PATH").RefersToRange.Value <> vbNullString Then
        Collect_Parameters = Collect_Parameters & "/result_folder_path:" & _
            Replace(ThisWorkbook.Names("SETTINGS_RESULT_FOLDER_PATH").RefersToRange.Value, "/", "|")
    End If
    
    If ThisWorkbook.Names("SETTINGS_RESULT_FILENAME").RefersToRange.Value <> vbNullString Then
        Collect_Parameters = Collect_Parameters & "/result_filename:" & ThisWorkbook.Names("SETTINGS_RESULT_FILENAME").RefersToRange.Value
    End If
    
    If ThisWorkbook.Names("SETTINGS_RESULT_FILE_EXTENSION").RefersToRange.Value <> vbNullString Then
        Collect_Parameters = Collect_Parameters & "/extension:" & ThisWorkbook.Names("SETTINGS_RESULT_FILE_EXTENSION").RefersToRange.Value
    End If
    
    If ThisWorkbook.Names("SETTINGS_ADD_DATETIME").RefersToRange.Value <> vbNullString Then
        Collect_Parameters = Collect_Parameters & "/add_datetime"
    End If
    
    ' when argument Scope is not null - we must pass it to new instance of Refresher
    If Scope = vbNullString Then
        If ThisWorkbook.Names("SETTINGS_SCOPES").RefersToRange.Value <> vbNullString Then
            Collect_Parameters = Collect_Parameters & "/scopes:" & Replace(ThisWorkbook.Names("SETTINGS_SCOPES").RefersToRange.Value, "/", "{|}")
        End If
    Else
        Collect_Parameters = Collect_Parameters & "/scopes:" & Replace(Trim(Scope), "/", "{|}")
    End If
    
    If ThisWorkbook.Names("SETTINGS_PARAMETERS").RefersToRange.Value <> vbNullString Then
        Collect_Parameters = Collect_Parameters & "/parameters:" & _
            Replace(ThisWorkbook.Names("SETTINGS_PARAMETERS").RefersToRange.Value, "/", "{|}")
    End If
    
    ' if we create new instance of Refresher for file - we have to pass calculated Time Limit per File
    If ThisWorkbook.Names("SETTINGS_TIME_LIMIT").RefersToRange.Value <> vbNullString Then
        If Target_File <> vbNullString Then
            Collect_Parameters = Collect_Parameters & "/time_limit:" & Time_Limit_Per_File
        Else
            Collect_Parameters = Collect_Parameters & "/time_limit:" & ThisWorkbook.Names("SETTINGS_TIME_LIMIT").RefersToRange.Value
        End If
    End If
        
    If ThisWorkbook.Names("SETTINGS_RESULT_SHEET_NAME").RefersToRange.Value <> vbNullString Then
        Collect_Parameters = Collect_Parameters & "/save_sheet:" & ThisWorkbook.Names("SETTINGS_RESULT_SHEET_NAME").RefersToRange.Value
    End If
    
'    Debug.Print Collect_Parameters
    
End Function

Sub Add_Child_Process(proc_id As Long)
    Dim rRow As Long
    ' TODO: re-write with ListRows.Add method
    
    Application.AutoCorrect.AutoExpandListRange = True
    
    With Child_Processes_Table
        If .DataBodyRange Is Nothing Then
            rRow = .HeaderRowRange.Row + 1
        Else
            rRow = .HeaderRowRange.Row + .DataBodyRange.Rows.count + 1
        End If
        
        .Parent.Cells(rRow, _
            .ListColumns("Child Process").Range.Column).Value = proc_id
        .Parent.Cells(rRow, _
            .ListColumns("Start Time").Range.Column).Value = Now()
    End With
End Sub

