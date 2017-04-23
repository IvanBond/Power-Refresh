Attribute VB_Name = "c_File_Refresher"
Option Explicit

Function Refresh_File() As Boolean
    ' general procedure that handles refreshing of target file
    ' considering all provided parameters
    
    ' only during test - comment for prod execution as these functions are executed in Workbook_Open
    'If Not Set_Global_Settings Then GoTo ErrHandler
    'If Not Check_Main_Parameters Then GoTo ErrHandler
    
    Dim arrIndex As Integer
    
    Call Write_Log("********************************* BEGIN *********************************", True)
    
    ' ALL is a keyword here
    ' if ALL - scan specific table for Scopes
    If ThisWorkbook.Names("SETTINGS_SCOPES").RefersToRange.Value = "ALL" Then
        ' open target file and scan for scopes
        If Not Open_Target_File Then GoTo ErrHandler
                
        Call Get_Scopes
        ' TODO - how to pass exclude scopes e.g. !UA, !US
        ' draft logic
        ' if Instr(..Scopes..., "!") >0  - exists exclude Scope
        ' it means automatically, that we need to scan ALL and then remove
        ' all excluded scopes
        
        If ScopesDictionary.count > 1 Then
            Call Refresh_File_For_Set_Of_Scopes
        ElseIf ScopesDictionary.count = 1 Then
            ' only one scope
            Call Refresh_File_One_or_No_Scopes(ScopesDictionary.Keys(0))
        Else
            ' no scopes found, maybe ControlTable is empty
            Call Refresh_File_One_or_No_Scopes_Try
            
            ' TODO - decide if it is an error case
            ' Scope 'ALL' was provided but nothing found
        End If
        
    ElseIf ThisWorkbook.Names("SETTINGS_SCOPES").RefersToRange.Value <> vbNullString Then
        
        ' some scopes were provided
        ' TODO: how to handle Exclude scope !
        If InStr(ThisWorkbook.Names("SETTINGS_SCOPES").RefersToRange.Value, ",") > 0 Then
            ' more than one scope
            arrScopes = Split(ThisWorkbook.Names("SETTINGS_SCOPES").RefersToRange.Value, ",")
            
            ' populate ScopesDictionary
            For arrIndex = LBound(arrScopes) To UBound(arrScopes)
                If Not ScopesDictionary.Exists(arrScopes(arrIndex)) Then
                    ScopesDictionary.Add arrScopes(arrIndex), arrScopes(arrIndex)
                End If
            Next arrIndex
            
            Call Refresh_File_For_Set_Of_Scopes
        Else
            ' only one scope
            Call Refresh_File_One_or_No_Scopes_Try(Trim(ThisWorkbook.Names("SETTINGS_SCOPES").RefersToRange.Value))
        End If ' check if more than one scope
    
    Else
        ' SCOPES cell is empty
        Call Refresh_File_One_or_No_Scopes_Try
    End If
            
Exit_Sub:
    Refresh_File = (Err.Number = 0) And (bGlobalError = False)
    Exit Function
    
ErrHandler:
    bGlobalError = True
    If Err.Number <> 0 Then
        Call Write_Log(Err.Number & ": " & Err.Description, True)
    End If
    GoTo Exit_Sub
    Resume ' for test purpose
End Function

Sub Refresh_File_For_Set_Of_Scopes()
    Dim kKey
    Dim BeforeRefresher
    Dim i As Byte
    Dim bStillRunning As Boolean
    
    ' at this point we have populated ScopesDictionary
    ' if parallel execution
    If ThisWorkbook.Names("SETTINGS_SCOPES_IN_PARALLEL").RefersToRange.Value = "Y" Then
        ' create child processes
        
        BeforeRefresher = Now()
        
        For Each kKey In ScopesDictionary.Keys
            Call Refresher_Caller(CStr(kKey), vbNullString)
        Next kKey
        
        ' monitor child processes
        ' until time_limit or processes are completed
        
        Do While 1 = 1
            Application.Wait (Now() + TimeValue("00:00:10")) ' 10 seconds
            'DoEvents ' ?
            
            ' we skip 0-element when populate array
            For i = LBound(arrChild_Processes, 2) + 1 To UBound(arrChild_Processes, 2)
                bStillRunning = False
                If arrChild_Processes(3, i).Status = 0 Then
                    bStillRunning = True
                End If
                
                ' TODO:
                ' if Debug Mode - display status in Child Processes Table
                
            Next i
            
            If Not bStillRunning Then
                Exit Do
            End If
            
            ' TODO:
            ' if general time limit exceeed -
            ' Write log Refresh Failed - Now # Child Process id # Scope. Refresh Failed #
            ' send email
            ' Terminate still active child processes
        Loop

    Else
        ' Refresh_File_For_Set_Of_Scopes one-by-one
        
        ' Run new process to avoid error stacking
        For Each kKey In ScopesDictionary.Keys
            Call Refresher_Caller(CStr(kKey), vbNullString)
        Next kKey
        
    End If ' if parallel
    
End Sub

Sub Refresh_File_One_or_No_Scopes_Try(Optional Scope As String)
    Refresh_Try = 1
    
    Do While Refresh_Try <= Refresh_Tries_Qty
        
        Call Refresh_File_One_or_No_Scopes(Scope)
        
        If bGlobalCriticalError Then
            ' do not run new tries in case of critical error that cannot be solved by new try
            GoTo ErrHandler
        End If
        
        If bGlobalError Then
            Call Write_Log("Failed refresh try " & Refresh_Try, bMandatoryLogRecord)
            
            Refresh_Try = Refresh_Try + 1
            If Refresh_Try > Refresh_Tries_Qty Then GoTo ErrHandler
            ' wait 10 min
            Call Write_Log("Waiting for next try... " & CStr(Delay_Between_Tries) & " min", bMandatoryLogRecord)
            bGlobalError = False
            
            ' Application.Wait makes pressure on CPU
            ' Application.Wait (Now() + TimeValue("00:" & Right("0" & CStr(Delay_Between_Tries), 2) & ":00"))
            Call WaitSeconds(600)
        Else
            Call Write_Log("Refresh Successful", bMandatoryLogRecord)
            Exit Do
        End If
    Loop

    Exit Sub
    
ErrHandler:
    Call Write_Log("Refresh Failed", bMandatoryLogRecord)
End Sub

Sub Refresh_File_One_or_No_Scopes(Optional Scope As String)
    ' General macro that is used to refresh files without Scopes
    ' or when only one Scope
    ' difference is only in set SCOPE cell or not in the beginning
    
    Dim BeforeAction
    
    If Not Open_Target_File Then GoTo ErrHandler

    Call Write_Log("Begin Refresh_File_One_or_No_Scopes")
    
    ' find SCOPE cell (named range) if was provided
    If Scope <> vbNullString Then
        On Error Resume Next
        Debug.Print target_wb.Names("SCOPE").Name
        If Err.Number <> 0 Then
            ' named range SCOPE not found
            Call Write_Log("Unexpected error: named range SCOPE was not found")
            Err.Clear
            bGlobalCriticalError = True
            GoTo ErrHandler
        Else
            target_wb.Names("SCOPE").RefersToRange.Value = Scope
        End If
        On Error GoTo 0
    End If
    
    ' TODO:
    ' apply Parameters to target workbook
    If ThisWorkbook.Names("SETTINGS_PARAMETERS").RefersToRange.Value <> vbNullString Then
        Call Set_Parameters
    End If
    
    Application.Calculate
    
    If ThisWorkbook.Names("SETTINGS_MACRO_BEFORE").RefersToRange.Value <> vbNullString Then
        BeforeAction = Now()
        Call Write_Log("Calling macro " & ThisWorkbook.Names("SETTINGS_MACRO_BEFORE").RefersToRange.Value)
        On Error Resume Next
        Run "'" & target_wb.Name & "'!" & ThisWorkbook.Names("SETTINGS_MACRO_BEFORE").RefersToRange.Value
        
        ' TODO: consider case with macro response
        ' e.g. macro checks that data source is not refreshed
        ' informs Refresher - it stops / delay execution
        If Err.Number <> 0 Then
            Call Write_Log("Unexpected error: " & Err.Number & ": " & Err.Description, True)
            Call Write_Log("Couldn't run macro_before " & ThisWorkbook.Names("SETTINGS_MACRO_BEFORE").RefersToRange.Value)
            Err.Clear
            bGlobalCriticalError = True
            GoTo ErrHandler
        End If
        On Error GoTo 0
        Call Write_Log("'Macro Before' completed # " & Round((Now() - BeforeAction) * 86400, 0) & "s")
    Else
        Call Write_Log("No Macro_Before")
    End If
        
    ' if do not skip RefreshAll
    If ThisWorkbook.Names("SETTINGS_SKIP_REFRESH_ALL").RefersToRange.Value <> "Y" Then
        BeforeAction = Now()
        Call Write_Log("Starting RefreshAll")
        If Not UpdateConnections Then
            ' error during refresh all
            Call Write_Log("Unexpected error during RefreshAll", True)
            GoTo ErrHandler
        End If
        Call Write_Log("RefreshAll Completed # " & Round((Now() - BeforeAction) * 86400, 0) & "s")
    Else
        Call Write_Log("Skipping RefreshAll")
    End If
    
    If ThisWorkbook.Names("SETTINGS_MACRO_AFTER").RefersToRange.Value <> vbNullString Then
        BeforeAction = Now()
        Call Write_Log("Calling macro " & ThisWorkbook.Names("SETTINGS_MACRO_AFTER").RefersToRange.Value)
        On Error Resume Next
        Run "'" & target_wb.Name & "'!" & ThisWorkbook.Names("SETTINGS_MACRO_AFTER").RefersToRange.Value
        ' macro can change parameter "Do not save" to "Y" to prevent save from Refresher
        ' in case of it understood that refresh failed, or no new data in data source
        
        If Err.Number <> 0 Then
            Call Write_Log("Unexpected error: " & Err.Number & ": " & Err.Description, True)
            Call Write_Log("Couldn't run macro_after" & ThisWorkbook.Names("SETTINGS_MACRO_AFTER").RefersToRange.Value)
            Err.Clear
            bGlobalCriticalError = True
            GoTo ErrHandler
        End If
        On Error GoTo 0
        Call Write_Log("'Macro After' completed # " & Round((Now() - BeforeAction) * 86400, 0) & "s")
    Else
        Call Write_Log("No Macro_After")
    End If
    
    Application.Calculate
    
    ' If parameter 'Do Not Save' is not Y
    If ThisWorkbook.Names("SETTINGS_DO_NOT_SAVE").RefersToRange.Value <> "Y" Then
        ' if save inplace
        If ThisWorkbook.Names("SETTINGS_SAVE_INPLACE").RefersToRange.Value = "Y" Then
            ' refresher does not check file format when save inplace
            If Not Save_Target_WB_Inplace Then GoTo Exit_Sub
        Else
            If Not Save_Target_WB_as_New(Scope) Then GoTo Exit_Sub
        End If
    Else
        Call Write_Log("Do_Not_Save option is enabled")
    End If
    
Exit_Sub:
    'If bGlobalError Then
    '    Call Write_Log("Refresh Failed", True)
    'Else
    '    Call Write_Log("Refresh Successful", True)
    'End If
    
    On Error Resume Next
    target_wb.Close SaveChanges:=False ' ? do we need to close it
    
    Exit Sub
    
ErrHandler:
    bGlobalError = True
    If Err.Number <> 0 Then
        Call Write_Log("Unexpected error: " & Err.Number & ": " & Err.Description, True)
    End If
    GoTo Exit_Sub
    Resume ' for test purpose
End Sub

Private Sub Set_Parameters()
    Dim arrParameters
    Dim i As Byte
    Dim delimiter_position As Long
    arrParameters = Split(ThisWorkbook.Names("SETTINGS_PARAMETERS").RefersToRange.Value, ",")
    ' TODO: add good parsing - what if comma is a part of parameter value?
    ' add instruciton that if comma used in parameter value - this value should be in quotes
    ' if quote is used in parameter value - then it must be doubled
    
    ' sample: FROM=01.01.2017, TO=17.01.2017
    
    For i = LBound(arrParameters) To UBound(arrParameters)
        delimiter_position = InStr(1, arrParameters(i), "=")
    
        If delimiter_position > 0 Then
            On Error Resume Next
            target_wb.Names(Trim(Left(arrParameters(i), delimiter_position - 1))).RefersToRange.Value = Trim(Mid(arrParameters(i), delimiter_position + 1))
            If Err.Number <> 0 Then
                Call Write_Log("Error. Couldn't find named range '" & Trim(Left(arrParameters(i), delimiter_position - 1)) & "' in target workbook", True)
                Err.Clear
            End If
            
            On Error GoTo 0
        Else
            Call Write_Log("Error. Wrong parameter structure " & arrParameters(i), True)
        End If
    Next i
End Sub
