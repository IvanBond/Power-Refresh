Attribute VB_Name = "Main"
Option Explicit
'
' procedure is operated by Timer, but can be executed manually as well
Sub Check_And_Run()

    Dim cell As Range
    Dim Field_Name As String
    Dim objShell, objScriptEngine, objProc As Object
    Dim log_row As Long
    Dim StartTime As Double
    Dim arrScopes
    Dim i As Long
    
    Dim sh As Worksheet
    ' exit from Edit mode (user edit cell) - just in case
    Set sh = ThisWorkbook.ActiveSheet
    Application.ScreenUpdating = False
    ThisWorkbook.Sheets("LOG").Activate
    sh.Activate ' should force Excel to exit from edit mode
    Application.ScreenUpdating = True
    
    Call Set_Global_Variables
    Set objShell = CreateObject("WScript.Shell")
    Set objScriptEngine = CreateObject("scriptcontrol")
    objScriptEngine.Language = "JScript"
    
    If Control_Table.DataBodyRange Is Nothing Then
        MsgBox "No reports for execution", vbExclamation + vbOKOnly, "Information"
        Exit Sub
    End If
    
    ' build parameters string
    ' loop through rows
    For Each cell In Control_Table.ListColumns("Report ID *").DataBodyRange
        ' if Status contains In Process
        If Left(Control_Table.Parent.Cells(cell.Row, Control_Table.ListColumns("Status").Range.Column).Value, 10) = "In Process" Then
            ' Update status if still in process
            ' check if process exists
            If CheckProcessExist(Get_Last_Log_Record(cell.Row, "Process ID")) Then
                StartTime = Get_Last_Log_Record(cell.Row, "Start Time") ' get process start time
                If StartTime <> -1 Then
                    Control_Table.Parent.Cells(cell.Row, _
                            Control_Table.ListColumns("Status").Range.Column).Value = "In Process: " & Format(Now() - StartTime, "hh:mm:ss")
                End If
            Else
            ' if process doesn't exist anymore
                Control_Table.Parent.Cells(cell.Row, Control_Table.ListColumns("Status").Range.Column).Value = Replace( _
                    Control_Table.Parent.Cells(cell.Row, Control_Table.ListColumns("Status").Range.Column).Value, _
                    "In Process", "Completed", compare:=vbTextCompare) & "+"
            End If
        End If
        
        ' check all conditions for run
        ' Enabled or Next Run time is passed
        ' or if Manual Trigger
        If (Control_Table.Parent.Cells(cell.Row, Control_Table.ListColumns("Enabled").Range.Column).Value = "Y" And _
            Control_Table.Parent.Cells(cell.Row, Control_Table.ListColumns("Next Run").Range.Column).Value < Now()) Or _
                Control_Table.Parent.Cells(cell.Row, Control_Table.ListColumns("Manual Trigger").Range.Column).Value <> vbNullString Then
            
            ' just in case check row validity
            If Is_Row_Valid(cell.Row) Then
                
                Control_Table.Parent.Cells(cell.Row, Control_Table.ListColumns("Last Run").Range.Column).Value = Now()
                
                ' if we run task that is planned on e.g. next week via putting Manual Trigger - we do not need to re-calc Next Run
                ' in other words, we re-calc next run only then Next Run < Now()
                
                If Control_Table.Parent.Cells(cell.Row, Control_Table.ListColumns("Next Run").Range.Column).Value < Now() Then
                    ' Calculate Next Run
                     Control_Table.Parent.Cells(cell.Row, _
                        Control_Table.ListColumns("Next Run").Range.Column).Value = Get_Next_Run_DateTime(cell.Row)
                     
                     ' Debug.Print Get_Next_Run_DateTime(cell.Row)
                     
                End If

                ' remove manual trigger - every time
                Control_Table.Parent.Cells(cell.Row, _
                 Control_Table.ListColumns("Manual Trigger").Range.Column).Value = vbNullString

                ' clear Status - as this code is executed only when we start new instance
                Control_Table.Parent.Cells(cell.Row, _
                    Control_Table.ListColumns("Status").Range.Column).Value = "In Process: 0:00"
                
                ' Run Excel with switches /x /r /e
                ' about switches https://support.microsoft.com/en-us/kb/291288
                'Debug.Print objScriptEngine.Run("encodeURIComponent", Collect_Parameters(cell.Row))
                
                ' for some reason when place /r parameter as first - /e and rest is ignored
                'Set objProc = objShell.Exec(Excel_Path & " /r """ & Refresher_Path & """ /x " & _
                    "/e" & objScriptEngine.Run("encodeURIComponent", Collect_Parameters(cell.Row)))
                
                ' therefore - /r is last parameter
                ' order is important for Parsing macro in Refresher.xlsb
                
                ' in case of Parallel refresh of Scopes we should start new Excel instance for each Scope
                Set objProc = objShell.Exec(Excel_Path & " /x " & _
                        "/e" & objScriptEngine.Run("encodeURIComponent", Collect_Parameters(cell.Row)) & _
                        " /r """ & Refresher_Path & """")

' OLD Part when Reports Controller had started separate processes for different scopes
' now Refresher starts it on his own
'                If Control_Table.Parent.Cells(cell.Row, _
'                    Control_Table.ListColumns("Parallel Refresh of Scopes").Range.Column).Value = "Y" Then
'
'                    arrScopes = Split(Control_Table.Parent.Cells(cell.Row, _
'                        Control_Table.ListColumns("Scopes").Range.Column).Value, ",")
'
'                    For i = LBound(arrScopes) To UBound(arrScopes)
'                        ' call Collect_Parameters with Scope param
'                        Set objProc = objShell.Exec(Excel_Path & " /x " & _
'                            "/e" & objScriptEngine.Run("encodeURIComponent", Collect_Parameters(cell.Row, Trim(arrScopes(i)))) & _
'                            " /r """ & Refresher_Path & """")
'
'                        Application.Wait Now() + TimeValue("00:00:03") ' just in case pause between start of Excel application
'
'                    Next i
'
'                Else
'                    ' not parallel execution
'                    ' just pass all parameters to Refresher
'                    Set objProc = objShell.Exec(Excel_Path & " /x " & _
'                        "/e" & objScriptEngine.Run("encodeURIComponent", Collect_Parameters(cell.Row)) & _
'                        " /r """ & Refresher_Path & """")
'                End If
                
                ' URL encoding is needed because otherwise no chance to pass values with spaces and special chars
                ' Excel execution through shell will trigger every value separated by space as new file to be opened
                
                ' can run without wait if no workbooks in
                ' C:\Users\<username>\AppData\Roaming\Microsoft\Excel\XLSTART
                Application.Wait Now() + TimeValue("00:00:03")
                
                ' populate log table
                Call Write_Log(cell.Row, objProc.ProcessID)

                Set objProc = Nothing
            Else
                ' row is not valid - function Is_Row_Valid put necessary comment to Status field
            End If ' if row is valid
            
            'Debug.Print cell.Row; Control_Table.Parent.Cells(cell.Row, Control_Table.ListColumns("Enabled").Range.Column).Value
        End If ' check if enabled

Next_Cell:
    Next cell
    
Exit_sub:
    
    Set objShell = Nothing
    Set objScriptEngine = Nothing
    
    Exit Sub
    
ErrHandler:
    Debug.Print Err.Number & Err.Description
    GoTo Exit_sub
    Resume
End Sub

Function Is_Row_Valid(row_id As Long) As Boolean
    Dim Field_Name  As String
    
    ' ******************** check mandatory fields ********************
    
    Field_Name = "File or Folder Path *"
    If Control_Table.Parent.Cells(row_id, Control_Table.ListColumns(Field_Name).Range.Column).Value = vbNullString Then
        Control_Table.Parent.Cells(row_id, Control_Table.ListColumns("Status").Range.Column).Value = "Cannot execute. Provide valid report path"
        Exit Function
    End If
    
    ' ******************** end of check mandatory fields ********************
    
    ' ******************** fields validation ********************
        'todo
        ' check conflicts of saving options
        ' check existence ot resulting path
        '
    ' ******************** end of fields validation ********************
    
    ' if passed
    Is_Row_Valid = True
End Function

Function encodeURL(str As String)
    Dim ScriptEngine As Object
    ' encode spaces, commas, other chars using URL encoded chars
    ' http://www.degraeve.com/reference/urlencoding.php
    
    Set ScriptEngine = CreateObject("scriptcontrol")
    ScriptEngine.Language = "JScript"
    encodeURL = ScriptEngine.Run("encodeURIComponent", str)
    Set ScriptEngine = Nothing
End Function

Function Collect_Parameters(report_row_id As Long, Optional Scope As String) As String
    Dim str As String
    Dim Field_Name As String
    
    ' Collect_Parameters = "/debug_mode/log_enabled/file_path:C:\Temp\Test.xlsx/" & _
        "result_folder_path:\\server_name\ssis in\/macro_before:test/macro_after:after" & _
        "/error_email_to:test@domain.com/success_email_to:test@domain.com/result_filename:reportX/extension:xlsb" & _
        "/add_datetime/scopes:KZ HR,UA AMS/parameters:FROM=2016-01-01,TO=2016-10-31"
    
    Field_Name = "Report ID *"
    Collect_Parameters = "/report_id:" & Control_Table.Parent.Cells(report_row_id, _
        Control_Table.ListColumns(Field_Name).Range.Column).Value
    
    Field_Name = "File or Folder Path *"
    Collect_Parameters = Collect_Parameters & "/target_path:" & _
    Replace(Control_Table.Parent.Cells(report_row_id, Control_Table.ListColumns(Field_Name).Range.Column).Value, _
            "/", "|") ' cheap solution - replace '/' before transfer parameters
                ' solves problem of web folder
    ' '|' cannot be used in path, so it will be replaced back in Refresher
    
    Field_Name = "Type"
        Collect_Parameters = Collect_Parameters & "/type:" & _
            IIf(Control_Table.Parent.Cells(report_row_id, _
                    Control_Table.ListColumns(Field_Name).Range.Column).Value = vbNullString, _
                "R", _
                Control_Table.Parent.Cells(report_row_id, _
                    Control_Table.ListColumns(Field_Name).Range.Column).Value)
    
    Field_Name = "Macro Before"
    If Control_Table.Parent.Cells(report_row_id, _
        Control_Table.ListColumns(Field_Name).Range.Column).Value <> vbNullString Then
        Collect_Parameters = Collect_Parameters & "/macro_before:" & Control_Table.Parent.Cells(report_row_id, _
            Control_Table.ListColumns(Field_Name).Range.Column).Value
    End If
    
    Field_Name = "Skip RefreshAll"
    If Control_Table.Parent.Cells(report_row_id, _
        Control_Table.ListColumns(Field_Name).Range.Column).Value <> vbNullString Then
        Collect_Parameters = Collect_Parameters & "/skip_refresh_all"
    End If
    
    Field_Name = "Macro After"
    If Control_Table.Parent.Cells(report_row_id, _
        Control_Table.ListColumns(Field_Name).Range.Column).Value <> vbNullString Then
        Collect_Parameters = Collect_Parameters & "/macro_after:" & Control_Table.Parent.Cells(report_row_id, _
            Control_Table.ListColumns(Field_Name).Range.Column).Value
    End If
    
    Field_Name = "Error Email"
    If Control_Table.Parent.Cells(report_row_id, _
        Control_Table.ListColumns(Field_Name).Range.Column).Value <> vbNullString Then
        Collect_Parameters = Collect_Parameters & "/error_email_to:" & Control_Table.Parent.Cells(report_row_id, _
            Control_Table.ListColumns(Field_Name).Range.Column).Value
    End If
    
    Field_Name = "Success Email"
    If Control_Table.Parent.Cells(report_row_id, _
        Control_Table.ListColumns(Field_Name).Range.Column).Value <> vbNullString Then
        Collect_Parameters = Collect_Parameters & "/success_email_to:" & Control_Table.Parent.Cells(report_row_id, _
            Control_Table.ListColumns(Field_Name).Range.Column).Value
    End If
        
    Field_Name = "Debug Mode"
    If Control_Table.Parent.Cells(report_row_id, _
        Control_Table.ListColumns(Field_Name).Range.Column).Value = "Y" Then
        Collect_Parameters = Collect_Parameters & "/debug_mode"
    End If
    
    Field_Name = "Log Enabled"
    If Control_Table.Parent.Cells(report_row_id, _
        Control_Table.ListColumns(Field_Name).Range.Column).Value = "Y" Then
        Collect_Parameters = Collect_Parameters & "/log_enabled"
    End If
    
    Field_Name = "Save Inplace"
    If Control_Table.Parent.Cells(report_row_id, _
        Control_Table.ListColumns(Field_Name).Range.Column).Value = "Y" Then
        Collect_Parameters = Collect_Parameters & "/save_inplace"
    End If
    
    Field_Name = "Do Not Save"
    If Control_Table.Parent.Cells(report_row_id, _
        Control_Table.ListColumns(Field_Name).Range.Column).Value = "Y" Then
        Collect_Parameters = Collect_Parameters & "/do_not_save"
    End If
    
    Field_Name = "Result Folder Path"
    If Control_Table.Parent.Cells(report_row_id, _
        Control_Table.ListColumns(Field_Name).Range.Column).Value <> vbNullString Then
        Collect_Parameters = Collect_Parameters & "/result_folder_path:" & _
         Replace(Control_Table.Parent.Cells(report_row_id, Control_Table.ListColumns(Field_Name).Range.Column).Value, _
                "/", "|") ' cheap solution - replace '/' before transfer parameters
                ' solves problem of web folder
                ' '|' cannot be used in path, so it will be replaced back in Refresher
    End If
    
    Field_Name = "Result FileName"
    If Control_Table.Parent.Cells(report_row_id, _
        Control_Table.ListColumns(Field_Name).Range.Column).Value <> vbNullString Then
        Collect_Parameters = Collect_Parameters & "/result_filename:" & Control_Table.Parent.Cells(report_row_id, _
            Control_Table.ListColumns(Field_Name).Range.Column).Value
    End If
    
    Field_Name = "Add Datetime"
    If Control_Table.Parent.Cells(report_row_id, _
        Control_Table.ListColumns(Field_Name).Range.Column).Value = "Y" Then
        Collect_Parameters = Collect_Parameters & "/add_datetime"
    End If
    
    Field_Name = "Extension"
    If Control_Table.Parent.Cells(report_row_id, _
        Control_Table.ListColumns(Field_Name).Range.Column).Value <> vbNullString Then
        Collect_Parameters = Collect_Parameters & "/extension:" & Control_Table.Parent.Cells(report_row_id, _
            Control_Table.ListColumns(Field_Name).Range.Column).Value
    End If
    
    Field_Name = "Time Limit"
    If Control_Table.Parent.Cells(report_row_id, _
        Control_Table.ListColumns(Field_Name).Range.Column).Value <> vbNullString Then
        Collect_Parameters = Collect_Parameters & "/time_limit:" & Control_Table.Parent.Cells(report_row_id, _
            Control_Table.ListColumns(Field_Name).Range.Column).Value
    End If
    
    Field_Name = "Parallel Refresh of Scopes"
    If Control_Table.Parent.Cells(report_row_id, _
        Control_Table.ListColumns(Field_Name).Range.Column).Value <> vbNullString Then
        Collect_Parameters = Collect_Parameters & "/scopes_in_parallel"
    End If
    
    Field_Name = "Parallel Refresh of files in folder"
    If Control_Table.Parent.Cells(report_row_id, _
        Control_Table.ListColumns(Field_Name).Range.Column).Value <> vbNullString Then
        Collect_Parameters = Collect_Parameters & "/files_in_parallel"
    End If
        
    ' if save only one sheet
    Field_Name = "Save Sheet"
    If Control_Table.Parent.Cells(report_row_id, _
        Control_Table.ListColumns(Field_Name).Range.Column).Value <> vbNullString Then
        Collect_Parameters = Collect_Parameters & "/save_sheet:" & Control_Table.Parent.Cells(report_row_id, _
            Control_Table.ListColumns(Field_Name).Range.Column).Value
    End If


    Field_Name = "Parameters"
    If Control_Table.Parent.Cells(report_row_id, _
        Control_Table.ListColumns(Field_Name).Range.Column).Value <> vbNullString Then
        Collect_Parameters = Collect_Parameters & "/parameters:" & _
            Replace(Control_Table.Parent.Cells(report_row_id, Control_Table.ListColumns(Field_Name).Range.Column).Value, _
                    "/", "{|}")
            ' potential risk of appear '/'
            ' changed back on Refresher side during params parsing
    End If
    
    ' Scope can be passed as argument
    If Scope = vbNullString Then
        Field_Name = "Scopes"
        If Control_Table.Parent.Cells(report_row_id, _
            Control_Table.ListColumns(Field_Name).Range.Column).Value <> vbNullString Then
            
            Collect_Parameters = Collect_Parameters & "/scopes:" & _
            Replace(Control_Table.Parent.Cells(report_row_id, Control_Table.ListColumns(Field_Name).Range.Column).Value, _
                "/", "{|}")
        End If
    Else
        Collect_Parameters = Collect_Parameters & "/scopes:" & Replace(Scope, "/", "{|}")
    End If
    
    Debug.Print Collect_Parameters
End Function
