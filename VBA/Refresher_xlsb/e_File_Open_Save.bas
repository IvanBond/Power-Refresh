Attribute VB_Name = "e_File_Open_Save"
Option Explicit

Function Open_Target_File() As Boolean
    Dim bReadOnly As Boolean
    On Error GoTo ErrHandler
    Dim BeforeAction
    BeforeAction = Now()
    Call Write_Log("Opening workbook... " & IIf(Left(ThisWorkbook.Names("SETTINGS_TARGET_PATH").RefersToRange.Value, 4) = "http", _
        Replace(ThisWorkbook.Names("SETTINGS_TARGET_PATH").RefersToRange.Value, " ", "%20"), _
        ThisWorkbook.Names("SETTINGS_TARGET_PATH").RefersToRange.Value))
    
    ' expression .Open(FileName, UpdateLinks, ReadOnly, Format, Password, WriteResPassword, IgnoreReadOnlyRecommended, Origin, Delimiter, Editable, Notify, Converter, AddToMru, Local, CorruptLoad)
    
    ' refresh for scope - always ReadOnly as result will be saved to a different workbook
    ' if need to refresh model with Scope and save inplace - use Parameters:SCOPE=XXX and Save_Inplace = Y in Reports Controller
    
    bReadOnly = (ThisWorkbook.Names("SETTINGS_SAVE_INPLACE").RefersToRange.Value = vbNullString)
    Set target_wb = Application.Workbooks.Open(Filename:=ThisWorkbook.Names("SETTINGS_TARGET_PATH").RefersToRange.Value, _
        UpdateLinks:=True, ReadOnly:=bReadOnly, IgnoreReadOnlyRecommended:=True, AddToMru:=False)
    
    ' for some reason, ReadOnly:=false doesn't work for files on SharePoint
    ' by default Excel 2016 opens files in read-only mode
    ' therefore, to be able to save them in place, we have to turn on "Edit Workbook" mode.
    
    On Error Resume Next
    If [SETTINGS_SAVE_INPLACE].Value <> vbNullString Then
        If Left([SETTINGS_TARGET_PATH].Value, 4) = "http" Then
            target_wb.LockServerFile
        End If
    End If
    Err.Clear
    On Error GoTo ErrHandler
    
    Application.Visible = (ThisWorkbook.Names("SETTINGS_DEBUG_MODE").RefersToRange.Value = "Y")
    
    target_wb.EnableAutoRecover = False
    ' other parameters can be added
     'If Not bReadOnly Then
     '   target_wb.LockServerFile
     'End If
     
     Call Write_Log("Workbook has been opened # " & Round((Now() - BeforeAction) * 86400, 0) & "s")
     
     Open_Target_File = True
Exit_Function:
    Exit Function

ErrHandler:
    Call Write_Log("Error on workbook opening")
    GoTo Exit_Function
End Function

Sub Get_Scopes()
' result can be - empty ScopesDictionary
' or having some values
    Dim sh As Worksheet
    Dim lo As ListObject
    Dim cell As Range
    ' procedure populates object ScopesDictionary with data from ControlTable
    
    For Each sh In target_wb.Sheets
        For Each lo In sh.ListObjects
        If lo.Name = CONTROL_TABLE_NAME Then
            On Error Resume Next
            ' check if column exists
            Debug.Print lo.ListColumns("Scope").Name
            If Err.Number = 0 Then
                On Error GoTo 0
                For Each cell In lo.ListColumns("Scope").DataBodyRange
                    If Not ScopesDictionary.Exists(cell.Value) Then
                        ScopesDictionary.Add cell.Value, cell.Value
                    End If
                Next cell
            End If ' If Err.Number = 0 Then
            On Error GoTo 0
            
            Exit Sub
        End If ' If lo.Name = CONTROL_TABLE_NAME Then
        Next lo
    Next sh
    
End Sub

Function Save_Target_WB_Inplace() As Boolean
    
    Call Write_Log("Saving as " & target_wb.FullName)
        
    ' just in case turn off display alerts
    Application.DisplayAlerts = False
    
    On Error Resume Next
    target_wb.Save
    Application.DisplayAlerts = (ThisWorkbook.Names("SETTINGS_DEBUG_MODE").RefersToRange.Value = "Y")
    
    target_wb.Close SaveChanges:=False
    
    If Err.Number <> 0 Then
        Call Write_Log("Unexpected error: " & Err.Number & ": " & Err.Description, True)
        Call Write_Log("Couldn't save workbook")
    Else
        Save_Target_WB_Inplace = True
    End If
End Function

Function Save_Target_WB_as_New(Optional Scope As String) As Boolean
    Dim result_path As String
    Dim separator As String
    Dim result_filename As String
    Dim new_wb As Workbook
    Dim BeforeAction
    
    On Error GoTo ErrHandler
    If ThisWorkbook.Names("SETTINGS_RESULT_FOLDER_PATH").RefersToRange.Value <> vbNullString Then
        ' attempt to handle save to SharePoint
        separator = IIf(InStr(ThisWorkbook.Names("SETTINGS_RESULT_FOLDER_PATH").RefersToRange.Value, "/") > 0, "/", "\") ' handle web address
        result_path = IIf(Right(ThisWorkbook.Names("SETTINGS_RESULT_FOLDER_PATH").RefersToRange.Value, 1) = "/" Or _
                                Right(ThisWorkbook.Names("SETTINGS_RESULT_FOLDER_PATH").RefersToRange.Value, 1) = "\", _
                           ThisWorkbook.Names("SETTINGS_RESULT_FOLDER_PATH").RefersToRange.Value, _
                           ThisWorkbook.Names("SETTINGS_RESULT_FOLDER_PATH").RefersToRange.Value & separator)
    Else
        separator = IIf(InStr(target_wb.path, "/") > 0, "/", "\")
        result_path = target_wb.path & separator
    End If
    
    result_filename = IIf(ThisWorkbook.Names("SETTINGS_RESULT_FILENAME").RefersToRange.Value <> vbNullString, _
        ThisWorkbook.Names("SETTINGS_RESULT_FILENAME").RefersToRange.Value, _
        ReportName) ' ReportName is name of target wb (without extension)
    
    ' include Scope
    result_filename = result_filename & IIf(Scope <> vbNullString, " " & Scope, vbNullString)
    
    ' get Extension
    Call Get_Resulting_Extension
    
    If ThisWorkbook.Names("SETTINGS_ADD_DATETIME").RefersToRange.Value = "Y" Then
        result_filename = result_filename & " " & Format(Now(), "yyyy-MM-dd hhmmss")
    Else
        ' as we are here - it is not Save_Inplace
        ' if no request for add datetime, if no scope and desired extension is the same - we need new name
        ' Reason: as target file was opened in ReadOnly mode - we need a new name
        ' check if
        ' no new folder provided and no new name or
        ' no new folder but new filename with same name
        ' new folder is the same and new filename is the same
        ' in other words: if same folder, same extension - we cannot re-write target file
        ' it is possible if only "Save Inplace" parameter is passed - then another procedure is called: Save_Target_WB_Inplace
        If Scope = vbNullString And Resulting_Extension = Right(target_wb.Name, Len(Resulting_Extension)) Then
            If (ThisWorkbook.Names("SETTINGS_RESULT_FOLDER_PATH").RefersToRange.Value = vbNullString And _
                        ThisWorkbook.Names("SETTINGS_RESULT_FILENAME").RefersToRange.Value = vbNullString) Or _
                    (ThisWorkbook.Names("SETTINGS_RESULT_FOLDER_PATH").RefersToRange.Value = target_wb.path And _
                        ThisWorkbook.Names("SETTINGS_RESULT_FILENAME").RefersToRange.Value = vbNullString) Or _
                    (ThisWorkbook.Names("SETTINGS_RESULT_FOLDER_PATH").RefersToRange.Value = target_wb.path And _
                        ThisWorkbook.Names("SETTINGS_RESULT_FILENAME").RefersToRange.Value = ReportName) Then
                result_filename = result_filename & " " & Format(Now(), "yyyy-MM-dd hhmmss")
            End If
        End If
        
    End If
        
    Set new_wb = target_wb
    
    ' (1) if Sheet Name is provided
    ' copy sheet to separate workbook and save it
    ' (2) if resulting fileformat is CSV - activate sheet with sheetname
    ' then - save active sheet as new file
    
    If ThisWorkbook.Names("SETTINGS_RESULT_SHEET_NAME").RefersToRange.Value <> vbNullString Then
        ' check if sheet exist
        On Error Resume Next
        target_wb.Sheets(ThisWorkbook.Names("SETTINGS_RESULT_SHEET_NAME").RefersToRange.Value).Activate
        If Err.Number = 0 Then
            target_wb.Sheets(ThisWorkbook.Names("SETTINGS_RESULT_SHEET_NAME").RefersToRange.Value).Copy
            Set new_wb = ActiveWorkbook
        Else
            ' sheet not found
            bGlobalError = True
            Call Write_Log("Couldn't find sheet " & ThisWorkbook.Names("SETTINGS_RESULT_SHEET_NAME").RefersToRange.Value)
            Exit Function
        End If
        On Error GoTo 0
    End If
    
    
    ' Change 2017-03-26
    ' Remove backward support of T-ransfer Result worksheet
    ' if type = T (transfer data)
    ' then sheet Result should be saved as a new workbook
'    If ThisWorkbook.Names("SETTINGS_SESSION_TYPE").RefersToRange.Value = "T" Then
'        On Error Resume Next
'        target_wb.Sheets("Result").Activate
'        If Err.Number = 0 Then
'            target_wb.Sheets("Result").Copy
'            Set new_wb = ActiveWorkbook
'        Else
'            ' sheet not found
'            bGlobalError = True
'            Call Write_Log("Couldn't find sheet 'Result'")
'            Exit Function
'        End If
'        On Error GoTo 0
'    End If
            
    ' SaveAs docu:
    ' https://msdn.microsoft.com/en-us/library/office/ff841185.aspx
        
    On Error Resume Next
    BeforeAction = Now()
    If Left(result_path, 4) = "http" Then
        Call Write_Log("Saving as " & Replace(result_path & result_filename, " ", "%20") & "." & Resulting_Extension)
    Else
        Call Write_Log("Saving as " & result_path & result_filename & "." & Resulting_Extension)
    End If
    
    Application.DisplayAlerts = False
    
    new_wb.SaveAs _
        Filename:=result_path & result_filename & "." & Resulting_Extension, _
        FileFormat:=Resulting_FileFormat, _
        ReadOnlyRecommended:=True, _
        ConflictResolution:=xlLocalSessionChanges, _
        AddToMru:=False, _
        AccessMode:=xlNoChange
    ' ToThink: what is a best practice for AccessMode - when save locally / network drive / SharePoint
    
    ' WriteResPassword - possible future improvement
    
    new_wb.Close SaveChanges:=False
    
    Application.DisplayAlerts = (ThisWorkbook.Names("SETTINGS_DEBUG_MODE").RefersToRange.Value = "Y")
    
    If Err.Number <> 0 Then
        bGlobalError = True
        Call Write_Log("Unexpected error: " & Err.Number & ": " & Err.Description, True)
        Call Write_Log("Couldn't save workbook")
    Else
        Call Write_Log("Saved succesfully # " & Round((Now() - BeforeAction) * 86400, 0) & "s")
        Save_Target_WB_as_New = True
    End If
    
    Exit Function
    
ErrHandler:
    ' TODO: handle errors
    On Error GoTo 0
    Call Write_Log("Unexpected error: " & Err.Number & ": " & Err.Description, True)
End Function
