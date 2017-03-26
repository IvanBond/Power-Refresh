Attribute VB_Name = "a_Command_Line_Args"
Option Explicit

Sub ParseArgs(sCmdLine As String)
    'Pulls the command line arguments/parameters and returns them as an array.
    'My method for command line arguments.  Toby Erkson.
    
    Dim iStart As Long
    Dim iEnd As Long
    Dim sArgs As String
    Dim vArgs, vTemp As Variant
    Dim x As Long
    Dim delimiter_position As Byte
    Dim param_key, param_value As String
    
    ' use " /x " after /r "file path" and only then /e[/par1:val1/par2:val2]
    iStart = InStr(1, sCmdLine, " /x /e") ' can be an issue with folders on SharePoint ?
    If iStart = 0 Then Exit Sub       'Couldn't find ' /e' so assume no parameters were given
    'sArgs = decodeURL(Mid(sCmdLine, iStart + 6)) ' pass URL encoded parameters to be able to use spaces in file / folder path, commas in lists and other chars
    
    iEnd = InStr(1, sCmdLine, " /r """) ' if file path provided in the end
    sArgs = decodeURL(Mid(sCmdLine, iStart + 6, iEnd - iStart - 6))
    
    If Len(sArgs) = 0 Then Exit Sub       'No command line parameters were given
            
    ' all '/' chars have to be replaced before pass to command line
    ' it helps to simplify 'split' process
    'Loop thru the arguments and fill array.
    'index(n, 0) is the key or defined parameter
    'index(n, 1) is the user supplied value
    vArgs = Split(sArgs, "/") ' expected vbs NamedArguments format, when parameters are preceding with "/"
    
    For x = 1 To UBound(vArgs) ' skip empty element
        delimiter_position = InStr(1, vArgs(x), ":")
        If delimiter_position > 0 Then
            param_key = Left(vArgs(x), delimiter_position - 1)
            param_value = Mid(vArgs(x), delimiter_position + 1)
        Else
            param_key = vArgs(x)
            param_value = vbNullString
        End If
        
        vTemp = Array(param_key, param_value) 'Split(vArgs(x), ":")  'Break up the arguements (key:value)
        
        If vTemp(0) = "debug_mode" Then
            ThisWorkbook.Names("SETTINGS_DEBUG_MODE").RefersToRange.Value = "Y"
        ElseIf vTemp(0) = "report_id" Then
            ThisWorkbook.Names("SETTINGS_REPORT_ID").RefersToRange.Value = Trim(vTemp(1))
        ElseIf vTemp(0) = "log_enabled" Then
            ThisWorkbook.Names("SETTINGS_LOG_ENABLED").RefersToRange.Value = "Y"
        'ElseIf vTemp(0) = "type" Then
        '    ThisWorkbook.Names("SETTINGS_SESSION_TYPE").RefersToRange.Value = Trim(vTemp(1))
        ElseIf vTemp(0) = "target_path" Then
            ThisWorkbook.Names("SETTINGS_TARGET_PATH").RefersToRange.Value = Replace(vTemp(1), "|", "/") ' return '/'
        ElseIf vTemp(0) = "macro_before" Then
            ThisWorkbook.Names("SETTINGS_MACRO_BEFORE").RefersToRange.Value = Trim(vTemp(1))
        ElseIf vTemp(0) = "macro_after" Then
            ThisWorkbook.Names("SETTINGS_MACRO_AFTER").RefersToRange.Value = Trim(vTemp(1))
        ElseIf vTemp(0) = "error_email_to" Then
            ThisWorkbook.Names("SETTINGS_ERROR_EMAIL_TO").RefersToRange.Value = Trim(vTemp(1))
        ElseIf vTemp(0) = "success_email_to" Then
            ThisWorkbook.Names("SETTINGS_SUCCESS_EMAIL_TO").RefersToRange.Value = Trim(vTemp(1))
        ElseIf vTemp(0) = "do_not_save" Then
            ThisWorkbook.Names("SETTINGS_DO_NOT_SAVE").RefersToRange.Value = "Y"
        ElseIf vTemp(0) = "save_inplace" Then
            ThisWorkbook.Names("SETTINGS_SAVE_INPLACE").RefersToRange.Value = "Y"
        ElseIf vTemp(0) = "result_folder_path" Then
            ThisWorkbook.Names("SETTINGS_RESULT_FOLDER_PATH").RefersToRange.Value = Replace(vTemp(1), "|", "/")
        ElseIf vTemp(0) = "result_filename" Then
            ThisWorkbook.Names("SETTINGS_RESULT_FILENAME").RefersToRange.Value = vTemp(1)
        ElseIf vTemp(0) = "extension" Then
            ThisWorkbook.Names("SETTINGS_RESULT_FILE_EXTENSION").RefersToRange.Value = Trim(vTemp(1))
        ElseIf vTemp(0) = "add_datetime" Then
            ThisWorkbook.Names("SETTINGS_ADD_DATETIME").RefersToRange.Value = "Y"
        ElseIf vTemp(0) = "scopes" Then
            ThisWorkbook.Names("SETTINGS_SCOPES").RefersToRange.Value = Replace(vTemp(1), "{|}", "/")
        ElseIf vTemp(0) = "parameters" Then
            ThisWorkbook.Names("SETTINGS_PARAMETERS").RefersToRange.Value = Replace(vTemp(1), "{|}", "/")
        ElseIf vTemp(0) = "skip_refresh_all" Then
            ThisWorkbook.Names("SETTINGS_SKIP_REFRESH_ALL").RefersToRange.Value = "Y"
        ElseIf vTemp(0) = "time_limit" Then
            ThisWorkbook.Names("SETTINGS_TIME_LIMIT").RefersToRange.Value = vTemp(1)
        ElseIf vTemp(0) = "files_in_parallel" Then
            ThisWorkbook.Names("SETTINGS_FILES_IN_PARALLEL").RefersToRange.Value = "Y"
        ElseIf vTemp(0) = "scopes_in_parallel" Then
            ThisWorkbook.Names("SETTINGS_SCOPES_IN_PARALLEL").RefersToRange.Value = "Y"
        ElseIf vTemp(0) = "save_sheet" Then
            ThisWorkbook.Names("SETTINGS_RESULT_SHEET_NAME").RefersToRange.Value = vTemp(1)
        End If
    Next x
    
End Sub

