Attribute VB_Name = "b_Globals"
Option Explicit

#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal lngMilliSeconds As Long)
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal lngMilliSeconds As Long)
#End If

Public Const Refresh_Tries_Qty = 3 ' after refresh fail script will try to refresh once again
Public Const Delay_Between_Tries = 10 ' in minutes, from 1 to 59

Public SMTP_Server As String
Public ErrorNotification_SendFrom As String
Public ErrorNotification_SendTo As String

Public LogsFolderPath As String ' in Refresher folder
Public Log_Enabled As Boolean

' T-ransfer constants
Public Const TRANSFER_SHEETNAME = "Result" ' sheet with this name will be saved as resulting file
Public Const DEFAULT_Transfer_FolderName = "Updated" ' in Refresher folder

' R-eport constants
Public Const CONTROL_TABLE_NAME = "ControlTable" ' table with list of Scopes
Public Const bMandatoryLogRecord = True

Public objFSO, objLog
Public objExcel
Public CurrentProcess
Public ProcessID As Long

Public ReportName As String
Public ScopesDictionary As Object
' Public StartPoint, BeforeAction, BeforeRefresher, StartTime
Public arrScopes, Scope
Public ExcelCreationTry As Byte
Public Refresh_Try As Byte
Public ReportID As String
Public bGlobalError As Boolean
Public bGlobalCriticalError As Boolean ' prevents new tries to refresh

Public Time_Limit_Per_File As Long ' in minutes
Public Time_Limit_Per_Scope As Long ' in minutes
Public Const Default_Time_Limit_Per_File = 360
Public Const Default_Time_Limit_Per_Scope = 60

Public target_wb As Workbook
Public Resulting_Extension As String
Public Resulting_FileFormat As Long

Public Child_Processes_Table As ListObject
Public arrChild_Processes() ' (3, 0) ' three parameters to store
 ' 0 - child process id
 ' 1 - Scope or Target File
 ' 2 - Start Time
 ' 3 - WshScriptExec Object - process, to get .Status
Public Child_Counter As Byte

Public Excel_Path As String
Public Refresher_Path As String
    
Function Check_Main_Parameters() As Boolean
    On Error Resume Next
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    ' check existence of files and folders
    If Left(ThisWorkbook.Names("SETTINGS_TARGET_PATH").RefersToRange.Value, 4) <> "http" Then
        If Not objFSO.FileExists(ThisWorkbook.Names("SETTINGS_TARGET_PATH").RefersToRange.Value) Then
            Call Write_Log("Error. Target file doesn't exists. " & ThisWorkbook.Names("SETTINGS_TARGET_PATH").RefersToRange.Value, True)
            GoTo ErrHandler
        End If
    End If
    
    ' check resulting folder if provided (and not save inplace)
    If ThisWorkbook.Names("SETTINGS_SAVE_INPLACE").RefersToRange.Value <> "Y" Then
        If ThisWorkbook.Names("SETTINGS_RESULT_FOLDER_PATH").RefersToRange.Value <> vbNullString Then
            If Left(ThisWorkbook.Names("SETTINGS_TARGET_PATH").RefersToRange.Value, 4) <> "http" Then
                If Not objFSO.FolderExists(ThisWorkbook.Names("SETTINGS_RESULT_FOLDER_PATH").RefersToRange.Value) Then
                    Call Write_Log("Error. Resulting folder doesn't exists. " & ThisWorkbook.Names("SETTINGS_RESULT_FOLDER_PATH").RefersToRange.Value, True)
                    GoTo ErrHandler
                End If
            End If
        End If
    End If
    
    ' check conflicts between parameters
    
    ' if SCOPE is provided and Save Inplace = Y
'    If ThisWorkbook.Names("SETTINGS_SAVE_INPLACE").RefersToRange.Value = "Y" And _
'            ThisWorkbook.Names("SETTINGS_SCOPES").RefersToRange.Value <> vbNullString Then
'        Call Write_Log("Error. Refresh for SCOPES cannot be done with 'Save InPlace' = 'Y' ", True)
'        GoTo ErrHandler
'    End If
    
    ' if format CSV but sheetname is empty
    If ThisWorkbook.Names("SETTINGS_RESULT_FILE_EXTENSION").RefersToRange.Value = "CSV" And _
            ThisWorkbook.Names("SETTINGS_RESULT_SHEET_NAME").RefersToRange.Value = vbNullString Then
        Call Write_Log("Error. For CSV format you should provide a name of sheet you want to save!", True)
        GoTo ErrHandler
    End If
    
    ' if SCOPES is not empty and Parameters contains SCOPE
    If ThisWorkbook.Names("SETTINGS_SCOPES").RefersToRange.Value <> vbNullString And _
        InStr(ThisWorkbook.Names("SETTINGS_PARAMETERS").RefersToRange.Value, "SCOPE=") > 0 Then
        Call Write_Log("Error. You cannot use Parameter SCOPE with non-empty 'Scopes'!", True)
        GoTo ErrHandler
    End If
        
    ' if Save Inplace and SheetName - danger, model can be rewritten by one-sheet-workbook
    If ThisWorkbook.Names("SETTINGS_SAVE_INPLACE").RefersToRange.Value = "Y" And _
        ThisWorkbook.Names("SETTINGS_RESULT_SHEET_NAME").RefersToRange.Value <> vbNullString Then
        Call Write_Log("Error. Refresher avoids scenario 'Save Inplace' only one worksheet.", True)
        GoTo ErrHandler
    End If
    
    
    ' warnings
    ' TOThink
    'if Save_Inplace - other parameters will be ignored (if non-empty - warning to log
    
Exit_Sub:
    Check_Main_Parameters = (Err.Number = 0)
    bGlobalError = (Err.Number <> 0)
    Exit Function
    
ErrHandler:
    bGlobalError = True
End Function

Function Set_Global_Settings() As Boolean
    On Error Resume Next
    With Application
    ' visible if only Debug Mode
        .Visible = (ThisWorkbook.Names("SETTINGS_DEBUG_MODE").RefersToRange.Value = "Y")
        .DisplayAlerts = (ThisWorkbook.Names("SETTINGS_DEBUG_MODE").RefersToRange.Value = "Y")
        .ScreenUpdating = (ThisWorkbook.Names("SETTINGS_DEBUG_MODE").RefersToRange.Value = "Y")
        
        '.Interactive = (ThisWorkbook.Names("SETTINGS_DEBUG_MODE").RefersToRange.Value <> vbNullString)
        ' .EnableCancelKey = xlDisabled  'Turn off the Esc key
        
        Excel_Path = Application.path & "\EXCEL.EXE"
        ' "C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"
    End With
    
    With ThisWorkbook
        Refresher_Path = .path & "\Refresher.xlsb"
        
        ' TODO: resolve issue when Refresher is stored in a folder which is synced via OneDrive or OneDrive for business
        ' in such case Workbook.Path gives URL to file
        ' Log file cannot be created in a web
        ' how to get local path that is synced by OneDrive?
        LogsFolderPath = .path & "\Logs\"
        Log_Enabled = (.Names("SETTINGS_LOG_ENABLED").RefersToRange.Value = "Y") Or _
                (.Names("SETTINGS_DEBUG_MODE").RefersToRange.Value = "Y")
                
        ReportID = .Names("SETTINGS_REPORT_ID").RefersToRange.Value
        
        SMTP_Server = .Names("SETTINGS_SMTP_SERVER").RefersToRange.Value
        ErrorNotification_SendFrom = .Names("SETTINGS_EMAIL_FROM").RefersToRange.Value
        ErrorNotification_SendTo = .Names("SETTINGS_GENERAL_EMAIL_TO").RefersToRange.Value
    End With
    
    ReportName = GetReportName
        
    Set ScopesDictionary = CreateObject("Scripting.Dictionary")
    ScopesDictionary.CompareMode = vbTextCompare
    
    ' calculate Time Limit per File / per Scope
    If ThisWorkbook.Names("SETTINGS_TIME_LIMIT").RefersToRange.Value <> vbNullString Then
        If Right(ThisWorkbook.Names("SETTINGS_TARGET_PATH").RefersToRange.Value, 1) = "\" Or _
            Right(ThisWorkbook.Names("SETTINGS_TARGET_PATH").RefersToRange.Value, 1) = "/" Then
            ' refresh folder
            If ThisWorkbook.Names("SETTINGS_FILES_IN_PARALLEL").RefersToRange.Value = "Y" Then
                Time_Limit_Per_File = ThisWorkbook.Names("SETTINGS_TIME_LIMIT").RefersToRange.Value
            Else
                Time_Limit_Per_File = Round(ThisWorkbook.Names("SETTINGS_TIME_LIMIT").RefersToRange.Value / _
                    objFSO.GetFolder(ThisWorkbook.Names("SETTINGS_TARGET_PATH").RefersToRange.Value).Files.count, 0)
            End If
        Else
            ' refresh file
            If ThisWorkbook.Names("SETTINGS_SCOPES_IN_PARALLEL").RefersToRange.Value = "Y" Then
                Time_Limit_Per_Scope = ThisWorkbook.Names("SETTINGS_TIME_LIMIT").RefersToRange.Value
            Else
                Time_Limit_Per_Scope = Round(ThisWorkbook.Names("SETTINGS_TIME_LIMIT").RefersToRange.Value / _
                    (Len(ThisWorkbook.Names("SETTINGS_SCOPES").RefersToRange.Value) - _
                        Len(Replace(ThisWorkbook.Names("SETTINGS_SCOPES").RefersToRange.Value, ",", "")) + 1 _
                     ), 0) ' calc number of scopes
            End If
        End If
    End If
    
    If Time_Limit_Per_File = 0 Then Time_Limit_Per_File = Default_Time_Limit_Per_File
    If Time_Limit_Per_Scope = 0 Then Time_Limit_Per_Scope = Default_Time_Limit_Per_Scope
    
    Set Child_Processes_Table = ThisWorkbook.Sheets("Refresher").ListObjects("CHILD_PROCESSES")
    
    bGlobalError = (Err.Number <> 0)
    Set_Global_Settings = (Err.Number = 0)
End Function

Sub Get_Resulting_Extension()
    ' https://msdn.microsoft.com/en-us/library/office/ff198017.aspx
    ' CSV = xlCSV = 6
    ' xlsx = xlOpenXMLWorkbook = 51
    ' xlsm = xlOpenXMLWorkbookMacroEnabled = 52
    ' xlsb = xlExcel12 = 50
    
    If ThisWorkbook.Names("SETTINGS_RESULT_FILE_EXTENSION").RefersToRange.Value = vbNullString Then
        ' just take initial extention
        ' as we refresh only Excel files - just take last 4 chars
        Resulting_Extension = Replace(Right(target_wb.Name, 4), ".", "")
    Else
        Resulting_Extension = ThisWorkbook.Names("SETTINGS_RESULT_FILE_EXTENSION").RefersToRange.Value
    End If
    
    Resulting_Extension = LCase(Resulting_Extension)
    
    ' Get Resulting FileFormat
    If Resulting_Extension = "xlsx" Then
        Resulting_FileFormat = xlOpenXMLWorkbook
    ElseIf Resulting_Extension = "xlsb" Then
        Resulting_FileFormat = xlExcel12
    ElseIf Resulting_Extension = "xlsm" Then
        Resulting_FileFormat = xlOpenXMLWorkbookMacroEnabled
    ElseIf Resulting_Extension = "csv" Then
        Resulting_FileFormat = xlCSV
    End If
    
End Sub
