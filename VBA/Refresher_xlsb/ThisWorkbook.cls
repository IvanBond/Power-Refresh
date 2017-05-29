VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()
    Dim str As String
    
    [START_TIME].Value = Now
    Call GetCurrentProcess
    
    ThisWorkbook.Names("SETTINGS_PROCESS_ID").RefersToRange.Value = CurrentProcess.ProcessID
    ThisWorkbook.Names("SETTINGS_PARENT_PROCESS_ID").RefersToRange.Value = CurrentProcess.ParentProcessID
    ProcessID = CurrentProcess.ProcessID
    
    str = CurrentProcess.CommandLine
    ' CreationDate = CurrentProcess.CreationDate
    ' CurrentProcess.ExecutablePath - to get Excel Path
    
    ' uncomment for debugging (in case of problem) - encoded string
    'ThisWorkbook.Names("COMMAND_LINE_DECODED").RefersToRange.Value = str
    
    ThisWorkbook.Names("COMMAND_LINE_DECODED").RefersToRange.Value = decodeURL(str)
    
    ' if manual opening - just exit sub
    ' assume combination of switches can be only when we call Refresher from Reports Controller
    If InStr(ThisWorkbook.Names("COMMAND_LINE_DECODED").RefersToRange.Value, " /x /e/") = 0 Then
        ThisWorkbook.Names("COMMAND_LINE_DECODED").RefersToRange.Value = vbNullString
        Exit Sub
    End If
    
    Call ParseArgs(str) ' get command line parameters to named ranges
        
    If Not Set_Global_Settings Then GoTo ErrHandler
    If Not Check_Main_Parameters Then GoTo ErrHandler
    
    If Right(ThisWorkbook.Names("SETTINGS_TARGET_PATH").RefersToRange.Value, 1) = "\" Or _
            Right(ThisWorkbook.Names("SETTINGS_TARGET_PATH").RefersToRange.Value, 1) = "/" Then
        Call Refresh_Folder
    Else
        'Stop
        Call Refresh_File
    End If

Exit_Sub:
    ' TODO: write to log following parameters (if debug_mode or log enabled)
    '         : Name : "PeakPageFileUsage" : String : ThisWorkbook.Workbook_Open
    '         : Name : "WorkingSetSize" : String : ThisWorkbook.Workbook_Open
    ' VirtualSize
    '         : Name : "PeakWorkingSetSize" : String : ThisWorkbook.Workbook_Open
    On Error Resume Next
    Application.DisplayAlerts = False
    Application.Quit
    
    Exit Sub
    
ErrHandler:
    Call Write_Log("Refresh Failed", bMandatoryLogRecord)
    GoTo Exit_Sub
    Resume ' test purpose
End Sub
