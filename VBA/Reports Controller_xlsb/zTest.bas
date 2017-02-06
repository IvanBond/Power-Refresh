Attribute VB_Name = "zTest"
Option Explicit

Sub Manual_Run_Test()
    Dim ReportPath As String
    Dim objShell, objScriptEngine, objProc As Object
    Dim custom_parameters As String
    
    Call Set_Global_Variables
    
    ReportPath = """" & ThisWorkbook.path & "\Refresher.xlsb"""
    
    'custom_parameters = Collect_Parameters
    
    Set objShell = CreateObject("WScript.Shell")
    Set objScriptEngine = CreateObject("scriptcontrol")
    objScriptEngine.Language = "JScript"
    
    ' Run Excel with switches /x /r /e
    ' about switches https://support.microsoft.com/en-us/kb/291288
    Set objProc = objShell.Exec(Excel_Path & " /r " & _
        ReportPath & " /x /e" & objScriptEngine.Run("encodeURIComponent", custom_parameters))
        
    ThisWorkbook.Sheets("LOG").Cells( _
        ThisWorkbook.Sheets("LOG").Cells(ThisWorkbook.Sheets("LOG").Rows.count, 3).End(xlUp).Row + 1, 3) = objProc.ProcessID
    
    Set objShell = Nothing
    Set objScriptEngine = Nothing
    Set objProc = Nothing
    
End Sub

