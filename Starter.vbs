' VBScript sample

Refresher_Path = "C:\Power Refresh\Refresher.xlsb"
Excel_Path = "C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"

Set objScriptEngine = CreateObject("ScriptControl")
objScriptEngine.Language = "JScript"

strParameters = "/report_id:Test Parameters/target_path:C:\Power Refresh\Tests\Test Parameters.xlsx/type:R/skip_refresh_all/save_inplace/parameters:From=vbs, TO=script"

Set objShell = CreateObject("WScript.Shell")
Set objProc = objShell.Exec(Excel_Path & " /x " & _
                        "/e" & objScriptEngine.Run("encodeURIComponent", strParameters) & _
                        " /r """ & Refresher_Path & """")

' check if process still exists
do while CheckProcessExist(objProc.ProcessID) = 1
	' WScript.Echo "Process Exists: " & objProc.ProcessID
	WScript.Sleep 5000
loop
						
set objScriptEngine = nothing
set objShell = nothing
set objProc = nothing

Function CheckProcessExist(ProcessID)
    Dim objWMIService, colProcess, process
    
    If ProcessID = -1 Then Exit Function
    
    Set objWMIService = GetObject("winmgmts:" _
        & "{impersonationLevel=impersonate}!\\.\root\cimv2")
     
    Set colProcess = objWMIService.ExecQuery _
        ("Select * from Win32_Process Where ProcessID = " & ProcessID)
    
    For Each process In colProcess
        CheckProcessExist = 1
        Exit For
    Next
    
    Set process = Nothing
    Set colProcess = Nothing
    Set objWMIService = Nothing
End Function

' relative path
'Refresher_Path = replace( WScript.ScriptFullName, WScript.ScriptName, "" ) & "Refresher.xlsb"

' Possible parameters
' "/report_id:" ' required, for log
' "/target_path:" ' required, path to target file / folder
' "/type:" required, R or T ' R-eport or T-ransfer
' "/macro_before:" ' optional, name of macro that will be executed before RefreshAll
' "/skip_refresh_all" ' optional, if you need to skip RefreshAll operation (by default will be executed)
' "/macro_after:" ' optional, name of macro that wiil be executed after RefreshAll
' "/error_email_to:" ' optional, email address to which Refresher will send a note in case of fail
' "/success_email_to:" ' optional, email address to which Refresher will send a note in case of success
' "/log_enabled" ' optional, if you need log

' Saving
' "/save_inplace" ' optional, if you want to save file after refresh inplace
' "/do_not_save" ' optional, if you don't want to save target file(s)
' "/result_folder_path:" ' optional, if you want to save file to a new location after refresh
' ' Replace(str, "/", "{|}")
' cheap solution - replace '/' before transfer parameters
                ' solves problem of web folder
                ' '|' cannot be used in path, so it will be replaced back in Refresher

' "/result_filename:" ' optional, if you need new file name
' "/add_datetime" ' optional, if you want to add datetime to your file
' "/extension:" ' optional, if you want to change extension: xlsx, xlsm, xlsb, csv (then provide /save_sheet)
' "/time_limit:" ' optional, max time for report, make sense when Refresher will start another instance of Excel to run Scopes / Files, in parallel / subsequently
' "/scopes_in_parallel" ' optional, if you want to run scopes in parallel
' "/files_in_parallel" ' optional, if you want to execute refresh of files in folder in parallel
' "/save_sheet:" ' optional, name of worksheet that you want to save, if you want to save only one worksheet
' "/parameters:" ' optional, vales for NamedRanges, e.g. FROM=12.12.2016, TO=31.12.2016
' Replace(str, "/", "{|}")
	' ' potential risk of appear '/'
            ' changed back on Refresher side during params parsing
' 
' "/scopes:" ' optional, scopes, separated by comma.
' 	Replace(str, "/", "{|}")
' 

' ideas for new parameters
' /formulas_to_values ' if you don't want to save formulas - formulas on all worksheets will be converted to values
' /email_method ' outlook, CDO, gmail
	' for CDO and Gmail - username and password? Or store them in Refresher.xlsb ?
' /add_datetime - change to "/add_datetime:" - parameter with value, e.g. 'yyyy-mm-dd', 'yy mm dd', 'yyyy-mm-dd hh:mi:ss' etc.
	' replace '/' in case of usage in format m/d/yyyy

' sample
' strParameters = "/debug_mode/log_enabled/file_path:C:\Temp\Test.xlsx/" & _
'        "result_folder_path:\\server_name\ssis in\/macro_before:test/macro_after:after" & _
'        "/error_email_to:test@domain.com/success_email_to:test@domain.com/result_filename:reportX/extension:xlsb" & _
'        "/add_datetime/scopes:KZ HR,UA AMS/parameters:FROM=2016-01-01,TO=2016-10-31"