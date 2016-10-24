Option Explicit
REM DESCRIPTION
	VBScript starts Excel application and store ProcessID
	Waits certain time, then check if ProcessID still exists (very small chace that another application will be executed with this ID)
	If process exists - informs user by email.
REM END OF DESCRIPTION


REM END OF SCRIPT
	Dim objProc
	Dim objShell
	Dim XLapp
	Dim ReportPath
	Dim objSysInfo, objUser
	Dim Email_To
	Dim Email_From		
	
	Const SmtpServer = "yourmailhost.com"
	
	' if call VBS with parameter - uncomment
	'ReportPath = """" & Wscript.Arguments(0) & """"
	
	ReportPath = """C:\Temp\Test.xlsx"""
	
	set objShell = CreateObject("WScript.Shell")
	Set objSysInfo = CreateObject("ADSystemInfo")
	Set objUser = GetObject("LDAP://" & objSysInfo.UserName)
	
	' if you run script under own account
	Email_From = objUser.mail ' "email@mail.com" ' use static email if needed
	Email_To = objUser.mail ' "email@mail.com"
	
	' Run Excel with switches /x /r /e
	' about switches https://support.microsoft.com/en-us/kb/291288
	' check your path to Excel - can be different ( depends on version of Excel )
	set objProc = objShell.Exec("C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE /x /e /r " & ReportPath)

	' delay until check if process exists
	Wscript.Sleep 60000 * 15 ' 15 minutes
	'Wscript.Sleep 10000 ' test time
	
	' if process still exists - something is wrong
	if CheckProcessExist(objProc.ProcessID) > 0 then
		' Wscript.Echo "Have to kill object " & objProc.ProcessID
		call Send_Email("Something went wrong on " & objShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" ) )
		'call Process_Killer( objProc.ProcessID )
	else
		' everything went fine	
	end if
	
	set objProc = nothing
	set objShell = nothing
	set objSysInfo = nothing
	set objUser = nothing 
REM END OF SCRIPT

Function CheckProcessExist(ProcessID)
	Dim objWMIService, colProcess
	
	Set objWMIService = GetObject("winmgmts:" _
		& "{impersonationLevel=impersonate}!\\.\root\cimv2")
	 
	Set colProcess = objWMIService.ExecQuery _
		("Select * from Win32_Process Where ProcessID = " & ProcessID )
	CheckProcessExist = colProcess.count
	Set colProcess = nothing
	Set objWMIService = nothing
end function

Sub Process_Killer(ProcessID)
	Dim objWMIService, objProcess, colProcess
	On Error Resume Next
	
	Set objWMIService = GetObject("winmgmts:" _
		& "{impersonationLevel=impersonate}!\\.\root\cimv2")
	 
	Set colProcess = objWMIService.ExecQuery _
		("Select * from Win32_Process Where ProcessID = " & ProcessID )
	
	For Each objProcess in colProcess
		objProcess.Terminate()
		Call Write_Log( ReportName & " # Process " & ProcessID & " was killed.")
	Next
	Set colProcess = nothing
	Set objWMIService = nothing
End Sub

Sub Send_EMail(subj)
	Dim oMyMail
	Dim iConf, Flds
	Dim szServer
	
	Set oMyMail = CreateObject("CDO.Message")
	Set iConf = CreateObject("CDO.Configuration")
	Set Flds = iConf.Fields
	szServer = "http://schemas.microsoft.com/cdo/configuration/"
	
	With Flds
		.Item(szServer & "sendusing") = "2"
		.Item(szServer & "smtpserver") = SmtpServer
		.Item(szServer & "smtpserverport") = "25"
		.Item(szServer & "smtpconnectiontimeout") = 100 ' quick timeout
		.Item(szServer & "smtpauthenticate") = 0
		.Item(szServer & "sendusername") = ""
		.Item(szServer & "sendpassword") = ""
		.Update
	End With
	
	With oMyMail			
            Set .Configuration = iConf
            .bodypart.Charset = "utf-8"
            .To = Email_To			
            .From = Email_From
			.Subject = subj
			.TextBody = ""			
			.Send
    End With
End Sub
