'****************** INSTRUCTION ******************
' Run this script with arguments
' First argument should be R or T
' R for Report
' T for Data Transfer
' Due to different logic of refresh for Report and Data Transfer workbooks we have to identify them
' For Data Transfer we provide only name of transferring object - name of workbook expected in Country folder 
' For Report - full path to report. Can be file on local drive, network drive, SharePoint etc. Any place that Excel can use for Workbook.Open method
' 
' samples: 
' 	Reports
' 	Refresher.vbs R "C:\Reports Folder\Report.xlsx"
' 		will refresh report in 2nd argument for all Scopes provided in workbook and save result in report folder
' 	Refresher.vbs R "C:\Reports Folder\Report.xlsx" KZ
' 		will refresh report in 2nd argument for Scope = KZ
'		then save to same folder with name "Report KZ.xlsx"
'
'	Transfers
' 	Refresher.vbs T "Customers"
' 		will refresh "Customers" transfer for all countries (loop through folders in DataTransfer_Folder)
' 	Refresher.vbs T "Customers" UA
' 		will refresh "Customers" transfer for UA
'	both results will be saved to DataTransfer_UpdatedPath
'
' * use quotes for arguments that contain spaces
'
' Scopes concept is defined in post: TODO - write a post
' also can be understood from sample Excel file attached to project - TODO
'
' Created On: 2015-10-25
' Author: Ivan Bondarenko
'****************** END OF INSTRUCTION ******************

Path="C:\Power Refresh\"
DataTransfer_Folder = "C:\Power Refresh\Data Transfer\" ' Expect subfolders with Source IDs
	' each SubFolder contains set of files
	' each file represents a workbook that query data from source and place output to worksheet Result
	' which then saved as separate file to Updated folder
	
DataTransfer_UpdatedPath = "C:\Power Refresh\Updated\"

LogsFolder = "C:\Power Refresh\Logs\"
Update_Macro_Text_Name = "Update Macro.txt" ' contains macro code for refresh Excel Connections / Data Model
WinAPI_Macro_Text_Name = "Declare WinAPI Macro.txt" ' contains code that declare WinAPI functions

Param_Tries_Qty = 3 ' on refresh fail script will try to refresh once again
Param_Delay_Between_Tries = 10 ' in minutes
Param_Delay_Paste_Data_On_Result_Sheet = 30 ' in seconds

' Emailing parameters
smtp_server = "smtp server"
ErrorNotification_SendFrom  = "Sender Email"
ErrorNotification_SendTo = "Recipients"

'===========================================================================================================================================================
' START OF SCRIPT
	Dim objExcel
	Dim ProcessID
	Dim objFSO
	Dim ReportName
	Dim ScopesDictionary
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	
	StartPoint = Timer()
	On Error Resume Next
	
	If WScript.Arguments.Count < 2 then
		WScript.Echo "You should call script with at least two arguments!"
		WScript.Quit
	end if
	
	If WScript.Arguments(0) <> "T" and WScript.Arguments(0) <> "t" and WScript.Arguments(0) <> "R" and WScript.Arguments(0) <> "r" then
		WScript.Echo "Expected R or T for first argument!"
		WScript.Quit
	end if
		
	ReportName = GetReportName()
	call Write_Log(ReportName & " *************************** START ***************************")
		
	WinAPI_Macro_Text = ReadTxt(Path & WinAPI_Macro_Text_Name)	
	Update_Macro_Text = ReadTxt(Path & Update_Macro_Text_Name)
	
	If Err.Number > 0 then
		call Write_Log(ReportName & " Couldn't load VBA Macro text")
		WScript.Quit
	End if	
	
	' if we refresh Report - it is always one report
	' Scopes listed in ControlTable on ControlPanel worksheet
	
	' if we refresh 'Transfer workbook' we usualy should refresh such workbook for every source
	' however, when one source failed we can launch script with a parameter manually (3rd argument)
	
	If WScript.Arguments(0) = "T" or WScript.Arguments(0) = "t" then
		' Data Transfer
		Set Folder = objFSO.GetFolder(DataTransfer_Folder)
		
		If WScript.Arguments.Count > 2 then
			' 3rd and following arguments - for manual call
			for i = 2 to WScript.Arguments.Count
				BeforeRefresher = Timer()
				call Write_Log( WScript.Arguments ( i ) & " " & ReportName & " # Start refresh.")
				call Refresher( DataTransfer_Folder & WScript.Arguments ( i ) & "\" & WScript.Arguments(1) , WScript.Arguments ( i ) )
				call Write_Log( WScript.Arguments ( i ) & " " & ReportName & " # Refresh finished. # " & FormatNumber( Int( (Timer() - BeforeRefresher) / 60 ), 0) & "m " & FormatNumber( (Timer() - BeforeRefresher) mod 60, 0) & "s")
			next
		else
			' if script was called without options - refresh all files
			For each SubFolder in Folder.Subfolders ' Folder = objFSO.GetFolder(DataTransfer_Folder)
				BeforeRefresher = Timer()
				call Write_Log( SubFolder.Name & " " & ReportName & " # Start refresh.")
				call Refresher( DataTransfer_Folder & SubFolder.Name & "\" & WScript.Arguments( 1 ), SubFolder.Name )
				call Write_Log( SubFolder.Name & " " & ReportName & " # Refresh finished. # " & FormatNumber( Int( (Timer() - BeforeRefresher) / 60 ), 0) & "m " & FormatNumber( (Timer() - BeforeRefresher) mod 60, 0) & "s")
			Next
		end if
	else
		' Report
		If WScript.Arguments.Count > 2 then					
			' 3rd and following arguments - for manual call - Scopes
			for i = 2 to WScript.Arguments.Count
				BeforeRefresher = Timer()
				call Write_Log( SubFolder.Name & " " & ReportName & " # Start refresh.")
				call Refresher( WScript.Arguments( 1 ), WScript.Arguments( i ) )
				call Write_Log( SubFolder.Name & " " & ReportName & " # Refresh finished. # " & FormatNumber( Int( (Timer() - BeforeRefresher) / 60 ), 0) & "m " & FormatNumber( (Timer() - BeforeRefresher) mod 60, 0) & "s")
			next
		else
			' no additional arguments - refresh all Scopes in ControlTable			
			call Refresher( WScript.Arguments( 1 ) , "")
		end if
	end if ' If WScript.Arguments(0) = "T" then
	
	call Write_Log(ReportName & " # Overall execution time: " & FormatNumber( Int( (Timer() - StartPoint) / 60 ), 0) & "m " & FormatNumber( (Timer() - StartPoint) mod 60, 0) & "s")
	call Write_Log(ReportName & " *************************** END ***************************" )		
	
' END OF SCRIPT

'===========================================================================================================================================================
Sub Refresher( File_Path, Scope )
	On Error Resume Next
	try = 1
	If WScript.Arguments(0) = "T" or WScript.Arguments(0) = "t" then
		' **************************************************** Data Transfer ****************************************************
		do while try <= Param_Tries_Qty
			call Write_Log( iif( Scope <> "", Scope & "_", "") & ReportName & " # Starting Try " & try)
			BeforeAction = Timer()
						
			result = iif( Refresh_T( File_Path, Scope ) , "Success", "Fail" )
			
			call Write_Log( iif( Scope <> "", Scope & "_", "") & ReportName & " # Try " & try & " finished with " & result & " " & FormatNumber( Int( (Timer() - BeforeAction) / 60 ), 0) & "m " & FormatNumber( (Timer() - BeforeAction) mod 60, 0) & "s")
			
			if result = "Success" then
				with objExcel
					.DisplayAlerts = false
					save_name = Replace( Replace( Replace( ReportName, ".xlsx", ""), ".xlsb", ""), ".xlsm", "") & iif( Scope <> "", " " & Scope, "")
					
					call Write_Log( iif( Scope <> "", Scope & "_", "") & ReportName & " # Copying sheet Result to new workbook")
					BeforeAction = Timer()					
					.Workbooks(1).Sheets("Result").Copy
					Wscript.Sleep 1000 * 5 ' just in case
					call Write_Log( iif ( Scope <> "", Scope & "_", "") & ReportName & " # Sheet has been copied. # " & FormatNumber(Timer() - BeforeAction, 0) & "s")					
					
					call Write_Log( iif( Scope <> "", Scope & "_", "") & ReportName & " # Saving workbook to " & DataTransfer_UpdatedPath & save_name & ".xlsx")
					BeforeAction = Timer()					
					.ActiveWorkBook.SaveAs DataTransfer_UpdatedPath & save_name & ".xlsx", 51					
					
					if Err.Number <> 0 then
						call Write_Log( iif ( Scope <> "", Scope & "_", "") & ReportName & " # Save failed. Error " & Err.Number & " " & Err.Description )
						Process_Killer(ProcessID)
						Exit Do
					end if
					
					call Write_Log( iif ( Scope <> "", Scope & "_", "") & ReportName & " # Workbook saved. # " & FormatNumber(Timer() - BeforeAction, 0) & "s")
					
				end with
				Process_Killer(ProcessID)
				Exit Do
			else 
				if try >= Param_Tries_Qty then
					call Write_Log( iif( Scope <> "", Scope & "_", "") & ReportName & " # Trying to send email")
					Call Send_Mail( iif ( Scope <> "", Scope & "_", ""), "ERROR", ReportName & " # Unable to refresh" )
				end if
			end if
			' kill Excel to clean up, new instance will be launched by Refresh_T
			Process_Killer(ProcessID)
			
			try = try + 1
			if try < Param_Tries_Qty then
				call Write_Log( iif ( Scope <> "", Scope & "_", "") & ReportName & " # Waiting between tries. " & Param_Delay_Between_Tries & " min")
				Wscript.Sleep ( 1000 * 60 ) * Param_Delay_Between_Tries ' in minutes
			end if
		loop
	else 
		' **************************************************** Report ****************************************************		
		if Scope <> "" then
			call Refresh_Try( File_Path, Scope )
		else			
			result = GetReportScopes( File_Path )
			
			if result = 1 then				
				for each Key in ScopesDictionary
					call Refresh_Try( File_Path, Key )
				next
			elseif result = 2 then
				' no ControlTable - refresh all
				' open workbook - refresh all
				call Refresh_Try( File_Path, "" )
			end if
		end if
	end if	 ' If WScript.Arguments(0) = "T" or WScript.Arguments(0) = "t" then
end sub

'===========================================================================================================================================================
Sub Refresh_Try( File_Path, Scope )
	try = 1
	do while try <= Param_Tries_Qty
		call Write_Log( iif( Scope <> "", Scope & "_", "") & ReportName & " # Starting Try " & try)
		BeforeAction = Timer()
		
		result = iif( Refresh_R( File_Path, Scope ) , "Success", "Fail" )
		
		if result = "Success" then
			with objExcel
				.DisplayAlerts = false
				
				save_name =  Replace( Replace( Replace( ReportName, ".xlsx", ""), ".xlsb", ""), ".xlsm", "") & iif( Scope <> "", " " & Scope, "") & ".xlsx" ' all reports in xlsx
				Report_Folder = GetReportFolder()
				
				call Write_Log( iif( Scope <> "", Scope & "_", "") & ReportName & " # Saving workbook to " & Report_Folder & save_name)
				BeforeAction = Timer()
				.ActiveWorkBook.SaveAs Report_Folder & save_name, 51
				
				if Err.Number <> 0 then
					call Write_Log( iif ( Scope <> "", Scope & "_", "") & ReportName & " # Save failed. Error " & Err.Number & " " & Err.Description )
					Process_Killer(ProcessID)
					Exit Do
				end if
			end with
						
			Process_Killer(ProcessID)
			
			Exit Do
		else
			if try >= Param_Tries_Qty then
				Call Send_Mail( Scope, "ERROR", ReportName & " # Unable to refresh." )
			end if
		end if
		' kill Excel, new instance will be launched by Refresh_T
		Process_Killer(ProcessID)
		
		try = try + 1
		if try < Param_Tries_Qty then
			call Write_Log( iif ( Scope <> "", Scope & "_", "") & ReportName & " # Waiting between tries. " & Param_Delay_Between_Tries & " min")
			Wscript.Sleep ( 1000 * 60 ) * Param_Delay_Between_Tries
		end if
	loop
end sub

'===========================================================================================================================================================
Function GetReportScopes(File_Path)
	' 0 - Fail
	' 1 - ControlTable found, Scopes collected
	' 2 - ControlTable was not found
	On Error Resume Next
	StartFunction = Timer()
	
	if letObjExcel("") = 1 then
		with objExcel
			call Write_Log( ReportName & " # Opening workbook")
			BeforeAction = Timer()						
			
			' TODO
			' separate logic for ReadOnly - reports that have scopes
			.Workbooks.Open File_Path ', True, True
			call Write_Log( ReportName & " # Workbook opened. " & FormatNumber(Timer() - BeforeAction, 0) & "s")
			
			call Write_Log( ReportName & " # Checking ControlTable existence")
			BeforeAction = Timer()
			For i =1 to .Workbooks(1).Sheets.Count
				set sh = .Workbooks(1).Sheets(i)
				if Err.Number <> 0 then call Write_Log( ReportName & " # Loop on sheets: " & Err.Number & " " & Err.Description)
				
				For j=1 to sh.ListObjects.Count
					set lo = sh.ListObjects(j)
					
					if Err.Number <> 0 then call Write_Log( ReportName & " # Loop on ListObjects: " & Err.Number & " " & Err.Description)
					If lo.Name = "ControlTable" Then												
						
						call Write_Log( ReportName & " # Getting Scopes from ControlTable")
						Set ScopesDictionary = CreateObject("Scripting.Dictionary") ' globally defined variable
						if Err.Number <> 0 then call Write_Log( ReportName & " # Cannot create ScopesDictionary object. Error: " & Err.Number & " " & Err.Description)
						
						For cell = 1 to lo.ListColumns("Scope").DataBodyRange.Rows.Count
							if Err.Number <> 0 then call Write_Log( ReportName & " # failed to get cell row: " & Err.Number & " " & Err.Description)							
							
							If Not ScopesDictionary.Exists( lo.ListColumns("Scope").DataBodyRange.Cells(cell, 1).Value ) Then
								if Err.Number <> 0 then call Write_Log( ReportName & " # failed check element in dictionary: " & Err.Number & " " & Err.Description)
								ScopesDictionary.Add lo.ListColumns("Scope").DataBodyRange.Cells(cell, 1).Value, lo.ListColumns("Scope").DataBodyRange.Cells(cell, 1).Value
							End If
						Next
						
						if Err.Number <> 0 then call Write_Log( ReportName & " # Cannot create ScopesDictionary object. Error: " & Err.Number & " " & Err.Description)
						
						for each Key in ScopesDictionary.Keys
							list_of_scopes = list_of_scopes & " " & Key
						next
						
						call Write_Log( ReportName & " # Scopes collected " & trim( list_of_scopes ) )
						Process_Killer(ProcessID)
						ProcessID = ""
						GetReportScopes = 1 ' 1 - ControlTable found, Scopes collected
						Exit Function
					End If
				Next ' For j=1 to sh.ListObjects.Count
			Next ' For i =1 to .Workbooks(1).Sheets.Count
			
			Process_Killer(ProcessID)
			ProcessID = ""
			GetReportScopes = 2 ' ControlTable not found - no scopes
			call Write_Log( iif ( Scope <> "", Scope & "_", "") & ReportName & " # ControlTable not found " & FormatNumber( Int( (Timer() - StartFunction) / 60 ), 0) & "m " & FormatNumber( (Timer() - StartFunction) mod 60, 0) & "s")
		end with
	else
		' couldn't create Excel app - Fail
		GetReportScopes = 0
		call Write_Log( ReportName & " # Unable to create Excel Application. " )		
	end if
end Function

'===========================================================================================================================================================
Function Refresh_T(File_Path, Scope)
	' Scope is mandatory parameter
	On Error Resume Next
	StartRefreshT = Timer()
	
	if letObjExcel( Scope ) = 1 then
		with objExcel			
			call Write_Log( Scope & "_" & ReportName & " # Opening workbook")
			BeforeAction = Timer()
			.Application.Workbooks.Open File_Path ' , True, True
			call Write_Log( Scope & "_" & ReportName & " # Workbook opened. " & FormatNumber(Timer() - BeforeAction, 0) & "s")
			
			call Write_Log( iif ( Scope <> "", Scope & "_", "") & ReportName & " # Adding macro")
			BeforeAction = Timer()
			.Workbooks(1).VBProject.VBComponents.Add(1).CodeModule.AddFromString Update_Macro_Text
			call Write_Log( iif ( Scope <> "", Scope & "_", "") & ReportName & " # Macros has been embedded. " & FormatNumber(Timer() - BeforeAction, 0) & "s")			
			
			call Write_Log( iif ( Scope <> "", Scope & "_", "") & ReportName & " # Starting Refresh Connections" )
			BeforeAction = Timer()
						
			macro_result = .Run("UpdateConnections")			
			
			if macro_result = 0 then
				Write_Log( iif ( Scope <> "", Scope & "_", "") & ReportName & " # Failed to refresh")				
			end if
			
			' actual for Data Transfer
			if  macro_result = 1 then
				Wscript.Sleep 1000 * Param_Delay_Paste_Data_On_Result_Sheet  ' wait while Excel paste data on sheet
				
				if .workbooks(1).sheets("Result").ListObjects(1).DataBodyRange is Nothing  then
					Write_Log( iif ( Scope <> "", Scope & "_", "") & ReportName & " # Rows loaded: 0")
					' can be used some logic here
					' for example, send email
				else
					Write_Log( iif ( Scope <> "", Scope & "_", "") & ReportName & " # Rows loaded: " & .workbooks(1).sheets("Result").ListObjects(1).DataBodyRange.Rows.Count )
				end if
				
			end if
			
			Refresh_T = ( macro_result = 1 )
			call Write_Log( iif ( Scope <> "", Scope & "_", "") & ReportName & " # Refresh finished " & FormatNumber( Int( (Timer() - StartRefreshT) / 60 ), 0) & "m " & FormatNumber( (Timer() - StartRefreshT) mod 60, 0) & "s")
		end with
	else
		call Write_Log( iif ( Scope <> "", Scope & "_", "") & ReportName & " # Unable to create Excel Application. " )
		Call Send_Mail( Scope, "1547", ReportName & " # Unable to create Excel Application. " )
	end if
end Function

'===========================================================================================================================================================
Function Refresh_R(File_Path, Scope)
	' Scope is optional parameter for reports
	' usually script collects scopes from ControlTable
	On Error Resume Next
	StartRefreshR = Timer()
	
	if letObjExcel( Scope ) = 1 then
		with objExcel						
			
			call Write_Log( Scope & "_" & ReportName & " # Opening workbook")
			BeforeAction = Timer()
			.Application.Workbooks.Open File_Path ', True, True			
			Wscript.Sleep 1000 * 15  ' wait while Excel load everything
			call Write_Log( Scope & "_" & ReportName & " # Workbook opened. " & FormatNumber(Timer() - BeforeAction, 0) & "s")
			
			call Write_Log( iif ( Scope <> "", Scope & "_", "") & ReportName & " # Adding macro")
			BeforeAction = Timer()
			.Workbooks(1).VBProject.VBComponents.Add(1).CodeModule.AddFromString Update_Macro_Text
			call Write_Log( iif ( Scope <> "", Scope & "_", "") & ReportName & " # Macros has been embedded. " & FormatNumber(Timer() - BeforeAction, 0) & "s")									
			
			' Set Scope if provided
			if Scope <> "" then
				.Workbooks(1).Names("SCOPE").RefersToRange.Value = Scope
			end if
			
			call Write_Log( iif ( Scope <> "", Scope & "_", "") & ReportName & " # Starting Refresh Connections" )
			BeforeAction = Timer()
									
			macro_result = .Run("UpdateConnections")
			
			if macro_result = 0 then
				Write_Log( iif ( Scope <> "", Scope & "_", "") & ReportName & " # Failed to refresh")
			else
				Wscript.Sleep 1000 * 15  ' wait while Excel paste data on sheets
				.Calculate
				.CalculateUntilAsyncQueriesDone
				' waiting for cube formulas
				while .CalculationState <> 0 ' xlDone = 0 , xlCalculating = 1, xlPending = 2
					WScript.Sleep 1000
				wend
			end if

			Refresh_R = ( macro_result = 1 )
			call Write_Log( iif ( Scope <> "", Scope & "_", "") & ReportName & " # Refresh finished " & FormatNumber( Int( (Timer() - StartRefreshR) / 60 ), 0) & "m " & FormatNumber( (Timer() - StartRefreshR) mod 60, 0) & "s")
		end with
	else
		call Write_Log( iif ( Scope <> "", Scope & "_", "") & ReportName & " # Unable to create Excel Application. " )
		Call Send_Mail( Scope, "1547", ReportName & " # Unable to create Excel Application. " )
	end if
End Function

'===========================================================================================================================================================
Function letObjExcel( Scope )
	' Creates empty Excel Application and remember its system ProcessID
	On Error Resume Next
	
	call Write_Log( iif( Scope <> "", Scope & "_", "") & ReportName & " # Creating Excel Object" )	
	
	StartTime = Timer()
	set objExcel = CreateObject("Excel.Application")
	
	if Err.Number <> 0 then 
		call Write_Log( iif( Scope <> "", Scope & "_", "") & ReportName & " # Error " & Err.Number & " " & Err.Description)
		Exit Function
	end if
	
	with objExcel
		.Visible = True ' can be hidded as well
		call Write_Log( iif( Scope <> "", Scope & "_", "") & ReportName & " # Adding workbook")
		BeforeAction = Timer()
		.Workbooks.Add
		if Err.Number <> 0 then Exit Function
		call Write_Log( iif( Scope <> "", Scope & "_", "") & ReportName & " # Workbook has been added. " & FormatNumber(Timer() - BeforeAction, 0) & "s")
		
		call Write_Log( iif( Scope <> "", Scope & "_", "") & ReportName & " # Adding macro")
		BeforeAction = Timer()
		
		.Workbooks(1).VBProject.VBComponents.Add(1).CodeModule.AddFromString WinAPI_Macro_Text
		
		call Write_Log( iif( Scope <> "", Scope & "_", "") & ReportName & " # Macros has been embedded. " & FormatNumber(Timer() - BeforeAction, 0) & "s")
		if Err.Number <> 0 then Exit Function
		
		do while try <= Param_Tries_Qty
			call Write_Log( iif( Scope <> "", Scope & "_", "") & ReportName & " # Getting ProcessID")
			ProcessID = .Run("GetProcessID")			
			
			if ProcessID = "" then
				call Write_Log( iif( Scope <> "", Scope & "_", "") & ReportName & " # Try " & Try & " cannot get ProcessID. Error " & Err.Number & " " & Err.Description)
				' try to quit to clean up
				.Quit
				if try >= Param_Tries_Qty then
					letObjExcel = 0
					.DisplayAlerts = false
					
					Exit Function
				end if
				
				try = try + 1
			else 
				Exit do
			end if
		loop
		call Write_Log( iif( Scope <> "", Scope & "_", "") & ReportName & " # ProcessID = " & ProcessID )
		
		.DisplayAlerts = false
		.Workbooks(1).Close
	end with
	
	call Write_Log( iif( Scope <> "", Scope & "_", "") & ReportName & " # Excel Object has been created. Overall time: " & FormatNumber( Int( (Timer() - StartTime) / 60 ), 0) & "m " & FormatNumber( (Timer() - StartTime) mod 60, 0) & "s")
	letObjExcel = 1
end Function

'===========================================================================================================================================================
Sub Write_Log(str)
	On Error Resume Next
	const ForAppending = 8	
	LogFile = LogsFolder & "Log_" & ReportName & ".txt"
	If not objFSO.FileExists(LogFile) Then objFSO.CreateTextFile(LogFile)
	Set objLog = objFSO.OpenTextFile(LogFile, ForAppending)
	objLog.WriteLine(now() & "# " & str)
	objLog.Close
end sub

'===========================================================================================================================================================	
Function ReadTxt(path)
	Const ForReading = 1
	Set objTextFile = objFSO.OpenTextFile(path, ForReading)
	ReadTxt = objTextFile.ReadAll
	objTextFile.Close
End function

'===========================================================================================================================================================
Sub Send_Mail(Scope, ErrNumber, ErrDescription)
	Dim oMyMail
	Set oMyMail = CreateObject("CDO.Message")
	Set iConf = CreateObject("CDO.Configuration")
	Set Flds = iConf.Fields
	szServer = "http://schemas.microsoft.com/cdo/configuration/"
	
	With Flds
		.Item(szServer & "sendusing") = "2"
		.Item(szServer & "smtpserver") = smtp_server
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
            .To = ErrorNotification_SendTo
            .From = ErrorNotification_SendFrom
			.Subject = "Power Refresh: " & ReportName & " " & Scope
			.TextBody = ErrNumber & " " & ErrDescription			
			.AddAttachment LogsFolder & "Log_" & ReportName & ".txt"
			.Send
    End With
End Sub

'===========================================================================================================================================================	
Sub Process_Killer(ProcessID)
	Dim objWMIService, objProcess, colProcess			
	 
	Set objWMIService = GetObject("winmgmts:" _
		& "{impersonationLevel=impersonate}!\\.\root\cimv2")
	 
	Set colProcess = objWMIService.ExecQuery _
		("Select * from Win32_Process Where ProcessID = " & ProcessID )
	
	For Each objProcess in colProcess
		objProcess.Terminate()
		Call Write_Log( ReportName & " # Process " & ProcessID & " was killed.")
	Next
End Sub

'===========================================================================================================================================================	
Function GetReportName()
	str = WScript.Arguments( 1 )
	str = Right(str, Len(str) - InStrRev(str, "/", -1, vbTextCompare) ) ' web address
	str = Right(str, Len(str) - InStrRev(str, "\", -1, vbTextCompare) ) ' file system address
	GetReportName = Replace (str, "%20", " ")
end function

'===========================================================================================================================================================	
Function GetReportFolder()
	str = WScript.Arguments( 1 )
	if InStr(str, "/") > 0 then
		GetReportFolder = Left(str, InStrRev(str, "/", -1, vbTextCompare) ) ' web address
	else
		GetReportFolder = Left(str, InStrRev(str, "\", -1, vbTextCompare) ) ' file system address
	end if	
end function

'===========================================================================================================================================================	
Function iif(psdStr, trueStr, falseStr)
  if psdStr then
    iif = trueStr
  else 
    iif = falseStr
  end if
end function
