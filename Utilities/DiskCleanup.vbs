' http://ss64.com/vb/syntax-profile.html

' cscript DiskCleanUp.vbs //nologo
' cscript "C:\MySoft\DiskCleanup.vbs"

Option Explicit

'Variables
Dim objShell,FSO,dtmStart,dtmEnd
Dim strUserProfile,strAppData
Dim objFolder,objFile,strOSversion

Wscript.echo "Profile cleanup starting"
dtmStart = Timer()

'Get the current users Profile and ApplicationData folders
Set objShell = CreateObject("WScript.Shell")
strUserProfile = objShell.ExpandEnvironmentStrings("%USERPROFILE%")
strAppData = objShell.ExpandEnvironmentStrings("%APPDATA%")
'Wscript.echo strAppData

'Set reference to the file system
Set FSO = createobject("Scripting.FileSystemObject")

'Get the windows version
strOSversion = objShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\CurrentVersion")
'Wscript.echo strOSversion
'Call the DeleteOlder subroutine for each folder

'Application temp files

DeleteOlder 14, strAppData & "\Microsoft\Office\Recent" 'Days to keep recent MS Office files
DeleteOlder 5, strAppData & "\Microsoft\CryptnetUrlCache\Content"  'IE certificate cache
DeleteOlder 5, strAppData & "\Microsoft\CryptnetUrlCache\MetaData" 'IE cert info
DeleteOlder 5, strAppData & "\Sun\Java\Deployment\cache" 'Days to keep Java cache
DeleteOlder 5, strAppData & "\Macromedia\Flash Player"   'Days to keep flash data

DeleteOlderWithSubFolders 2, strUserProfile & "\AppData\Local\Temp\"   'Days to keep temp

'OS specific temp files
if Cint(Left(strOSversion,1)) > 5 Then
   Wscript.echo "Windows Vista/7/2008..."
   DeleteOlder 90, strAppData & "\Microsoft\Windows\Cookies"  'Days to keep cookies
   DeleteOlder 14, strAppData & "\Microsoft\Windows\Recent"   'Days to keep recent files
Else
   Wscript.echo "Windows 2000/2003/XP..."
   DeleteOlder 90, strUserProfile & "\Cookies"  'Days to keep cookies
   DeleteOlder 14, strUserProfile & "\Recent"   'Days to keep recent files
End if

'Print completed message

dtmEnd = Timer()
Wscript.echo "Profile cleanup complete, elapsed time: " & FormatNumber(dtmEnd-dtmStart,2) & " seconds"

'Subroutines below

' The runtime error of fso.DeleteFolder is caused by pathlengths > 259 characters. 
' If such a long path exists DeleteFolder will only say "Path not found". 
' After shortening the long paths manually everythings works (including on UNC Paths). 

Sub DeleteOlder(intDays,strPath)
' Delete files from strPath that are more than intDays old
If FSO.FolderExists(strPath) = True Then
   on error resume next
   Set objFolder = FSO.GetFolder(strPath)
   For each objFile in objFolder.files
      If DateDiff("d", objFile.DateLastModified,Now) > intDays Then
         Wscript.echo "File: " & objFile.Name
         objFile.Delete(True)
      End If
   Next
End If
End Sub

Sub DeleteOlderWithSubFolders(intDays,strPath)
Dim objSubFolder
Dim strRun

' Delete files from strPath that are more than intDays old
' Wscript.Echo "Starting: " & strPath
If FSO.FolderExists(strPath) = True Then   
   Set objFolder = FSO.GetFolder(strPath)
   
	Wscript.echo "Folder: " & objFolder.path
	Wscript.echo "Len: " & Len(objFolder.path)
	ON error resume next
	Wscript.echo "Subfolders: " & objFolder.SubFolders.Count
	Wscript.echo "Files: " & objFolder.Files.Count
	On error goto 0
	
	'Wscript.echo "Folder: " & objFolder.name			  	
	' For each objFile in objFolder.files
		' If DateDiff("d", objFile.DateLastModified,Now) > intDays Then
			' Wscript.echo "Deleting file: " & objFile.Path
			' on error resume next
			' objFile.Delete(True)
			' if err.number <> 0 then
				' Wscript.echo "Cannot delete file: " & Err.number & " : " & Err.Description
				' err.clear
			' else
				' Wscript.echo Chr(13) & "Deleted" & Chr(13)
			' end if
			' on error goto 0
		' End If
	' Next ' objFile				
	
	if left(objFolder.name, 9) = "Vertipaq_" then
		Wscript.echo "Deleting folder (vertipaq): " & objFolder.Path & "\"
		on error resume next
		FSO.deletefolder objFolder.path, true
		if err.number <> 0 then
			Wscript.echo "Cannot delete folder: " & Err.number & " : " & Err.Description
			err.clear
		end if
	else
		' not vertipaq
		if objFolder.SubFolders.Count = 0 and objFolder.Files.Count = 0 then
			Wscript.echo "Deleting folder (no sub-items): " & objFolder.Path & "\"
			on error resume next
			FSO.DeleteFolder objFolder.Path, true
			if err.number <> 0 then
				Wscript.echo "Cannot delete folder: " & Err.number & " : " & Err.Description
				err.clear
			else 
				Wscript.echo Chr(13) & "Deleted" & Chr(13)
			end if
			on error goto 0
		else
			' delete by time	
			If DateDiff("d", objFolder.DateLastModified,Now) > intDays Then
				 Wscript.echo "Deleting folder (time): " & objFolder.Name
				 Wscript.echo "Attributes: " & objFolder.Attributes
				 objFolder.Attributes = 0
				 Wscript.echo "Attributes2: " & objFolder.Attributes
				 on error resume next
				 FSO.DeleteFolder objFolder.path & "\", true
				 				
				'strRun = "cmd /c rd /s /q """ & strPath & """"
				'strRun = "cmd /k rd /s """ & strPath & """"
				strRun = "cmd rd /c /s """ & strPath & """"
				objShell.Run strRun, 1, True 

				 if err.number <> 0 then
					Wscript.echo "Cannot delete folder: " & Err.number & " : " & Err.Description
					err.clear
				else
					Wscript.echo Chr(13) & "Deleted" & Chr(13)
				end if
				on error goto 0
			End If
		end if
	end if
	
	If FSO.FolderExists(strPath) = True Then   
		for each objSubFolder in objFolder.SubFolders		
			'Call DeleteOlderWithSubFolders(intDays,objSubFolder.Path)				
			
			If DateDiff("d", objSubFolder.DateLastModified,Now) > intDays Then
				 Wscript.echo "Deleting folder (time): " & objSubFolder.Name				 
				 on error resume next
				 FSO.DeleteFolder objSubFolder.path & "\", true				 								
				
				strRun = "cmd /c del /s /q """ & objSubFolder.path & "\" & """"
				'strRun = "cmd /k del /s /q """ & objSubFolder.path & "\" & """"
				objShell.Run strRun, 1, True 
				
				strRun = "cmd /c rd /q /s """ & objSubFolder.path & "\" & """"
				'strRun = "cmd /k rd /q /s """ & objSubFolder.path & "\" & """"
				
				'strRun = "cmd /c rd /s """ & objSubFolder.path & "\" & """"
				'strRun = "cmd /k rd /s """ & objSubFolder.path & "\" & """"
				
				objShell.Run strRun, 1, True 

				 if err.number <> 0 then
					Wscript.echo "Cannot delete folder: " & Err.number & " : " & Err.Description
					err.clear
				else
					Wscript.echo Chr(13) & "Deleted" & Chr(13)
				end if
				on error goto 0
			End If
			
		next ' objSubFolder
	end if 
else
	Wscript.echo "Folder doesn't exist"
End If
End Sub