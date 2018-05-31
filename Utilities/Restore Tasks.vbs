rem Script to restore Windows Task Scheduler tasks
rem from backup - copy of tasks from C:\Windows\system32\tasks\

TasksFolder = "C:Power Refresh\Tasks\" ' folder with backup of tasks without extension (how Task Scheduler stores them)
RestoreTasksToFolder = MyTasks" ' folder in Task Scheduler

Set oShell = WScript.CreateObject ("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set Fldr = objFSO.GetFolder(TasksFolder)

ProcessBackupTasksFolder Fldr, RestoreTasksToFolder

Sub ProcessBackupTasksFolder(OfFolder, TaskFolderPath)
    For Each SubFolder In OfFolder.SubFolders
        ProcessBackupTasksFolder SubFolder, TaskFolderPath & "\" & SubFolder.Name
    Next
    
	for each fl in ofFolder.Files
		' backup files do not contain extension
		' rename - for testing
		'if Instr(fl.name, ".") = 0 then
		'	objFSO.MoveFile fl, fl & ".xml"
		'end if
		
		' schtasks help
		' http://ss64.com/nt/schtasks.html
		' test 		
		'oShell.Run "cmd.exe /K schtasks /create /xml """ & fl.path & """ /tn """ & TaskFolderPath & "\" & Replace( fl.name, ".xml", "" ) & """"
		'WScript.echo "cmd.exe /K schtasks /create /xml """ & fl.path & """ /tn """ & TaskFolderPath & "\" & Replace( fl.name, ".xml", "" ) & """"
		
		' cmd.exe switches:
		' http://ss64.com/nt/cmd.html
		' prod
		
		oShell.Run "cmd.exe /C schtasks /create /xml """ & fl.path & """ /tn """ & TaskFolderPath & "\" & Replace( fl.name, ".xml", "" ) & """"
	next
    
End Sub