Attribute VB_Name = "Process"
Option Explicit

Function CheckProcessExist(ProcessID)
    Dim objWMIService, colProcess, process
    On Error GoTo ErrHandler
    
    ' -1 gives "wrong request"
    If ProcessID = -1 Then Exit Function
    
    Set objWMIService = GetObject("winmgmts:" _
        & "{impersonationLevel=impersonate}!\\.\root\cimv2")
     
    Set colProcess = objWMIService.ExecQuery _
        ("Select * from Win32_Process Where ProcessID = " & ProcessID)
    
    For Each process In colProcess
        CheckProcessExist = 1
        Exit For
    Next

ErrHandler:
    Set process = Nothing
    Set colProcess = Nothing
    Set objWMIService = Nothing
End Function

' Kill child processes
' http://stackoverflow.com/questions/20379723/how-to-kill-child-processes-with-vbscript

Sub KillProcessWithChildren(ProcessID)
    Dim objWMIService, objProcess, colProcess
    On Error Resume Next
    
    Set objWMIService = GetObject("winmgmts:" _
        & "{impersonationLevel=impersonate}!\\.\root\cimv2")
    
    ' collection of children
    Set colProcess = objWMIService.ExecQuery _
        ("Select * from Win32_Process Where ParentProcessId = " & ProcessID)
    
    For Each objProcess In colProcess
        objProcess.Terminate
    Next
    
    ' main process
    Set colProcess = objWMIService.ExecQuery _
        ("Select * from Win32_Process Where ProcessId = " & ProcessID)
    
    For Each objProcess In colProcess
        objProcess.Terminate
    Next
    
    Set colProcess = Nothing
    Set objWMIService = Nothing
    Set objProcess = Nothing
End Sub

Sub Process_Killer(ProcessID)
    Dim objWMIService, objProcess, colProcess
    On Error Resume Next
    
    Set objWMIService = GetObject("winmgmts:" _
        & "{impersonationLevel=impersonate}!\\.\root\cimv2")
    
    ' collection
    Set colProcess = objWMIService.ExecQuery _
        ("Select * from Win32_Process Where ProcessID = " & ProcessID)
    
    For Each objProcess In colProcess
        objProcess.Terminate
    Next
    Set colProcess = Nothing
    Set objWMIService = Nothing
    Set objProcess = Nothing
End Sub
