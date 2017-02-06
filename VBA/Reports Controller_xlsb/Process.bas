Attribute VB_Name = "Process"
Option Explicit

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

' Kill child processes
' http://stackoverflow.com/questions/20379723/how-to-kill-child-processes-with-vbscript

