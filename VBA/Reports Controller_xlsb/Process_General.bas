Attribute VB_Name = "Process_General"
Option Explicit

' Kill child processes
' http://stackoverflow.com/questions/20379723/how-to-kill-child-processes-with-vbscript

Sub KillProcessWithDependents(ProcessID As String)
    Dim objWMIService, objProcess, colProcess
    
    On Error GoTo ErrHandler
    
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

Exit_Sub:
    On Error Resume Next
    Set colProcess = Nothing
    Set objWMIService = Nothing
    Set objProcess = Nothing
    Err.Clear
    Exit Sub
    
ErrHandler:
    Debug.Print Now, "KillProcessWithDependents", Err.Number & ": " & Err.Description
    Err.Clear
    GoTo Exit_Sub
    Resume
End Sub

Function GetRunningProcessesCountByName(sName As String) As Integer
' Function returns number of running processes with specified name
' called for 'excel.exe' returns number of running Excel processes
' to prevent launch overlimit number of processes

    Dim objWMIService As Object
    Dim colItems As Object
    Dim objItem As Object
    Dim i As Integer
    
    On Error GoTo ErrHandler
    
    Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
    
    Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_Process" & _
               " WHERE Name = '" & sName & "'", , 48)
    
    i = 0
    For Each objItem In colItems
        i = i + 1
    Next objItem
    GetRunningProcessesCountByName = i

Exit_Function:
    On Error Resume Next
    Set objWMIService = Nothing
    Set colItems = Nothing
    Set objItem = Nothing
    Err.Clear
    Exit Function
    
ErrHandler:
    Debug.Print Now, "GetRunningProcessesCountByName", Err.Number & ": " & Err.Description
    Err.Clear
    GoTo Exit_Function
End Function

Function GetProcessStartTime(ProcessID As String) As Date
    Dim objWMIService, objProcess, colProcess
    On Error Resume Next
    
    On Error GoTo ErrHandler
    
    Set objWMIService = GetObject("winmgmts:" _
        & "{impersonationLevel=impersonate}!\\.\root\cimv2")
    
    Set colProcess = objWMIService.ExecQuery _
        ("Select * from Win32_Process Where ProcessId = " & ProcessID)
    
    For Each objProcess In colProcess
        'objProcess.Terminate
        GetProcessStartTime = DateValue(WMIDateStringToDateTime(objProcess.CreationDate)) + _
                TimeValue(WMIDateStringToDateTime(objProcess.CreationDate))
        GoTo Exit_Sub
    Next

Exit_Sub:
    On Error Resume Next
    Set colProcess = Nothing
    Set objWMIService = Nothing
    Set objProcess = Nothing
    Err.Clear
    Exit Function
    
ErrHandler:
    Debug.Print Now, "GetProcessStartTime", Err.Number & ": " & Err.Description
    Err.Clear
    GoTo Exit_Sub
    Resume
End Function

Function WMIDateStringToDateTime(str As String) As String
' https://technet.microsoft.com/en-au/library/ee198928.aspx
    WMIDateStringToDateTime = Left(str, 4) & "-" & _
        Mid(str, 5, 2) & "-" & _
        Mid(str, 7, 2) & " " & _
        Mid(str, 9, 2) & ":" & _
        Mid(str, 11, 2) & ":" & _
        Mid(str, 13, 2)
End Function

Function CheckProcessExist(ProcessID As String) As Boolean
    Dim objWMIService, colProcess, process
    On Error GoTo ErrHandler
    
    ' -1 gives "wrong request"
    If ProcessID = -1 Then Exit Function
    
    Set objWMIService = GetObject("winmgmts:" _
        & "{impersonationLevel=impersonate}!\\.\root\cimv2")
     
    Set colProcess = objWMIService.ExecQuery _
        ("Select * from Win32_Process Where ProcessID = " & ProcessID)
    
    For Each process In colProcess
        CheckProcessExist = True
        Exit For
    Next

Exit_Sub:
    On Error Resume Next
    Set process = Nothing
    Set colProcess = Nothing
    Set objWMIService = Nothing
    Err.Clear
    Exit Function
    
ErrHandler:
    Debug.Print Now, "CheckProcessExist", Err.Number & ": " & Err.Description
    Err.Clear
    GoTo Exit_Sub
    Resume
End Function

