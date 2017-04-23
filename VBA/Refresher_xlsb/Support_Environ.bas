Attribute VB_Name = "Support_Environ"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As LongPtr
#Else
    Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
#End If

Sub GetCurrentProcess()
    Dim objWMIService, colProcess, process
    On Error GoTo ErrHandler
    Set objWMIService = GetObject("winmgmts:" _
        & "{impersonationLevel=impersonate}!\\.\root\cimv2")
    
    ' collection
    Set colProcess = objWMIService.ExecQuery _
        ("Select * from Win32_Process Where ProcessID = " & GetCurrentProcessId) ' winAPI call
    
    ' must be only one process
    For Each process In colProcess
        Set CurrentProcess = process
    Next

ErrHandler:
    Set process = Nothing
    Set colProcess = Nothing
    Set objWMIService = Nothing
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
        Call Write_Log(ReportName & " # Process " & ProcessID & " was killed.")
    Next
    Set colProcess = Nothing
    Set objWMIService = Nothing
End Sub

Function CheckProcessExist(ProcessID)
    Dim objWMIService, colProcess
    On Error GoTo ErrHandler
    Set objWMIService = GetObject("winmgmts:" _
        & "{impersonationLevel=impersonate}!\\.\root\cimv2")
    
    ' collection
    Set colProcess = objWMIService.ExecQuery _
        ("Select * from Win32_Process Where ProcessID = " & ProcessID)
    
    If Not IsEmpty(colProcess) Then
        CheckProcessExist = colProcess.count
    End If
    
ErrHandler:
    Set colProcess = Nothing
    Set objWMIService = Nothing
End Function


' objShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )
' Environ$("computername")

' user's email
' Set objSysInfo = CreateObject("ADSystemInfo")
' Set objUser = GetObject("LDAP://" & objSysInfo.UserName)
' objUser.mail ' sends email to logged user

Function GetReportFolder() As String
    Dim str As String
    str = ThisWorkbook.Names("SETTINGS_TARGET_PATH").RefersToRange.Value
    If InStr(str, "/") > 0 Then
        GetReportFolder = Left(str, InStrRev(str, "/", -1, vbTextCompare))  ' web address
    Else
        GetReportFolder = Left(str, InStrRev(str, "\", -1, vbTextCompare))  ' file system address
    End If
End Function

Function GetReportName() As String
    Dim str As String
    str = ThisWorkbook.Names("SETTINGS_TARGET_PATH").RefersToRange.Value
    str = Right(str, Len(str) - InStrRev(str, "/", -1, vbTextCompare))  ' web address
    str = Right(str, Len(str) - InStrRev(str, "\", -1, vbTextCompare))  ' file system address
    str = Left(str, InStrRev(str, ".", -1, vbTextCompare) - 1) ' remove extension
    GetReportName = Replace(str, "%20", " ")
    ' TOThink: use decodeURL here. Low prio
End Function
