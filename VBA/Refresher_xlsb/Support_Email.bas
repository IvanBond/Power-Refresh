Attribute VB_Name = "Support_Email"
Option Explicit

Sub Send_Mail(bSuccess As Boolean, Msg_Text As String, Optional Scope As String)
    Dim oMyMail, iConf, Flds, szServer
        
    On Error GoTo ErrHandler
    Set oMyMail = CreateObject("CDO.Message")
    Set iConf = CreateObject("CDO.Configuration")
    Set Flds = iConf.Fields
    szServer = "http://schemas.microsoft.com/cdo/configuration/"
    
    With Flds
        .Item(szServer & "sendusing") = "2"
        .Item(szServer & "smtpserver") = ThisWorkbook.Names("SETTINGS_SMTP_SERVER").RefersToRange.Value
        .Item(szServer & "smtpserverport") = ThisWorkbook.Names("SETTINGS_SMTP_SERVER_PORT").RefersToRange.Value
        .Item(szServer & "smtpconnectiontimeout") = 100 ' quick timeout
        
        If ThisWorkbook.Names("SETTINGS_SMTP_SERVER_USERNAME").RefersToRange.Value <> vbNullString Then
            .Item(szServer & "smtpauthenticate") = 1
        Else
            .Item(szServer & "smtpauthenticate") = 0
        End If
                
        .Item(szServer & "sendusername") = ThisWorkbook.Names("SETTINGS_SMTP_SERVER_USERNAME").RefersToRange.Value
        .Item(szServer & "sendpassword") = ThisWorkbook.Names("SETTINGS_SMTP_SERVER_PASSWORD").RefersToRange.Value
        .Update
    End With
    
    With oMyMail
            Set .Configuration = iConf
            .bodypart.Charset = "utf-8"

            .To = IIf(bSuccess, _
                        ThisWorkbook.Names("SETTINGS_SUCCESS_EMAIL_TO").RefersToRange.Value, _
                        ThisWorkbook.Names("SETTINGS_ERROR_EMAIL_TO").RefersToRange.Value)

            ' by default - send to owner of Reports Controller
            .To = IIf(.To = vbNullString, ThisWorkbook.Names("SETTINGS_EMAIL_FROM").RefersToRange.Value, .To)
            
            .From = ThisWorkbook.Names("SETTINGS_EMAIL_FROM").RefersToRange.Value
            
            .Subject = "Power Refresh on " & Environ$("computername") & ": " & ReportName & _
                IIf(Scope <> vbNullString, " " & Scope, "") & IIf(bSuccess, " - Success", " - Fail")
            
            If bSuccess Then
                .TextBody = "Process " & ProcessID & " succesfully finished."
            Else
                .TextBody = "Error during process " & ProcessID & ". " & Msg_Text
            End If
            
            If ThisWorkbook.Names("SETTINGS_DEBUG_MODE").RefersToRange.Value = "Y" Or _
                ThisWorkbook.Names("SETTINGS_LOG_ENABLED").RefersToRange.Value = "Y" Then

                On Error Resume Next
                .TextBody = .TextBody & vbCrLf & vbCrLf & Get_Log_Records_For_Process(CStr(ProcessID))
                
                ' .AddAttachment LogsFolderPath & ReportID & ".log" ' (1) option for logging process
                Err.Clear
                On Error GoTo 0
            End If
            
            On Error Resume Next
            .Send
            ' TOThink - what to do if error? write log?
            
            If Err.Number <> 0 Then
                Call Write_Log("Couldn't send email. " & Err.Number & ": " & Err.Description)
            End If
            On Error GoTo 0
    End With
    
    Exit Sub

ErrHandler:
    On Error GoTo 0
    Call Write_Log("Couldn't send email. " & Err.Number & ": " & Err.Description)
End Sub
