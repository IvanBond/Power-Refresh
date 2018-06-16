Attribute VB_Name = "Support_Email"
Option Explicit
Option Compare Text

Enum enumMailImportance
    Low = 0
    Normal = 1
    High = 2
End Enum

Sub SendNotification(sSubject As String, _
                     sMessageText As String, _
                     Optional Report_Row_ID As Long)
    
    If [SETTINGS_EMAIL_ERRORS_TO].Value <> vbNullString Then

        If [SETTINGS_EMAIL_METHOD].Value = "Outlook" Then

            Call Send_Email_Outlook([SETTINGS_EMAIL_ERRORS_TO].Value, _
                 sSubject, _
                sMessageText, _
                IIf([SETTINGS_EMAIL_ATTACH_LOGFILE].Value = "Y", _
                    IIf(Report_Row_ID <> 0, _
                        ThisWorkbook.path & "\Logs\" & GetReportParameter(Report_Row_ID, "Report ID *") & ".log", _
                        vbNullString), _
                    vbNullString), _
                IIf([SETTINGS_EMAIL_IMPORTANCE].Value <> vbNullString, [SETTINGS_EMAIL_IMPORTANCE].Value, "Normal"))

        ElseIf [SETTINGS_EMAIL_METHOD].Value = "SMTP" Then
            
            If CheckSMTPSettings Then
                Call Send_EMail_CDO([SETTINGS_SMTP_FROM].Value, _
                    [SETTINGS_EMAIL_ERRORS_TO].Value, _
                    sSubject, _
                    sMessageText, _
                    IIf([SETTINGS_EMAIL_ATTACH_LOGFILE].Value = "Y", _
                        IIf(Report_Row_ID <> 0, _
                            ThisWorkbook.path & "\Logs\" & GetReportParameter(Report_Row_ID, "Report ID *") & ".log", _
                            vbNullString), _
                        vbNullString), _
                    IIf([SETTINGS_EMAIL_IMPORTANCE].Value <> vbNullString, [SETTINGS_EMAIL_IMPORTANCE].Value, "Normal"))
            End If ' If CheckSMTPSettings Then
            
        End If ' If [SETTINGS_EMAIL_METHOD].Value = "Outlook" Then

    End If ' If [SETTINGS_EMAIL_ERRORS_TO].Value <> vbNullString Then
End Sub

Function Send_EMail_CDO(sFrom As String, _
                    sRecipients As String, _
                    sSubject As String, _
                    Optional sMessage As String, _
                    Optional sAttachmentPath As String, _
                    Optional Importance As String = "Normal", _
                    Optional sCC As String, _
                    Optional sBCC As String)
                    
    Dim iMsg As Object
    Dim iConf As Object
    Dim strbody As String
    Dim sSendUsing As String
    Dim sAuthentication As String
    Dim Flds
    Dim szServer As String
        
    'https://www.experts-exchange.com/questions/23044027/CDO-Message-sendusing-and-smtpauthenticate.html
    Const cdoSendUsingPickup = 1 'Send message using the local SMTP service pickup directory.
    Const cdoSendUsingPort = 2 'Send the message using the network (SMTP over the network).
    
    Const cdoAnonymous = 0 'Do not authenticate
    Const cdoBasic = 1 'basic (clear-text) authentication
    Const cdoNTLM = 2 'NTLM
 
    Dim oMyMail As Object
    Dim objShell As Object
    
    On Error GoTo ErrHandler
    
    sSendUsing = cdoSendUsingPort
    Select Case [SETTINGS_SMTP_SENDUSING].Value
        Case "Network"
            sSendUsing = cdoSendUsingPort
        Case "Local Catalog"
            sSendUsing = cdoSendUsingPickup
    End Select

    sAuthentication = cdoAnonymous
    Select Case [SETTINGS_SMTP_AUTHENTICATION].Value
        Case "Basic"
            sAuthentication = cdoBasic
        Case "Anonymous"
            sAuthentication = cdoAnonymous
        Case "NTLM"
            sAuthentication = cdoNTLM
    End Select
    
    Set objShell = CreateObject("WScript.Shell")
    Set oMyMail = CreateObject("CDO.Message")
    Set iConf = CreateObject("CDO.Configuration")
    Set Flds = iConf.Fields
    szServer = "http://schemas.microsoft.com/cdo/configuration/"
    
    Select Case "Network"  '[SETTINGS_SMTP_SENDUSING].Value
        Case "Network"
            sSendUsing = cdoSendUsingPort
        Case "Local Catalog"
            sSendUsing = cdoSendUsingPickup
    End Select
            
        With Flds
        .Item(szServer & "sendusing") = sSendUsing
        .Item(szServer & "smtpserver") = [SETTINGS_SMTP_SERVER].Value
        .Item(szServer & "smtpserverport") = [SETTINGS_SMTP_PORT].Value
        
        .Item(szServer & "smtpconnectiontimeout") = 60
        If [SETTINGS_SMTP_TIMEOUT].Value <> vbNullString Then
            If IsNumeric([SETTINGS_SMTP_TIMEOUT].Value) Then
                .Item(szServer & "smtpconnectiontimeout") = [SETTINGS_SMTP_TIMEOUT].Value
            End If
        End If
        
        .Item(szServer & "smtpauthenticate") = sAuthentication
        .Item(szServer & "smtpusessl") = IIf([SETTINGS_SMTP_USESSL].Value = "Y", True, False)
        .Item(szServer & "sendusername") = [SETTINGS_SMTP_USERNAME].Value
        .Item(szServer & "sendpassword") = [SETTINGS_SMTP_PASSWORD].Value
        .Update
    End With
    
    With oMyMail
            Set .Configuration = iConf
            .bodypart.Charset = "utf-8"
            .to = sRecipients
            
            If sCC <> vbNullString Then
            .cc = sCC
            End If
            
            If sBCC <> vbNullString Then
                .bcc = sBCC
            End If
            
            .From = IIf([SETTINGS_SMTP_FROM].Value <> vbNullString, _
                        [SETTINGS_SMTP_FROM].Value, _
                        IIf(sFrom <> vbNullString, sFrom, "Reports Controller"))
                        
            .Subject = sSubject
            .TextBody = sMessage & vbCrLf & _
                vbCrLf & _
                ThisWorkbook.Name & vbCrLf & _
                objShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
            
            If sAttachmentPath <> vbNullString Then
                .AddAttachment sAttachmentPath
            End If
            
            .send
    End With
    
    Send_EMail_CDO = True
    
Exit_sub:
    Set oMyMail = Nothing
    Set iConf = Nothing
    Set Flds = Nothing
    Exit Function

ErrHandler:
    ' write_log...
    Debug.Print Now, "Send_EMail_CDO", Err.Number, Err.Description
    Err.Clear
    
    GoTo Exit_sub
    Resume
End Function

Function Send_Email_Outlook(sRecipients As String, _
                    sSubject As String, _
                    Optional sMessage As String, _
                    Optional sAttachmentPath As String, _
                    Optional Importance As String = "Normal", _
                    Optional sFrom As String, _
                    Optional sCC As String, _
                    Optional sBCC As String)

    Dim oOutlook As Object
    Dim oMyMail As Object
    Dim objShell As Object
    
    Const olMailItem = 0
    
    ' get or create outlook
    On Error Resume Next
    Set oOutlook = GetObject(, "Outlook.Application")
    Err.Clear
    
    On Error GoTo ErrHandler
    If oOutlook Is Nothing Then
        Set oOutlook = CreateObject("Outlook.Application")
    End If
    
    Set objShell = CreateObject("WScript.Shell")
    Set oMyMail = oOutlook.CreateItem(olMailItem)
    
    With oMyMail
        If sFrom <> vbNullString Then
            .SentOnBehalfOfName = sFrom
        End If
    
        .to = sRecipients
        If sCC <> vbNullString Then
            .cc = sCC
        End If
        
        If sBCC <> vbNullString Then
            .bcc = sBCC
        End If
        
        .Subject = sSubject
    
        .BodyFormat = 2 ' 2 olFormatHTML
         
        .htmlbody = Replace(Replace(sMessage, vbCrLf, "<br>"), Chr(10), "<br>") _
            & "<br>" & _
            "<br>" & _
            ThisWorkbook.Name & "<br>" & _
            objShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
        
        If sAttachmentPath <> vbNullString Then
            .Attachments.Add sAttachmentPath
        End If
        
        ' Normal importance is default
        Select Case Importance
            Case "High"
                .Importance = enumMailImportance.High
            Case "Low"
                .Importance = enumMailImportance.Low
        End Select
                
        '.display
        .send
    End With
    
    Send_Email_Outlook = True

Exit_sub:
    Set oMyMail = Nothing
    Set oOutlook = Nothing
    Set objShell = Nothing
    Exit Function

ErrHandler:
    ' write_log...
    Debug.Print Now, "Send_Email_Outlook", Err.Number, Err.Description
    Err.Clear
        
    GoTo Exit_sub
    Resume
End Function

'
'Sub Send_Email_Outlook(Subject As String, _
'                        Recipients As String, _
'                        Optional sMessage As String, _
'                        Optional Importance As String = "High", _
'                        Optional AttachmentPath As String)
'
'    Dim oOutlook As Object
'    Dim oMail As Object
'
'    On Error Resume Next
'    Set oOutlook = GetObject(, "Outlook.Application")
'    On Error GoTo 0
'
'    On Error GoTo ErrHandler
'    If oOutlook Is Nothing Then
'        Set oOutlook = CreateObject("Outlook.Application")
'    End If
'
'    Set oMail = oOutlook.CreateItem(0) ' oMailItem
'    With oMail
'        .Subject = Subject
'        .To = Recipients
'
'        ' https://msdn.microsoft.com/en-us/library/office/ff866430.aspx
'        Select Case Importance
'            Case "Low"
'                .Importance = 0 ' olImportanceLow
'            Case "High"
'                .Importance = 2 ' olImportanceHigh
'        End Select
'
'        If AttachmentPath <> vbNullString Then
'            On Error Resume Next
'            .Attachments.Add AttachmentPath
'            On Error GoTo 0
'            Err.Clear
'        End If
'
'        .TextBody = sMessage
'
'        On Error GoTo ErrHandler
'        .Save
'        .display
'        .send
'    End With
'
'ErrHandler:
'    Set oMail = Nothing
'    Set oOutlook = Nothing
'    Err.Clear
'End Sub
'
'Sub Send_EMail_CDO(sSubject As String, _
'        sRecipients As String, _
'        Optional sMessage As String, _
'        Optional Importance As String = "High", _
'        Optional AttachmentPath As String)
'
'    Dim iMsg As Object
'    Dim iConf As Object
'    Dim strbody As String
'    Dim sSendUsing As String
'    Dim sAuthentication As String
'    Dim Flds
'    Dim szServer As String
'
'    'https://www.experts-exchange.com/questions/23044027/CDO-Message-sendusing-and-smtpauthenticate.html
'    Const cdoSendUsingPickup = 1 'Send message using the local SMTP service pickup directory.
'    Const cdoSendUsingPort = 2 'Send the message using the network (SMTP over the network).
'
'    Const cdoAnonymous = 0 'Do not authenticate
'    Const cdoBasic = 1 'basic (clear-text) authentication
'    Const cdoNTLM = 2 'NTLM
'
'    Dim oMyMail
'
'    On Error GoTo ErrHandler
'
'    Set oMyMail = CreateObject("CDO.Message")
'    Set iConf = CreateObject("CDO.Configuration")
'    Set Flds = iConf.Fields
'    szServer = "http://schemas.microsoft.com/cdo/configuration/"
'
'    Select Case [SETTINGS_SMTP_SENDUSING].Value
'        Case "Network"
'            sSendUsing = cdoSendUsingPort
'        Case "Local Catalog"
'            sSendUsing = cdoSendUsingPickup
'    End Select
'
'    Select Case [SETTINGS_SMTP_AUTHENTICATION].Value
'        Case "Basic"
'            sAuthentication = cdoBasic
'        Case "Anonymous"
'            sAuthentication = cdoAnonymous
'        Case "NTLM"
'            sAuthentication = cdoNTLM
'    End Select
'
'    With Flds
'        .Item(szServer & "sendusing") = sSendUsing
'        .Item(szServer & "smtpserver") = [SETTINGS_SMTP_SERVER].Value
'        .Item(szServer & "smtpserverport") = [SETTINGS_SMTP_PORT].Value
'        .Item(szServer & "smtpconnectiontimeout") = [SETTINGS_SMTP_TIMEOUT].Value ' quick timeout
'        .Item(szServer & "smtpauthenticate") = sAuthentication
'        .Item(szServer & "sendusername") = [SETTINGS_SMTP_USERNAME].Value
'        .Item(szServer & "sendpassword") = [SETTINGS_SMTP_PASSWORD].Value
'        .Update
'    End With
'
'    With oMyMail
'            Set .Configuration = iConf
'            .bodypart.Charset = "utf-8"
'            .To = sRecipients
'            .from = [SETTINGS_SMTP_FROM].Value
'            .Subject = sSubject
'
'            .TextBody = sMessage
'            .send
'    End With
'
'ErrHandler:
'    Set oMyMail = Nothing
'    Set iConf = Nothing
'    Set Flds = Nothing
'    Err.Clear
'End Sub
'
