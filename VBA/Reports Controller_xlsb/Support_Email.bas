Attribute VB_Name = "Support_Email"
Option Explicit

Sub Send_Email_Outlook(Subject As String, Recipients As String, Optional Importance As String = "Normal", Optional AttachmentPath As String)
    Dim oOutlook As Object
    Dim oMail As Object

    On Error Resume Next
    Set oOutlook = GetObject(, "Outlook.Application")
    On Error GoTo 0
    
    On Error GoTo ErrHandler
    If oOutlook Is Nothing Then
        Set oOutlook = CreateObject("Outlook.Application")
    End If
    
    Set oMail = oOutlook.CreateItem(0) ' oMailItem
    With oMail
        .Subject = Subject
        .to = Recipients
                
        ' https://msdn.microsoft.com/en-us/library/office/ff866430.aspx
        Select Case Importance
            Case "Low"
                .Importance = 0 ' olImportanceLow
            Case "High"
                .Importance = 2 ' olImportanceHigh
        End Select
                
        If AttachmentPath <> vbNullString Then
            On Error Resume Next
            .attachments.Add AttachmentPath
            On Error GoTo 0
            Err.Clear
        End If
        
        On Error GoTo ErrHandler
        .Save
        .display
        .send
    End With
    
ErrHandler:
    Set oMail = Nothing
    Set oOutlook = Nothing
End Sub
