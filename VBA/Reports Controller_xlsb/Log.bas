Attribute VB_Name = "Log"
Option Explicit

Sub Write_Log(report_row_id As Long, proc_id As String)
    Dim log_row As Long
    
    Application.AutoCorrect.AutoExpandListRange = True
    
    If LOG_Table.DataBodyRange Is Nothing Then
        log_row = LOG_Table.HeaderRowRange.Row + 1
    Else
        log_row = LOG_Table.HeaderRowRange.Row + LOG_Table.DataBodyRange.Rows.count + 1
    End If
    
    LOG_Table.Parent.Cells(log_row, _
        LOG_Table.ListColumns("Report ID").Range.Column).Value = _
            Control_Table.Parent.Cells(report_row_id, Control_Table.ListColumns("Report ID *").Range.Column).Value
    LOG_Table.Parent.Cells(log_row, _
        LOG_Table.ListColumns("Process ID").Range.Column).Value = proc_id
    LOG_Table.Parent.Cells(log_row, _
        LOG_Table.ListColumns("Start Time").Range.Column).Value = Now()
End Sub

Function Get_Last_Log_Record(report_row_id As Long, ColumnName As String) As Double
    Dim SearchResult As Range
    If LOG_Table.DataBodyRange Is Nothing Then
        Get_Last_Log_Record = -1
    Else
        Set SearchResult = LOG_Table.ListColumns("Report ID").DataBodyRange.Find( _
            Control_Table.Parent.Cells(report_row_id, _
                Control_Table.ListColumns("Report ID *").Range.Column).Value, _
            searchdirection:=xlPrevious)
        
        If Not SearchResult Is Nothing Then
            Get_Last_Log_Record = LOG_Table.Parent.Cells(SearchResult.Row, _
                    LOG_Table.ListColumns(ColumnName).Range.Column).Value
        Else
            Get_Last_Log_Record = -1
        End If
    End If
    Set SearchResult = Nothing
End Function

' get .log records
' http://stackoverflow.com/questions/13598691/read-number-of-lines-in-large-text-file-vb6
' or simple version - filesystemobject - fine for our small logs
' https://blogs.technet.microsoft.com/heyscriptingguy/2006/03/03/how-can-i-read-just-the-last-line-of-a-text-file/
