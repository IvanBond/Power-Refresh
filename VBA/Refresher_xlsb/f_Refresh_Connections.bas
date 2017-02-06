Attribute VB_Name = "f_Refresh_Connections"
Option Explicit
' Update executed from Excel to be able to handle errors
'
Function UpdateConnections() As Boolean
    Dim cnct As Variant
    Dim slc As SlicerCache
    Dim BeforeAction
    
    On Error Resume Next
    If IsError(target_wb.Model.ModelTables.count) Then
        ' cannot access model
        ' do nothing
    Else
        If target_wb.Model.ModelTables.count > 0 Then
            target_wb.Model.Initialize
            Application.Wait Now + TimeValue("0:00:05")
        End If
    End If
    
    On Error GoTo ErrHandler
    
    ' deny background refresh
    ' ToThink - probably worth to restore initial settings
    ' however, if workbook is done for Power Refresh solution, it should not contain "background" connections
    ' create 2D array, restore settings after update
    For Each cnct In target_wb.Connections
        Select Case cnct.Type
            Case xlConnectionTypeODBC
                cnct.ODBCConnection.BackgroundQuery = False
            Case xlConnectionTypeOLEDB
                cnct.OLEDBConnection.BackgroundQuery = False
        End Select
    Next cnct
    
    ' no need to turn on RefreshOnRefreshAll
    ' can be scenario with own macro, that uses some connections
    BeforeAction = Now()
    target_wb.RefreshAll
    Application.CalculateUntilAsyncQueriesDone
    DoEvents
        
    For Each cnct In target_wb.Connections
        Select Case cnct.Type
            Case xlConnectionTypeODBC
                Do While cnct.ODBCConnection.Refreshing
                  DoEvents
                Loop
            Case xlConnectionTypeOLEDB
                Do While cnct.OLEDBConnection.Refreshing
                  DoEvents
                Loop
        End Select
    Next cnct
        
    Application.Calculate
    Application.CalculateUntilAsyncQueriesDone
    
    ' if faster than 2 second
    If Round((Now() - BeforeAction) * 86400, 0) < 2 Then
        ' too fast, most probably error
        ' TOThink: think about new parameter: Trigger too short refresh time as error
        ' or manual parameter in seconds, e.g. 5 sec, if refresh < 5 sec - count as error
        ' if Too_Short_Time_Is_Error then
        Call Write_Log("Error. Too short time of RefreshAll", bMandatoryLogRecord)
        GoTo ErrHandler
    End If
    
    ' wait for refresh of cube formulas
    Application.Wait (Now + TimeValue("0:00:10"))
    
    If Err.Number <> 0 Then GoTo ErrHandler
    
    ' update cache after Model refresh
    ' TOThink: what method is better ?
    For Each slc In target_wb.SlicerCaches
        slc.ClearManualFilter
        slc.ClearAllFilters
    Next slc
    ' if needed, slicer default value can be set in BeforeSave event of target workbook, or in custom macro
    
    ' for each
    ' pivotCache.Refresh ?
    ' I've seen issues with slicers after this method. Maybe fixed...
    ' or maybe only if Pivot on DataModel...
    
    ' wait for refresh of cube formulas
    If target_wb.SlicerCaches.count > 0 Then
        Application.Calculate
        Application.CalculateUntilAsyncQueriesDone
        Application.Wait (Now + TimeValue("0:00:10"))
    End If
    
    If Not Application.CalculationState = xlDone Then
        DoEvents
        Application.Wait (Now + TimeValue("0:00:02"))
    End If
    
    UpdateConnections = True

Exit_Function:
    
    Exit Function
    
ErrHandler:
    bGlobalError = True
    Call Write_Log(Err.Number & ": " & Err.Description)
    GoTo Exit_Function
    Resume
End Function
