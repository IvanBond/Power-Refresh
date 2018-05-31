Attribute VB_Name = "f_Refresh_Connections"
Option Explicit
Option Compare Text

Function RefreshWorkbook(Optional Wb As Workbook) As Boolean
    Dim cnct As Variant
    Dim slc As SlicerCache
    Dim BeforeAction
    Dim target_wb As Workbook
    Dim bCubeFormulasFound As Boolean
    Dim bScreenUpdatingInitial As Boolean
    Dim bEnableEventsInitial As Boolean
    Dim CalcModeInitial As Double
    Dim CursorStateInitial As Double
    
    On Error GoTo ErrHandler
    Debug.Print Now, "Updating connections..."
    
    With Application
        bScreenUpdatingInitial = .ScreenUpdating
        bEnableEventsInitial = .EnableEvents
        CalcModeInitial = .Calculation
        CursorStateInitial = .Cursor
        
        ' switch everything off
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
        .Cursor = xlWait
    End With
    
    If Wb Is Nothing Then
        Set target_wb = ThisWorkbook
    Else
        Set target_wb = Wb
    End If
    
    On Error Resume Next
    If IsError(target_wb.Model.ModelTables.count) Then
        ' cannot access model
        ' do nothing
    Else
        If target_wb.Model.ModelTables.count > 0 Then
            Application.StatusBar = "Initializing Data Model..."
            target_wb.Model.Initialize
            WaitSeconds 5
        End If
    End If
    
    Err.Clear
    On Error GoTo ErrHandler
    
    ' deny background refresh
    ' ToThink - probably worth to restore initial settings
    ' however, if workbook is done for Power Refresh solution, it should not contain "background" connections
    ' create 2D array, restore settings after update
    Application.StatusBar = "Switching off background refresh..."
    For Each cnct In target_wb.Connections
        Select Case cnct.Type
            Case xlConnectionTypeODBC
                cnct.ODBCConnection.BackgroundQuery = False
            Case xlConnectionTypeOLEDB
                cnct.OLEDBConnection.BackgroundQuery = False
        End Select
    Next cnct
    
    Application.StatusBar = "Refreshing Data Model and Connections..."
    target_wb.RefreshAll
    WaitSeconds 1
    Application.CalculateUntilAsyncQueriesDone
    WaitSeconds 1
    
    For Each cnct In target_wb.Connections
        Select Case cnct.Type
            Case xlConnectionTypeODBC
                Do While cnct.ODBCConnection.Refreshing
                    WaitSeconds 1
                Loop
            Case xlConnectionTypeOLEDB
                Do While cnct.OLEDBConnection.Refreshing
                    WaitSeconds 1
                Loop
        End Select
    Next cnct
    
    Application.StatusBar = "Calculating after connections refresh..."
    Application.Calculate
    Application.CalculateUntilAsyncQueriesDone
    WaitSeconds 1
    
    Application.StatusBar = "Checking existence of cube formulas..."
    bCubeFormulasFound = IsWBHasCubeFormulas(target_wb)
    
    ' update cache after Model refresh
    ' ignore all possible errors with slicers
    On Error Resume Next
    Application.StatusBar = "Updating slicers..."
    For Each slc In target_wb.SlicerCaches
        slc.ClearManualFilter
        slc.ClearAllFilters
        'slc.ClearDateFilter
    Next slc
    Err.Clear
    On Error GoTo ErrHandler
    ' if needed, slicer default value can be set in BeforeSave event of target workbook, or in custom macro
    
    If bCubeFormulasFound Then
        ' wait for refresh of cube formulas
        If target_wb.SlicerCaches.count > 0 Then
            Application.StatusBar = "Calculating after slicers refresh..."
            Application.Calculate
            Application.CalculateUntilAsyncQueriesDone
        End If
        
        Application.StatusBar = "Waiting for cube formulas..."
        WaitSeconds 20
    End If
    
    If Not Application.CalculationState = xlDone Then
        ' infinite loop can be trully infinite
        ' so just delay
        Application.StatusBar = "Waiting for application to calculate..."
        WaitSeconds 5
    End If
        
    RefreshWorkbook = True

Exit_Function:
    On Error Resume Next
    
    ' restore initial state
    With Application
        .ScreenUpdating = bScreenUpdatingInitial
        .EnableEvents = bEnableEventsInitial
        .Calculation = CalcModeInitial
        .Cursor = CursorStateInitial
        .StatusBar = vbNullString
    End With
    
    Exit Function
    
ErrHandler:
    Debug.Print Now, "RefreshWorkbook", Err.Number, Err.Description, Application.StatusBar
    
    Err.Clear
    GoTo Exit_Function
    Resume ' for debug purpose
End Function

Private Function IsWBHasCubeFormulas(Optional Wb As Workbook) As Boolean
    Dim sh As Worksheet
    Dim cell As Range
    Dim bFound As Boolean
    Dim bScreenUpdatingInitial As Boolean
    Dim bEnableEventsInitial As Boolean
    Dim CalcModeInitial As Integer
    Dim rngFormulas As Range
    
    On Error GoTo ErrHandler
    
    If Wb Is Nothing Then
        Set Wb = ThisWorkbook ' ActiveWorkbook ' alternatively
    End If
    
    With Application
        bScreenUpdatingInitial = .ScreenUpdating
        bEnableEventsInitial = .EnableEvents
        CalcModeInitial = .Calculation
        
        ' switch everything off
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With
    
    For Each sh In Wb.Sheets
        'Debug.Print sh.Name
        
        Err.Clear
        On Error Resume Next
        Set rngFormulas = sh.Cells.SpecialCells(xlCellTypeFormulas)
        bFound = (Err.Number = 0) ' no error, means SpecialCells returned non-empty range
        Err.Clear
        On Error GoTo ErrHandler
        
        ' if result of SpecialCells was non-empty - check formulas
        If bFound Then
            For Each cell In rngFormulas
                'Debug.Print cell.Formula
                If Left(cell.Formula, 5) = "=CUBE" Then
                    IsWBHasCubeFormulas = True
                    GoTo Exit_Function
                End If
            Next cell
        End If
    Next sh
    
Exit_Function:
    On Error Resume Next
    
    ' restore initial state
    With Application
        .ScreenUpdating = bScreenUpdatingInitial
        .EnableEvents = bEnableEventsInitial
        .Calculation = CalcModeInitial
    End With
    
    Exit Function
    
ErrHandler:
    If Err.Number <> 0 Then
        Debug.Print Now, "IsWBHasCubeFormulas", Err.Number & ": " & Err.Description
        Err.Clear
    End If
    
    GoTo Exit_Function
    Resume ' for debug purpose
End Function

' ================ OBSOLETE

' Update executed from Excel to be able to handle errors
'
'Function UpdateConnections() As Boolean
'    Dim cnct As Variant
'    Dim slc As SlicerCache
'    Dim BeforeAction
'
'    On Error Resume Next
'    If IsError(target_wb.Model.ModelTables.count) Then
'        ' cannot access model
'        ' do nothing
'    Else
'        If target_wb.Model.ModelTables.count > 0 Then
'            target_wb.Model.Initialize
'            Application.Wait Now + TimeValue("0:00:05")
'        End If
'    End If
'
'    On Error GoTo ErrHandler
'
'    ' deny background refresh
'    ' ToThink - probably worth to restore initial settings
'    ' however, if workbook is done for Power Refresh solution, it should not contain "background" connections
'    ' create 2D array, restore settings after update
'    For Each cnct In target_wb.Connections
'        Select Case cnct.Type
'            Case xlConnectionTypeODBC
'                cnct.ODBCConnection.BackgroundQuery = False
'            Case xlConnectionTypeOLEDB
'                cnct.OLEDBConnection.BackgroundQuery = False
'        End Select
'    Next cnct
'
'    ' no need to turn on RefreshOnRefreshAll
'    ' can be scenario with own macro, that uses some connections
'    BeforeAction = Now()
'    target_wb.RefreshAll
'    Application.CalculateUntilAsyncQueriesDone
'    DoEvents
'
'    For Each cnct In target_wb.Connections
'        Select Case cnct.Type
'            Case xlConnectionTypeODBC
'                Do While cnct.ODBCConnection.Refreshing
'                  DoEvents
'                Loop
'            Case xlConnectionTypeOLEDB
'                Do While cnct.OLEDBConnection.Refreshing
'                  DoEvents
'                Loop
'        End Select
'    Next cnct
'
'    Application.Calculate
'    Application.CalculateUntilAsyncQueriesDone
'
'    ' if faster than 2 second
'    If Round((Now() - BeforeAction) * 86400, 0) < 2 Then
'        ' too fast, most probably error
'        ' TOThink: think about new parameter: Trigger too short refresh time as error
'        ' or manual parameter in seconds, e.g. 5 sec, if refresh < 5 sec - count as error
'        ' if Too_Short_Time_Is_Error then
'        Call Write_Log("Error. Too short time of RefreshAll", bMandatoryLogRecord)
'        GoTo ErrHandler
'    End If
'
'    ' wait for refresh of cube formulas
'    Application.Wait (Now + TimeValue("0:00:10"))
'
'    If Err.Number <> 0 Then GoTo ErrHandler
'
'    ' update cache after Model refresh
'    ' TOThink: what method is better ?
'    For Each slc In target_wb.SlicerCaches
'        slc.ClearManualFilter
'        slc.ClearAllFilters
'    Next slc
'    ' if needed, slicer default value can be set in BeforeSave event of target workbook, or in custom macro
'
'    ' for each
'    ' pivotCache.Refresh ?
'    ' I've seen issues with slicers after this method. Maybe fixed...
'    ' or maybe only if Pivot on DataModel...
'
'    ' wait for refresh of cube formulas
'    If target_wb.SlicerCaches.count > 0 Then
'        Application.Calculate
'        Application.CalculateUntilAsyncQueriesDone
'        Application.Wait (Now + TimeValue("0:00:10"))
'    End If
'
'    If Not Application.CalculationState = xlDone Then
'        DoEvents
'        Application.Wait (Now + TimeValue("0:00:02"))
'    End If
'
'    UpdateConnections = True
'
'Exit_Function:
'
'    Exit Function
'
'ErrHandler:
'    bGlobalError = True
'    Call Write_Log(Err.Number & ": " & Err.Description)
'    GoTo Exit_Function
'    Resume
'End Function

