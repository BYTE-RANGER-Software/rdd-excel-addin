Attribute VB_Name = "modRibbon"
Option Explicit
Option Private Module

' ==== Ribbon onLoad ====
Sub RB75dd2c44_Ribbon_OnLoad(ribbon As IRibbonUI)
    Set clsState.RibbonUI = ribbon
    Call InitCellCtxMnu
End Sub

' ============================================================================
' Rooms group
' ============================================================================

Sub RB75dd2c44_btnAddRoom_OnAction(control As IRibbonControl)
    Call modMain.AddNewRoom
End Sub


Sub RB75dd2c44_btnAddRoom_getEnabled(control As IRibbonControl, ByRef returnedVal)
    returnedVal = (Workbooks.Count > 0)
End Sub

' -----------------------------------------------------------------------------

Sub RB75dd2c44_BtnRemoveRoom_OnAction(control As IRibbonControl)
    Call modMain.RemoveCurrentRoom
End Sub

Sub RB75dd2c44_btnRemoveRoom_getEnabled(control As IRibbonControl, ByRef returnedVal)
If modMain.IsAddinWorkbook(ActiveWorkbook) And modRooms.IsRoomSheet(ActiveSheet) Then
    returnedVal = (Workbooks.Count > 0)
End If
End Sub

' -----------------------------------------------------------------------------

Sub RB75dd2c44_btnSyncLists_OnAction(control As IRibbonControl)
    MsgBox "Sync Lists is not implemented yet.", vbInformation, "Sync Lists"
    'modRooms.SyncLists ' TODO: Sub still needs to be adjusted, supplemented, and tested.
End Sub

Sub RB75dd2c44_btnSyncLists_getEnabled(control As IRibbonControl, ByRef returnedVal)
If modMain.IsAddinWorkbook(ActiveWorkbook) Then
    returnedVal = (Workbooks.Count > 0)
    End If
End Sub

' -----------------------------------------------------------------------------

Sub RB75dd2c44_btnValidate_OnAction(control As IRibbonControl)
    MsgBox "Validate Data is not implemented yet.", vbInformation, "Validate"
    'modPDC.ValidateModel ' TODO: Sub still needs to be adjusted, supplemented, and tested.
End Sub

Sub RB75dd2c44_btnValidate_getEnabled(control As IRibbonControl, ByRef returnedVal)
If modMain.IsAddinWorkbook(ActiveWorkbook) Then
    ' TODO: I still need to set conditions for displaying.
End If
End Sub

' ============================================================================
' Dependency Chart group
' ============================================================================

Sub RB75dd2c44_btnBuildData_OnAction(control As IRibbonControl)
    MsgBox "Build PDC Data is not implemented yet.", vbInformation, "Build Data"
    'modPDC.BuildPdcData 'TODO: Sub still needs to be adjusted, supplemented, and tested.
End Sub

Sub RB75dd2c44_btnBuildData_getEnabled(control As IRibbonControl, ByRef returnedVal)
    If modMain.IsAddinWorkbook(ActiveWorkbook) Then
        ' TODO: I still need to set conditions for displaying.
    End If
End Sub

' -----------------------------------------------------------------------------

Sub RB75dd2c44_btnBuildChart_OnAction(control As IRibbonControl)
    MsgBox "Build Chart is not implemented yet.", vbInformation, "Build Chart"
    'modPDC.GeneratePuzzleChart 'TODO: Sub still needs to be adjusted, supplemented, and tested.
End Sub

Sub RB75dd2c44_btnBuildChart_getEnabled(control As IRibbonControl, ByRef returnedVal)
    If modMain.IsAddinWorkbook(ActiveWorkbook) Then
        ' TODO: I still need to set conditions for displaying.
    End If
End Sub

' -----------------------------------------------------------------------------

Sub RB75dd2c44_btnUpdateChart_OnAction(control As IRibbonControl)
    MsgBox "Update Chart is not implemented yet.", vbInformation, "Update Chart"
    'modPDC.SyncPuzzleChart 'TODO: Sub still needs to be adjusted, supplemented, and tested.
End Sub

Sub RB75dd2c44_btnUpdateChart_getEnabled(control As IRibbonControl, ByRef returnedVal)
    If modMain.IsAddinWorkbook(ActiveWorkbook) Then
        ' TODO: I still need to set conditions for displaying.
    End If
End Sub

' ============================================================================
' Export group
' ============================================================================

Sub RB75dd2c44_btnExportPdf_OnAction(control As IRibbonControl)
    ' Export current RDD views and chart to PDF.
    ' TODO: Implement PDF export.
    MsgBox "PDF export is not implemented yet.", vbInformation, "Export PDF"
End Sub


Sub RB75dd2c44_btnExportPdf_getEnabled(control As IRibbonControl, ByRef returnedVal)
If modMain.IsAddinWorkbook(ActiveWorkbook) Then
End If
End Sub

Sub RB75dd2c44_btnExportCsv_OnAction(control As IRibbonControl)
    ' Export puzzles, edges, rooms to CSV.
    ' TODO: Implement CSV export.
    MsgBox "CSV export is not implemented yet.", vbInformation, "Export CSV"
End Sub

Sub RB75dd2c44_btnExportCsv_getEnabled(control As IRibbonControl, ByRef returnedVal)
If modMain.IsAddinWorkbook(ActiveWorkbook) Then
End If
End Sub

' ============================================================================
' About group
' ============================================================================

Sub RB75dd2c44_btnShowOptions_OnAction(control As IRibbonControl)
    Call modMain.ShowOptions
End Sub

Sub RB75dd2c44_btnShowLog_OnAction(control As IRibbonControl)
    Call modMain.ShowLog
End Sub

Sub RB75dd2c44_btnShowManual_OnAction(control As IRibbonControl)
    Call modMain.ShowManual
End Sub

Sub RB75dd2c44_btnAddInVersion_GetLabel(control As IRibbonControl, ByRef returnedVal)
    Dim strVer As String: strVer = modMain.AppVersion
    returnedVal = strVer
End Sub

Sub RB75dd2c44_btnAddInVersion_OnAction(control As IRibbonControl)
    Call modMain.ShowAbout
End Sub

' ============================================================================
' Cell Context menus
' ============================================================================

Sub RB75dd2c44_btnDynCtxMnu1_getLabel(control As IRibbonControl, ByRef returnedVal)
    
    If clsState.CellCtxMenuType = CCM_Objects Then
        returnedVal = "Add New Object"
    ElseIf clsState.CellCtxMenuType = CCM_Rooms Then
        returnedVal = "Add New Room"
    End If
    
End Sub

Sub RB75dd2c44_btnDynCtxMnu1_getVisible(control As IRibbonControl, ByRef returnedVal)
    If clsState.CellCtxMenuType <> 0 Then
    Call EnsureCellCtxMnuReady
    returnedVal = True
    End If
End Sub

Sub RB75dd2c44_btnDynCtxMnu1_onAction(control As IRibbonControl)
    Select Case clsState.CellCtxMenuType
        Case CCM_Rooms
            modMain.AddNewRoom (False)
        Case Else
            ' TODO: I need to defined actions for other contexts
    End Select
End Sub

Sub RB75dd2c44_btnDynCtxMnu2_getLabel(control As IRibbonControl, ByRef returnedVal)
    If clsState.CellCtxMenuType = CCM_Rooms Then
        returnedVal = "Goto Room..."
    End If
End Sub

Sub RB75dd2c44_btnDynCtxMnu2_getVisible(control As IRibbonControl, ByRef returnedVal)
    If clsState.CellCtxMenuType = CCM_Rooms Then
        Call EnsureCellCtxMnuReady
        returnedVal = True
    End If
End Sub

Sub RB75dd2c44_btnDynCtxMnu2_onAction(control As IRibbonControl)
    If clsState.CellCtxMenuType = CCM_Rooms Then
        modMain.GotoRoomFromCell
    End If
End Sub

