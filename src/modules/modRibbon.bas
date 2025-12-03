Attribute VB_Name = "modRibbon"
' ====================================================================================
' Module    : modRibbon
' Purpose   : Ribbon callback procedures for the RDD Add-In custom Ribbon tab.
'             Contains all callback methods referenced in the Ribbon XML customUI.
'             Acts as the routing layer between Ribbon UI and business logic.
'
' Public API:
'   === Ribbon Lifecycle ===
'   - RB75dd2c44_Ribbon_OnLoad          : Ribbon initialization callback
'
'   === Button OnAction Callbacks ===
'   - RB75dd2c44_btnAddRoom_OnAction          : Add new room sheet
'   - RB75dd2c44_btnEditRoomIdentity_OnAction : Edit room identity
'   - RB75dd2c44_BtnRemoveRoom_OnAction       : Remove current room sheet
'
'   === Button getEnabled Callbacks ===
'   - RB75dd2c44_btnAddRoom_getEnabled          : Controls Add Room button state
'   - RB75dd2c44_btnEditRoomIdentity_getEnabled : Controls Edit Room button state
'
'   === Dynamic Content Callbacks ===
'   - (getLabel, getImage, getVisible callbacks if present)
'
' Dependencies:
'   - clsState        : Stores Ribbon UI reference, state flags
'   - modMain         : Business logic for all Ribbon actions
'   - modRooms        : Room-specific operations
'   - modCellCtxMnu   : Context menu initialization
'
' Notes:
'   - Keep callbacks THIN - delegate all business logic to modMain or feature modules
'   - Return values via ByRef parameters (not Function returns)
'   - Use clsState.InvalidateControl(id) to refresh specific controls
'   - Use clsState.InvalidateRibbon to refresh entire Ribbon
'   - getEnabled callbacks determine button availability (True/False)
'   - Check Workbooks.Count > 0 before accessing ActiveWorkbook/ActiveSheet
'   - Callbacks should not raise errors to Ribbon (causes Excel instability)
'   - Use On Error Resume Next or proper error handling
'   - Log errors via modErr.ReportError if needed
'   - Callback names must match EXACTLY with Ribbon XML definitions
'   - Ribbon XML file location: customUI/customUI14.xml (in XLAM)
'
' ====================================================================================

Option Explicit
Option Private Module

' ==== Ribbon onLoad ====
Sub RB75dd2c44_Ribbon_OnLoad(ribbon As IRibbonUI)
    Set clsState.RibbonUI = ribbon
    Call InitCellCtxMenu
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

Sub RB75dd2c44_btnEditRoomIdentity_OnAction(control As IRibbonControl)
    Call modMain.EditRoomIdentity
End Sub


Sub RB75dd2c44_btnEditRoomIdentity_getEnabled(control As IRibbonControl, ByRef returnedVal)
    If Workbooks.Count > 0 Then
        If modMain.IsRDDWorkbook(ActiveWorkbook) And modRooms.IsRoomSheet(ActiveSheet) Then
            returnedVal = True
        End If
    End If
End Sub

' -----------------------------------------------------------------------------

Sub RB75dd2c44_BtnRemoveRoom_OnAction(control As IRibbonControl)
    Call modMain.RemoveCurrentRoom
End Sub

Sub RB75dd2c44_btnRemoveRoom_getEnabled(control As IRibbonControl, ByRef returnedVal)
    If Workbooks.Count > 0 Then
        If modMain.IsRDDWorkbook(ActiveWorkbook) And modRooms.IsRoomSheet(ActiveSheet) Then
            returnedVal = True
        End If
    End If
End Sub

' -----------------------------------------------------------------------------

Sub RB75dd2c44_btnSyncLists_OnAction(control As IRibbonControl)
    MsgBox "Sync Lists is not implemented yet.", vbInformation, "Sync Lists"
    modMain.SyncAllLists
End Sub

Sub RB75dd2c44_btnSyncLists_getEnabled(control As IRibbonControl, ByRef returnedVal)
    If Workbooks.Count > 0 Then
        If modMain.IsRDDWorkbook(ActiveWorkbook) Then
            returnedVal = True
        End If
    End If
End Sub

' -----------------------------------------------------------------------------

Sub RB75dd2c44_btnValidate_OnAction(control As IRibbonControl)
    modMain.ValidateRoomData
End Sub

Sub RB75dd2c44_btnValidate_getEnabled(control As IRibbonControl, ByRef returnedVal)
    If Workbooks.Count > 0 Then
        If modMain.IsRDDWorkbook(ActiveWorkbook) Then
            ' Check if there's at least one room sheet
            Dim sheet As Worksheet
            For Each sheet In ActiveWorkbook.Worksheets
                If modRooms.IsRoomSheet(sheet) Then
                    returnedVal = True
                    Exit Sub
                End If
            Next sheet
        End If
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
    If modMain.IsRDDWorkbook(ActiveWorkbook) Then
        ' TODO: I still need to set conditions for displaying.
    End If
End Sub

' -----------------------------------------------------------------------------

Sub RB75dd2c44_btnBuildChart_OnAction(control As IRibbonControl)
    MsgBox "Build Chart is not implemented yet.", vbInformation, "Build Chart"
    'modPDC.GeneratePuzzleChart 'TODO: Sub still needs to be adjusted, supplemented, and tested.
End Sub

Sub RB75dd2c44_btnBuildChart_getEnabled(control As IRibbonControl, ByRef returnedVal)
    If modMain.IsRDDWorkbook(ActiveWorkbook) Then
        ' TODO: I still need to set conditions for displaying.
    End If
End Sub

' -----------------------------------------------------------------------------

Sub RB75dd2c44_btnUpdateChart_OnAction(control As IRibbonControl)
    MsgBox "Update Chart is not implemented yet.", vbInformation, "Update Chart"
    'modPDC.SyncPuzzleChart 'TODO: Sub still needs to be adjusted, supplemented, and tested.
End Sub

Sub RB75dd2c44_btnUpdateChart_getEnabled(control As IRibbonControl, ByRef returnedVal)
    If modMain.IsRDDWorkbook(ActiveWorkbook) Then
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
If modMain.IsRDDWorkbook(ActiveWorkbook) Then
End If
End Sub

Sub RB75dd2c44_btnExportCsv_OnAction(control As IRibbonControl)
    ' Export puzzles, edges, rooms to CSV.
    ' TODO: Implement CSV export.
    MsgBox "CSV export is not implemented yet.", vbInformation, "Export CSV"
End Sub

Sub RB75dd2c44_btnExportCsv_getEnabled(control As IRibbonControl, ByRef returnedVal)
If modMain.IsRDDWorkbook(ActiveWorkbook) Then
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
    
    If clsState.CellCtxMenuType = CCM_Rooms Then
        returnedVal = "Add New Room"
    End If
    
End Sub

Sub RB75dd2c44_btnDynCtxMnu1_getVisible(control As IRibbonControl, ByRef returnedVal)
    If clsState.CellCtxMenuType <> 0 Then
    Call EnsureCellCtxMenuReady
    returnedVal = True
    End If
End Sub

Sub RB75dd2c44_btnDynCtxMnu1_onAction(control As IRibbonControl)
    Select Case clsState.CellCtxMenuType
        Case CCM_Rooms
            modMain.AddNewRoomFromCellCtxMnu
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
        Call EnsureCellCtxMenuReady
        returnedVal = True
    End If
End Sub

Sub RB75dd2c44_btnDynCtxMnu2_onAction(control As IRibbonControl)
    If clsState.CellCtxMenuType = CCM_Rooms Then
        modMain.GotoRoomFromCell
    End If
End Sub

