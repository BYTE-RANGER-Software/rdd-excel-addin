Attribute VB_Name = "modRibbon"
Option Explicit
Option Private Module

'Callback for customUI.onLoad
Sub RB75dd2c44_Ribbon_OnLoad(ribbon As IRibbonUI)
    Set clsState.RibbonUI = ribbon
    Call InitCellCtxMnu
End Sub

'Callback for RB75dd2c44_btnAddRoom onAction
Sub RB75dd2c44_btnAddRoom_OnAction(control As IRibbonControl)
    Call modMain.AddNewRoom
End Sub

'Callback for RB75dd2c44_btnAddRoom getEnabled
Sub RB75dd2c44_btnAddRoom_getEnabled(control As IRibbonControl, ByRef returnedVal)
    returnedVal = (Workbooks.Count > 0)
End Sub

'Callback for RB75dd2c44_btnRemoveRoom onAction
Sub RB75dd2c44_BtnRemoveRoom_OnAction(control As IRibbonControl)
    Call modRooms.RemoveRoom
End Sub

'Callback for RB75dd2c44_btnRemoveRoom getEnabled
Sub RB75dd2c44_btnRemoveRoom_getEnabled(control As IRibbonControl, ByRef returnedVal)
If modMain.IsAddinWorkbook(ActiveWorkbook) And modRooms.IsRoomSheet(ActiveSheet) Then
    returnedVal = (Workbooks.Count > 0)
End If
End Sub

'Callback for RB75dd2c44_BtnUpdateLists onAction
Sub RB75dd2c44_btnUpdateLists_OnAction(control As IRibbonControl)
    Call modRooms.UpdateLists '
End Sub

'Callback for RB75dd2c44_BtnUpdateLists getEnabled
Sub RB75dd2c44_btnUpdateLists_getEnabled(control As IRibbonControl, ByRef returnedVal)
If modMain.IsAddinWorkbook(ActiveWorkbook) Then
    returnedVal = (Workbooks.Count > 0)
    End If
End Sub

'Callback for RB75dd2c44_BtnSyncLists onAction
Sub RB75dd2c44_btnSyncLists_OnAction(control As IRibbonControl)
End Sub

'Callback for RB75dd2c44_BtnSyncLists getEnabled
Sub RB75dd2c44_btnSyncLists_getEnabled(control As IRibbonControl, ByRef returnedVal)
If modMain.IsAddinWorkbook(ActiveWorkbook) Then
    returnedVal = (Workbooks.Count > 0)
    End If
End Sub

'Callback for RB75dd2c44_btnValidate onAction
Sub RB75dd2c44_btnValidate_OnAction(control As IRibbonControl)
End Sub

'Callback for RB75dd2c44_btnValidate getEnabled
Sub RB75dd2c44_btnValidate_getEnabled(control As IRibbonControl, ByRef returnedVal)
If modMain.IsAddinWorkbook(ActiveWorkbook) Then
End If
End Sub

'Callback for RB75dd2c44_btnRefreshList onAction
Sub RB75dd2c44_btnRefreshList_OnAction(control As IRibbonControl)
End Sub

'Callback for RB75dd2c44_btnRefreshList getEnabled
Sub RB75dd2c44_btnRefreshList_getEnabled(control As IRibbonControl, ByRef returnedVal)
If modMain.IsAddinWorkbook(ActiveWorkbook) Then
End If
End Sub

'Callback for RB75dd2c44_btnBuildData onAction
Sub RB75dd2c44_btnBuildData_OnAction(control As IRibbonControl)
End Sub

'Callback for RB75dd2c44_btnBuildData getEnabled
Sub RB75dd2c44_btnBuildData_getEnabled(control As IRibbonControl, ByRef returnedVal)
If modMain.IsAddinWorkbook(ActiveWorkbook) Then
End If
End Sub

'Callback for RB75dd2c44_btnBuildChart onAction
Sub RB75dd2c44_btnBuildChart_OnAction(control As IRibbonControl)
End Sub

'Callback for RB75dd2c44_btnBuildChart getEnabled
Sub RB75dd2c44_btnBuildChart_getEnabled(control As IRibbonControl, ByRef returnedVal)
If modMain.IsAddinWorkbook(ActiveWorkbook) Then
End If
End Sub

'Callback for RB75dd2c44_btnUpdateChart onAction
Sub RB75dd2c44_btnUpdateChart_OnAction(control As IRibbonControl)
End Sub

'Callback for RB75dd2c44_btnUpdateChart getEnabled
Sub RB75dd2c44_btnUpdateChart_getEnabled(control As IRibbonControl, ByRef returnedVal)
End Sub

'Callback for RB75dd2c44_btnExportPdf onAction
Sub RB75dd2c44_btnExportPdf_OnAction(control As IRibbonControl)
End Sub

'Callback for RB75dd2c44_btnExportPdf getEnabled
Sub RB75dd2c44_btnExportPdf_getEnabled(control As IRibbonControl, ByRef returnedVal)
If modMain.IsAddinWorkbook(ActiveWorkbook) Then
End If
End Sub

'Callback for RB75dd2c44_btnExportCsv onAction
Sub RB75dd2c44_btnExportCsv_OnAction(control As IRibbonControl)
End Sub

'Callback for RB75dd2c44_btnExportCsv getEnabled
Sub RB75dd2c44_btnExportCsv_getEnabled(control As IRibbonControl, ByRef returnedVal)
If modMain.IsAddinWorkbook(ActiveWorkbook) Then
End If
End Sub

'Callback for RB75dd2c44_BtnShowOptions onAction
Sub RB75dd2c44_btnShowOptions_OnAction(control As IRibbonControl)
    Call modMain.ShowOptions
End Sub

'Callback for RB75dd2c44_BtnShowLog onAction
Sub RB75dd2c44_btnShowLog_OnAction(control As IRibbonControl)
    Call modMain.ShowLog
End Sub


'Callback for RB75dd2c44_BtnShowManual onAction
Sub RB75dd2c44_btnShowManual_OnAction(control As IRibbonControl)
    Call modMain.ShowManual
End Sub

'Callback for RB75dd2c44_lblAddInVersion getLabel
Sub RB75dd2c44_btnAddInVersion_GetLabel(control As IRibbonControl, ByRef returnedVal)
    Dim strVer As String: strVer = modMain.AppVersion
    returnedVal = strVer
End Sub

Sub RB75dd2c44_btnAddInVersion_OnAction(control As IRibbonControl)
    Call modMain.ShowAbout
End Sub

'  Context menus
'---------------------------------------------------------------------------------

'Callback for RB75dd2c44_btnDynCtxMnu1 getLabel
Sub RB75dd2c44_btnDynCtxMnu1_getLabel(control As IRibbonControl, ByRef returnedVal)
    
    If clsState.CellCtxMenuType = CCM_Objects Then
        returnedVal = "Add New Object"
    ElseIf clsState.CellCtxMenuType = CCM_Rooms Then
        returnedVal = "Add New Room"
    End If
    
End Sub

'Callback for RB75dd2c44_btnDynCtxMnu1 getVisible
Sub RB75dd2c44_btnDynCtxMnu1_getVisible(control As IRibbonControl, ByRef returnedVal)
    If clsState.CellCtxMenuType <> 0 Then
    Call EnsureCellCtxMnuReady
    returnedVal = True
    End If
End Sub

'Callback for RB75dd2c44_btnDynCtxMnu1 onAction
Sub RB75dd2c44_btnDynCtxMnu1_onAction(control As IRibbonControl)
End Sub

'Callback for RB75dd2c44_btnDynCtxMnu2 getLabel
Sub RB75dd2c44_btnDynCtxMnu2_getLabel(control As IRibbonControl, ByRef returnedVal)
    If clsState.CellCtxMenuType = CCM_Rooms Then
        returnedVal = "Goto Room..."
    End If
End Sub

'Callback for RB75dd2c44_btnDynCtxMnu2 getVisible
Sub RB75dd2c44_btnDynCtxMnu2_getVisible(control As IRibbonControl, ByRef returnedVal)
    If clsState.CellCtxMenuType = CCM_Rooms Then
        Call EnsureCellCtxMnuReady
        returnedVal = True
    End If
End Sub

'Callback for RB75dd2c44_btnDynCtxMnu2 onAction
Sub RB75dd2c44_btnDynCtxMnu2_onAction(control As IRibbonControl)
End Sub

