VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Public WithEvents App As Application
Attribute App.VB_VarHelpID = -1

Private Sub App_SheetActivate(ByVal Sh As Object)
    Dim wb As Workbook: Set wb = Sh.Parent
    Dim wks As Worksheet: Set wks = Sh
    
    
    ' Only react if the sheet is NOT in the add-in.
    If Not wb Is RDDAddInWkBk And Not wb.IsAddin Then
        ' Check whether the target workbook has a sheet with the Tag SHEET_LISTS
        If modTags.SheetWithTagExists(wb, SHEET_DISPATCHER) Then
            'Don't update on sheet Lists
            If Sh.codeName <> SHEET_DISPATCHER And Not modTags.HasSheetTag(wks, SHEET_DISPATCHER) Then
                If clsState.RoomSheetChanged Then
                    clsState.RoomSheetChanged = False
                    Call UpdateLists
                End If
                If modRooms.IsRoomSheet(wks) Then
                    modRooms.ApplyParallaxRangeCover wks
                End If
            End If
        End If
    End If
    
    clsState.InvalidateRibbon
End Sub

Private Sub App_SheetChange(ByVal Sh As Object, ByVal target As Range)
    Dim wb As Workbook: Set wb = Sh.Parent
    Dim wks As Worksheet: Set wks = Sh
    
    
    ' Only react if the sheet is NOT in the add-in.
    If Not wb Is RDDAddInWkBk And Not wb.IsAddin Then
        ' Check whether the target workbook has a sheet with the Tag SHEET_LISTS
        If modTags.SheetWithTagExists(wb, SHEET_DISPATCHER) Then
            ' Only set the change flag if a room sheet has been changed.
            If modTags.HasSheetTag(wks, ROOM_SHEET_ID_TAG_NAME) Then
                clsState.RoomSheetChanged = True
                If Not Intersect(target, wks.Range(NAME_CELL_PARALLAX)) Is Nothing Then
                    modRooms.ApplyParallaxRangeCover wks
                End If
            End If
        End If
    End If
End Sub

Private Sub App_SheetBeforeRightClick(ByVal Sh As Object, ByVal target As Range, Cancel As Boolean)
    Debug.Print "App_SheetBeforeRightClick()"
    clsState.CellCtxMnuNeedsPrepare = True

    EvaluateCellCtxMnu Sh, target
  
    clsState.InvalidateControl "RB75dd2c44_btnDynCtxMnu1"
    clsState.InvalidateControl "RB75dd2c44_btnDynCtxMnu2"
    
End Sub



Private Sub App_WorkbookBeforeSave(ByVal wb As Workbook, ByVal SaveAsUI As Boolean, Cancel As Boolean)
    If Not (wb Is RDDAddInWkBk) Then

        Call SaveWorkbookOptions(wb)

    End If
End Sub


