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
            If Sh.codeName <> SHEET_DISPATCHER And Not modTags.HasSheetTag(Sh, SHEET_DISPATCHER) Then
                If clsState.RoomSheetChanged Then
                    clsState.RoomSheetChanged = False
                    Call UpdateLists
                End If
                If Left$(wks.Name, Len(ROOM_SHEET_PREFIX)) = ROOM_SHEET_PREFIX Then
                    modRangeCover.EnsureRangeCover wks, wks.Range(NAME_RANGE_ADD_PARALLAX), ROOM_ADD_PARALLAX_SET_COVER_NAME
                    modRangeCover.ApplyCoverVisibilityByDropdown wks, NAME_CELL_PARALLAX, ROOM_ADD_PARALLAX_SET_HIDE_TOKEN, ROOM_ADD_PARALLAX_SET_COVER_NAME
                End If
            End If
        End If
    End If
    
    clsState.InvalidateRibbon
End Sub

Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    Dim wb As Workbook: Set wb = Sh.Parent
    Dim wks As Worksheet: Set wks = Sh
    
    
    ' Only react if the sheet is NOT in the add-in.
    If Not wb Is RDDAddInWkBk And Not wb.IsAddin Then
        ' Check whether the target workbook has a sheet with the Tag SHEET_LISTS
        If modTags.SheetWithTagExists(wb, SHEET_DISPATCHER) Then
            ' Only set the change flag if a room sheet has been changed.
            If Left$(wks.Name, Len(ROOM_SHEET_PREFIX)) = ROOM_SHEET_PREFIX Then
                clsState.RoomSheetChanged = True
                If Not Intersect(Target, wks.Range(NAME_CELL_PARALLAX)) Is Nothing Then
                    modRangeCover.EnsureRangeCover wks, wks.Range(NAME_RANGE_ADD_PARALLAX), ROOM_ADD_PARALLAX_SET_COVER_NAME
                    modRangeCover.ApplyCoverVisibilityByDropdown wks, NAME_CELL_PARALLAX, ROOM_ADD_PARALLAX_SET_HIDE_TOKEN, ROOM_ADD_PARALLAX_SET_COVER_NAME
                End If
            End If
        End If
    End If
End Sub

Private Sub App_SheetBeforeRightClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    Debug.Print "App_SheetBeforeRightClick()"
    clsState.CellCtxMnuNeedsPrepare = True

    EvaluateCellCtxMnu Sh, Target
  
    clsState.InvalidateControl "RB75dd2c44_btnDynCtxMnu1"
    clsState.InvalidateControl "RB75dd2c44_btnDynCtxMnu2"
    
End Sub



Private Sub App_WorkbookBeforeSave(ByVal wb As Workbook, ByVal SaveAsUI As Boolean, Cancel As Boolean)
    If Not (wb Is RDDAddInWkBk) Then

        Call SaveWorkbookOptions(wb)

    End If
End Sub


