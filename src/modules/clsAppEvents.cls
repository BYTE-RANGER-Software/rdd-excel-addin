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

Private Sub App_SheetActivate(ByVal sh As Object)
    Dim wb As Workbook: Set wb = sh.Parent
    Dim wks As Worksheet: Set wks = sh
    
    
    ' Only react if the sheet is NOT in the add-in.
    If Not wb Is RDDAddInWkBk And Not wb.IsAddin Then
        ' Check whether the target workbook has a sheet with the code name “Lists”
        If Not modUtil.GetSheetByCodeName(SHEET_LISTS) Is Nothing Then
            'Don't update on sheet Lists
            If sh.codeName <> SHEET_LISTS Then
                If gblnRoomSheetChanged Then
                    gblnRoomSheetChanged = False
                    Call UpdateLists
                End If
                If Left$(wks.name, Len(ROOM_SHEET_PREFIX)) = ROOM_SHEET_PREFIX Then
                    modRangeCover.EnsureRangeCover wks, wks.Range(ROOM_ADD_PARALLAX_SET_RANGE), ROOM_ADD_PARALLAX_SET_COVER_NAME
                    modRangeCover.ApplyCoverVisibilityByDropdown wks, ROOM_PARALLAX_CELL, ROOM_ADD_PARALLAX_SET_HIDE_TOKEN, ROOM_ADD_PARALLAX_SET_COVER_NAME
                End If
            End If
        End If
    End If
End Sub

Private Sub App_SheetChange(ByVal sh As Object, ByVal Target As Range)
    Dim wb As Workbook: Set wb = sh.Parent
    Dim wks As Worksheet: Set wks = sh
    
    
    ' Only react if the sheet is NOT in the add-in.
    If Not wb Is RDDAddInWkBk And Not wb.IsAddin Then
        ' Check whether the target workbook has a sheet with the code name “Lists”
        If Not modUtil.GetSheetByCodeName(SHEET_LISTS) Is Nothing Then
            ' Only set the change flag if a room sheet has been changed.
            If Left$(wks.name, Len(ROOM_SHEET_PREFIX)) = ROOM_SHEET_PREFIX Then
                gblnRoomSheetChanged = True
                If Not Intersect(Target, wks.Range(ROOM_PARALLAX_CELL)) Is Nothing Then
                    modRangeCover.EnsureRangeCover wks, wks.Range(ROOM_ADD_PARALLAX_SET_RANGE), ROOM_ADD_PARALLAX_SET_COVER_NAME
                    modRangeCover.ApplyCoverVisibilityByDropdown wks, ROOM_PARALLAX_CELL, ROOM_ADD_PARALLAX_SET_HIDE_TOKEN, ROOM_ADD_PARALLAX_SET_COVER_NAME
                End If
            End If
        End If
    End If
End Sub

Private Sub App_SheetBeforeRightClick(ByVal sh As Object, ByVal Target As Range, Cancel As Boolean)

    gblnCellCtxMnuNeedsPrepare = True

    
    EvaluateCellCtxMnu sh, Target

    
    If Not gRibbonUI Is Nothing Then
        gRibbonUI.InvalidateControl "RB75dd2c44_btnDynCtxMnu1"
        gRibbonUI.InvalidateControl "RB75dd2c44_btnDynCtxMnu2"
    End If
End Sub






