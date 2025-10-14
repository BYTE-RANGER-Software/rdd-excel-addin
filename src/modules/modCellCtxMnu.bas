Attribute VB_Name = "modCellCtxMnu"
Option Explicit
Option Private Module

Public Enum CellCtxMnu
    CellCtxMnu_Default      ' Default menu behavior
    CellCtxMnu_Rooms        ' Context menu for cells validated against room IDs
    CellCtxMnu_Objects      ' Context menu for cells validated against objects list
    CellCtxMnu_Actors       ' Context menu for cells validated against actors list
End Enum

Public gblnCellCtxMnuNeedsPrepare As Boolean        ' True to (re)prepare the menu on next call when Excel changed the "Cell" CommandBar
Public gintCellCtxMenuType As Integer               ' Resolved context menu type according to CellCtxMnu, used by EvaluateCellCtxMnu and callers
Public gblnCellCtxMnuHideDefault As Boolean         ' If True, built-in entries will be hidden during the next prepare pass

Private mblnCtxCacheInited As Boolean               ' Cache guard, True once macbcCellCtxCtrls/mastrCellCtxCaps reflect the current "Cell" menu
Private mlngCtlCountSig As Long                     ' Signature of live controls count to detect Excel-driven menu mutations
Private macbcCellCtxCtrls() As CommandBarControl    ' Cached controls of the "Cell" CommandBar, 1-based, index aligned with mastrCellCtxCaps()
Private mastrCellCtxCaps()  As String               ' Cached captions for the corresponding controls, used for case-insensitive matching
Private mvntCtxMenuWhitelist As Variant             ' Whitelist of caption substrings to re-show after hiding built-in entries


' -----------------------------------------------------------------------------------
' Procedure : InitCellCtxMnu
' Purpose   : Initializes the cell context menu handling, builds the cache of the
'             current "Cell" CommandBar entries, hides built-in items, and re-shows
'             a whitelisted subset by caption.
'
' Parameters:
'   (none)
'
' Notes     :
'   - Populates macbcCellCtxCtrls/mastrCellCtxCaps via BuildCellCtxMenuCache.
'   - Hides built-in entries, then re-shows captions in mvntCtxMenuWhitelist.
'   - Stores a control-count signature (mlngCtlCountSig) to detect later changes.
'   - Resets gblnCellCtxMnuNeedsPrepare to False.
' -----------------------------------------------------------------------------------
Public Sub InitCellCtxMnu()
    Dim vntCtxCpt As Variant
    BuildCellCtxMenuCache
    HideAllBuildInCellCtxMenuEntries

    mvntCtxMenuWhitelist = Array("K&opieren", "&Copy", "Kom&mentar einfügen", "Insert Co&mment", "Neuer Kommentar", "New Co&mment", "Neue Notiz", "&New Note")

    For Each vntCtxCpt In mvntCtxMenuWhitelist
        ShowCellCtxByCachedCaption vntCtxCpt
    Next
    
    mlngCtlCountSig = Application.CommandBars("Cell").Controls.Count
    gblnCellCtxMnuNeedsPrepare = False
End Sub

' -----------------------------------------------------------------------------------
' Procedure : EnsureCellCtxMnuReady
' Purpose   : Ensures the cell context menu is prepared for the current state,
'             optionally rebuilding the cache and re-applying visibility rules if
'             Excel has modified the menu for the current context.
'
' Parameters:
'   (none)
'
' Behavior  :
'   - Exits early if gblnCellCtxMnuNeedsPrepare is False.
'   - If the live control count differs from mlngCtlCountSig, rebuilds cache and
'     re-applies default visibility and whitelist.
'   - If gblnCellCtxMnuHideDefault is True, hides built-in items again.
'
' Notes     :
'   - Updates mlngCtlCountSig to the current control count.
'   - Resets gblnCellCtxMnuNeedsPrepare to False at the end.
' -----------------------------------------------------------------------------------
Public Sub EnsureCellCtxMnuReady()
    If Not gblnCellCtxMnuNeedsPrepare Then Exit Sub

    Dim cbrCell As CommandBar
    Dim vntCtxCpt As Variant
    Set cbrCell = Application.CommandBars("Cell")

    ' Menu has been modified by Excel for the current context,
    ' then rebuild context menu cache
    If cbrCell.Controls.Count <> mlngCtlCountSig Then
        mblnCtxCacheInited = False
        BuildCellCtxMenuCache
        'Hide everything built-in (RibbonX buttons remain visible)
        HideAllBuildInCellCtxMenuEntries
        ' Show standards again (by caption substring, case-insensitive)
        For Each vntCtxCpt In mvntCtxMenuWhitelist
            ShowCellCtxByCachedCaption vntCtxCpt
        Next
        mlngCtlCountSig = cbrCell.Controls.Count
    ElseIf gblnCellCtxMnuHideDefault Then
        HideAllBuildInCellCtxMenuEntries
        gblnCellCtxMnuHideDefault = False
    End If

    gblnCellCtxMnuNeedsPrepare = False
End Sub

' -----------------------------------------------------------------------------------
' Function  : EvaluateCellCtxMnu
' Purpose   : Determines the context menu type for the given cell based on the
'             worksheet name prefix and data validation (list) source, then prepares
'             the menu accordingly.
'
' Parameters:
'   ws     [Worksheet] - Worksheet containing the target cell
'   Target [Range]     - Target cell whose context is evaluated
'
' Returns:
'   Integer - One of the CellCtxMnu enumeration values
'             (CellCtxMnu_Default, CellCtxMnu_Rooms, CellCtxMnu_Objects, CellCtxMnu_Actors).
'
' Behavior  :
'   - If the sheet name starts with ROOM_SHEET_PREFIX and the target has a List validation
'     driven by a named range, maps NAME_LIST_ROOM_IDS / NAME_LIST_OBJECTS / NAME_LIST_ACTORS
'     to the respective menu type.
'   - For Default: ensures cache exists and matches the live menu, then shows all cached items.
'   - For non-default: sets gblnCellCtxMnuHideDefault to True to hide built-ins on prepare.
'
' Notes:
'   - Uses On Error Resume Next around Validation access.
'   - Updates global gintCellCtxMenuType and returns it.
' -----------------------------------------------------------------------------------
Public Function EvaluateCellCtxMnu(ws As Worksheet, Target As Range) As Integer
    On Error Resume Next
    
    Dim vldTarget As Validation
    gintCellCtxMenuType = CellCtxMnu_Default
    
    If Left$(ws.name, Len(ROOM_SHEET_PREFIX)) = ROOM_SHEET_PREFIX Then
        Set vldTarget = Target.Validation
        If vldTarget.Type = xlValidateList Then
        
            If Left(vldTarget.Formula1, 1) = "=" Then
                Dim nameRef As String
                nameRef = Mid(vldTarget.Formula1, 2) ' remove '='
                If nameRef = NAME_LIST_ROOM_IDS Then
                    gintCellCtxMenuType = CellCtxMnu_Rooms
                ElseIf nameRef = NAME_LIST_OBJECTS Then
                    gintCellCtxMenuType = CellCtxMnu_Objects
                ElseIf nameRef = NAME_LIST_ACTORS Then
                    gintCellCtxMenuType = CellCtxMnu_Actors
                End If
            End If
        End If
    End If
    EvaluateCellCtxMnu = gintCellCtxMenuType
    
    If gintCellCtxMenuType = CellCtxMnu_Default Then
        ' Cache für Default-Fall aktualisieren
        If Not mblnCtxCacheInited Then BuildCellCtxMenuCache
        
        If Application.CommandBars("Cell").Controls.Count <> mlngCtlCountSig Then
            mblnCtxCacheInited = False
            BuildCellCtxMenuCache
            mlngCtlCountSig = Application.CommandBars("Cell").Controls.Count
        End If
        ShowAllCachedCellCtx
    Else
        gblnCellCtxMnuHideDefault = True
    End If
    
    
End Function

' -----------------------------------------------------------------------------------
' Procedure : BuildCellCtxMenuCache
' Purpose   : Builds (or rebuilds) an in-memory cache of the current "Cell"
'             CommandBar controls and their captions, for later show/hide operations.
'
' Parameters:
'   (none)
'
' Behavior  :
'   - Reads Application.CommandBars("Cell").Controls into macbcCellCtxCtrls(), stores captions in mastrCellCtxCaps().
'   - Skips rebuilding if mblnCtxCacheInited is already True.
'
' Notes     :
'   - Sets mblnCtxCacheInited = True after successful caching.
'   - Debug.Prints captions of built-in controls (for diagnostics).
' -----------------------------------------------------------------------------------
Private Sub BuildCellCtxMenuCache()
    Dim cbrCell As CommandBar, lngIdx As Long
    If Not mblnCtxCacheInited Then
        Set cbrCell = Application.CommandBars("Cell")

        ReDim macbcCellCtxCtrls(1 To cbrCell.Controls.Count)
        ReDim mastrCellCtxCaps(1 To cbrCell.Controls.Count)

        For lngIdx = 1 To cbrCell.Controls.Count
            Set macbcCellCtxCtrls(lngIdx) = cbrCell.Controls(lngIdx)
            If cbrCell.Controls(lngIdx).BuiltIn Then Debug.Print cbrCell.Controls(lngIdx).Caption
            mastrCellCtxCaps(lngIdx) = macbcCellCtxCtrls(lngIdx).Caption
        Next
    End If
    mblnCtxCacheInited = True
End Sub

' -----------------------------------------------------------------------------------
' Procedure : HideAllBuildInCellCtxMenuEntries
' Purpose   : Hides all built-in context menu entries using the cached control list.
'
' Parameters:
'   (none)
'
' Notes:
'   - Requires macbcCellCtxCtrls() to be populated by BuildCellCtxMenuCache.
'   - Uses On Error Resume Next to be resilient to stale references.
' -----------------------------------------------------------------------------------
Private Sub HideAllBuildInCellCtxMenuEntries()
    Dim lngIdx As Long
    On Error Resume Next
    For lngIdx = 1 To UBound(macbcCellCtxCtrls)
        If macbcCellCtxCtrls(lngIdx).BuiltIn Then macbcCellCtxCtrls(lngIdx).Visible = False
    Next
End Sub

' -----------------------------------------------------------------------------------
' Procedure : ShowCellCtxByCachedCaption
' Purpose   : Shows cached context menu entries whose stored captions contain the
'             given substring (case-insensitive).
'
' Parameters:
'   part [String] - Caption substring to match (case-insensitive)
'
' Notes:
'   - Operates on the cached captions mastrCellCtxCaps() and corresponding controls macbcCellCtxCtrls().
'   - Uses On Error Resume Next for robustness.
' -----------------------------------------------------------------------------------
Public Sub ShowCellCtxByCachedCaption(ByVal part As String)
    Dim lngIdx As Long, strPrt As String
    strPrt = part
    On Error Resume Next
    For lngIdx = 1 To UBound(mastrCellCtxCaps)
        If InStr(1, mastrCellCtxCaps(lngIdx), strPrt, vbTextCompare) > 0 Then macbcCellCtxCtrls(lngIdx).Visible = True
    Next
End Sub

' -----------------------------------------------------------------------------------
' Procedure : ShowAllCachedCellCtx
' Purpose   : Makes all cached context menu entries visible again.
'
' Parameters:
'   (none)
'
' Notes:
'   - Operates on the cached control array macbcCellCtxCtrls().
'   - Uses On Error Resume Next for robustness.
' -----------------------------------------------------------------------------------
Public Sub ShowAllCachedCellCtx()
    Dim lngIdx As Long
    On Error Resume Next
    For lngIdx = 1 To UBound(macbcCellCtxCtrls)
        macbcCellCtxCtrls(lngIdx).Visible = True
    Next
End Sub

' -----------------------------------------------------------------------------------
' Procedure : HideAllCellCtxMnuEntries
' Purpose   : Hides all current entries of the live "Cell" CommandBar (not the cache).
'
' Parameters:
'   (none)
'
' Notes:
'   - Iterates Application.CommandBars("Cell").Controls directly.
'   - Uses On Error Resume Next for robustness.
' -----------------------------------------------------------------------------------
Private Sub HideAllCellCtxMnuEntries()
    Dim cbcCtrl As CommandBarControl
    On Error Resume Next
    For Each cbcCtrl In Application.CommandBars("Cell").Controls
        cbcCtrl.Visible = False
    Next cbcCtrl
End Sub

' -----------------------------------------------------------------------------------
' Procedure : HideCellCtxByCaption
' Purpose   : Hides live context menu entries whose current captions contain the
'             given substring (case-insensitive).
'
' Parameters:
'   part [String] - Caption substring to match (case-insensitive)
'
' Notes:
'   - Operates on the live CommandBar controls (not the cache).
'   - Debug.Prints each processed caption for diagnostics.
'   - Uses On Error Resume Next for robustness.
' -----------------------------------------------------------------------------------
Private Sub HideCellCtxByCaption(ByVal part As String)
    Dim cbcCtrl As CommandBarControl
    On Error Resume Next
    For Each cbcCtrl In Application.CommandBars("Cell").Controls
        If InStr(1, cbcCtrl.Caption, part, vbTextCompare) > 0 Then cbcCtrl.Visible = False
        'Debug.Print cbcCtrl.Caption
    Next cbcCtrl
End Sub

' -----------------------------------------------------------------------------------
' Procedure : ShowCellCtxByCaption
' Purpose   : Shows live context menu entries whose current captions contain the
'             given substring (case-insensitive).
'
' Parameters:
'   part [String] - Caption substring to match (case-insensitive)
'
' Notes:
'   - Operates on the live CommandBar controls (not the cache).
'   - Uses On Error Resume Next for robustness.
' -----------------------------------------------------------------------------------
Private Sub ShowCellCtxByCaption(ByVal part As String)
    Dim cbcCtrl As CommandBarControl
    On Error Resume Next
    For Each cbcCtrl In Application.CommandBars("Cell").Controls
        If InStr(1, cbcCtrl.Caption, part, vbTextCompare) > 0 Then cbcCtrl.Visible = True
    Next cbcCtrl
End Sub

