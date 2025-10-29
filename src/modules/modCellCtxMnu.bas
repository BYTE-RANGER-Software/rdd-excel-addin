Attribute VB_Name = "modCellCtxMnu"
Option Explicit
Option Private Module

Public Enum CellCtxMnu
    CCM_Default      ' Default menu behavior
    CCM_Rooms        ' Context menu for cells validated against room IDs
    CCM_Objects      ' Context menu for cells validated against objects list
    CCM_Actors       ' Context menu for cells validated against actors list
End Enum

Private m_blnCtxCacheInited As Boolean               ' Cache guard, True once m_acbcCellCtxCtrls/m_astrCellCtxCaps reflect the current "Cell" menu
Private m_lngCtlCountSig As Long                     ' Signature of live controls count to detect Excel-driven menu mutations
Private m_acbcCellCtxCtrls() As CommandBarControl    ' Cached controls of the "Cell" CommandBar, 1-based, index aligned with m_astrCellCtxCaps()
Private m_astrCellCtxCaps()  As String               ' Cached captions for the corresponding controls, used for case-insensitive matching
Private m_vntCtxMenuWhitelist As Variant             ' Whitelist of caption substrings to re-show after hiding built-in entries


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
'   - Populates m_acbcCellCtxCtrls/m_astrCellCtxCaps via BuildCellCtxMenuCache.
'   - Hides built-in entries, then re-shows captions in m_vntCtxMenuWhitelist.
'   - Stores a control-count signature (m_lngCtlCountSig) to detect later changes.
'   - Resets clsState.CellCtxMnuNeedsPrepare to False.
' -----------------------------------------------------------------------------------
Public Sub InitCellCtxMnu()
    Dim vntCtxCpt As Variant
    BuildCellCtxMenuCache
    HideAllBuildInCellCtxMenuEntries

    m_vntCtxMenuWhitelist = Array("K&opieren", "&Copy", "Kom&mentar einfügen", "Insert Co&mment", "Neuer Kommentar", "New Co&mment", "Neue Notiz", "&New Note", "K&ommentare ein-/ausblenden", "&Kommentar löschen")

    For Each vntCtxCpt In m_vntCtxMenuWhitelist
        ShowCellCtxByCachedCaption vntCtxCpt
    Next
    
    m_lngCtlCountSig = Application.CommandBars("Cell").Controls.Count
    clsState.CellCtxMnuNeedsPrepare = False
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
'   - Exits early if clsState.CellCtxMnuNeedsPrepare is False.
'   - If the live control count differs from m_lngCtlCountSig, rebuilds cache and
'     re-applies default visibility and whitelist.
'   - If clsState.CellCtxMnuHideDefault is True, hides built-in items again.
'
' Notes     :
'   - Updates m_lngCtlCountSig to the current control count.
'   - Resets clsState.CellCtxMnuNeedsPrepare to False at the end.
' -----------------------------------------------------------------------------------
Public Sub EnsureCellCtxMnuReady()
    If Not clsState.CellCtxMnuNeedsPrepare Then Exit Sub

    Dim cbrCell As CommandBar
    Dim vntCtxCpt As Variant
    Set cbrCell = Application.CommandBars("Cell")

    ' Menu has been modified by Excel for the current context,
    ' then rebuild context menu cache
    If cbrCell.Controls.Count <> m_lngCtlCountSig Then
        m_blnCtxCacheInited = False
        BuildCellCtxMenuCache
        'Hide everything built-in (RibbonX buttons remain visible)
        HideAllBuildInCellCtxMenuEntries
        ' Show standards again (by caption substring, case-insensitive)
        For Each vntCtxCpt In m_vntCtxMenuWhitelist
            ShowCellCtxByCachedCaption vntCtxCpt
        Next
        m_lngCtlCountSig = cbrCell.Controls.Count
    ElseIf clsState.CellCtxMnuHideDefault Then
        HideAllBuildInCellCtxMenuEntries
        clsState.CellCtxMnuHideDefault = False
    End If

    clsState.CellCtxMnuNeedsPrepare = False
End Sub

' -----------------------------------------------------------------------------------
' Function  : EvaluateCellCtxMnu
' Purpose   : Determines the context menu type for the given cell based on the
'             worksheet name prefix and data validation (list) source, then prepares
'             the menu accordingly.
'
' Parameters:
'   wks     [Worksheet] - Worksheet containing the target cell
'   Target [Range]     - Target cell whose context is evaluated
'
' Returns:
'   Integer - One of the CellCtxMnu enumeration values
'             (CCM_Default, CCM_Rooms, CCM_Objects, CCM_Actors).
'
' Behavior  :
'   - If the sheet name starts with ROOM_SHEET_PREFIX and the target has a List validation
'     driven by a named range, maps NAME_LIST_ROOM_IDS / NAME_LIST_OBJECTS / NAME_LIST_ACTORS
'     to the respective menu type.
'   - For Default: ensures cache exists and matches the live menu, then shows all cached items.
'   - For non-default: sets clsState.CellCtxMnuHideDefault to True to hide built-ins on prepare.
'
' Notes:
'   - Uses On Error Resume Next around Validation access.
'   - Updates global clsState.CellCtxMenuType and returns it.
' -----------------------------------------------------------------------------------
Public Function EvaluateCellCtxMnu(wks As Worksheet, target As Range) As Integer
    On Error Resume Next
    
    Dim vldTarget As Validation
    clsState.CellCtxMenuType = CCM_Default
    
    If modRooms.IsRoomSheet(wks) Then
        Set vldTarget = target.Validation
        If vldTarget.Type = xlValidateList Then
        
            If Left(vldTarget.Formula1, 1) = "=" Then
                Dim nameRef As String
                nameRef = Mid(vldTarget.Formula1, 2) ' remove '='
                If nameRef = NAME_LIST_ROOM_IDS Then
                    clsState.CellCtxMenuType = CCM_Rooms
                ElseIf nameRef = NAME_LIST_OBJECTS Then
                    clsState.CellCtxMenuType = CCM_Objects
                ElseIf nameRef = NAME_LIST_ACTORS Then
                    clsState.CellCtxMenuType = CCM_Actors
                End If
            End If
        End If
    End If
    EvaluateCellCtxMnu = clsState.CellCtxMenuType
    
    If clsState.CellCtxMenuType = CCM_Default Then
        ' Cache für Default-Fall aktualisieren
        If Not m_blnCtxCacheInited Then BuildCellCtxMenuCache
        
        If Application.CommandBars("Cell").Controls.Count <> m_lngCtlCountSig Then
            m_blnCtxCacheInited = False
            BuildCellCtxMenuCache
            m_lngCtlCountSig = Application.CommandBars("Cell").Controls.Count
        End If
        ShowAllCachedCellCtx
    Else
        clsState.CellCtxMnuHideDefault = True
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
'   - Reads Application.CommandBars("Cell").Controls into m_acbcCellCtxCtrls(), stores captions in m_astrCellCtxCaps().
'   - Skips rebuilding if m_blnCtxCacheInited is already True.
'
' Notes     :
'   - Sets m_blnCtxCacheInited = True after successful caching.
'   - Debug.Prints captions of built-in controls (for diagnostics).
' -----------------------------------------------------------------------------------
Private Sub BuildCellCtxMenuCache()
    Dim cbrCell As CommandBar, lngIdx As Long
    If Not m_blnCtxCacheInited Then
        Set cbrCell = Application.CommandBars("Cell")

        ReDim m_acbcCellCtxCtrls(1 To cbrCell.Controls.Count)
        ReDim m_astrCellCtxCaps(1 To cbrCell.Controls.Count)

        For lngIdx = 1 To cbrCell.Controls.Count
            Set m_acbcCellCtxCtrls(lngIdx) = cbrCell.Controls(lngIdx)
            'If cbrCell.Controls(lngIdx).BuiltIn Then Debug.Print cbrCell.Controls(lngIdx).Caption
            m_astrCellCtxCaps(lngIdx) = m_acbcCellCtxCtrls(lngIdx).Caption
        Next
    End If
    m_blnCtxCacheInited = True
End Sub

' -----------------------------------------------------------------------------------
' Procedure : HideAllBuildInCellCtxMenuEntries
' Purpose   : Hides all built-in context menu entries using the cached control list.
'
' Parameters:
'   (none)
'
' Notes:
'   - Requires m_acbcCellCtxCtrls() to be populated by BuildCellCtxMenuCache.
'   - Uses On Error Resume Next to be resilient to stale references.
' -----------------------------------------------------------------------------------
Private Sub HideAllBuildInCellCtxMenuEntries()
    Dim lngIdx As Long
    On Error Resume Next
    For lngIdx = 1 To UBound(m_acbcCellCtxCtrls)
        If m_acbcCellCtxCtrls(lngIdx).BuiltIn Then m_acbcCellCtxCtrls(lngIdx).Visible = False
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
'   - Operates on the cached captions m_astrCellCtxCaps() and corresponding controls m_acbcCellCtxCtrls().
'   - Uses On Error Resume Next for robustness.
' -----------------------------------------------------------------------------------
Public Sub ShowCellCtxByCachedCaption(ByVal part As String)
    Dim lngIdx As Long, strPrt As String
    strPrt = part
    On Error Resume Next
    For lngIdx = 1 To UBound(m_astrCellCtxCaps)
        If InStr(1, m_astrCellCtxCaps(lngIdx), strPrt, vbTextCompare) > 0 Then m_acbcCellCtxCtrls(lngIdx).Visible = True
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
'   - Operates on the cached control array m_acbcCellCtxCtrls().
'   - Uses On Error Resume Next for robustness.
' -----------------------------------------------------------------------------------
Public Sub ShowAllCachedCellCtx()
    Dim lngIdx As Long
    On Error Resume Next
    For lngIdx = 1 To UBound(m_acbcCellCtxCtrls)
        m_acbcCellCtxCtrls(lngIdx).Visible = True
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

