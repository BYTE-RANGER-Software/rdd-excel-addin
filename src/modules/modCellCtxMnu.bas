Attribute VB_Name = "modCellCtxMnu"
Option Explicit
Option Private Module

' ============================================================================
' Module: modCellContextMenu
' Purpose: Encapsulates building, caching, and showing/hiding of the Excel
'          "Cell" CommandBar context menu, including default entries handling
'          and whitelist-based re-showing of selected items.
'
' Public API:
'   - InitCellCtxMenu()
'   - EnsureCellCtxMenuReady()
'   - EvaluateCellCtxMenu(wks As Worksheet, Target As Range) As Integer
'   - ShowCellCtxByCachedCaption(part As String)
'   - ShowAllCachedCellCtx()
'
' Dependencies:
'   - Uses Application.CommandBars("Cell")
'   - Expects clsState global with:
'       * CellCtxMnuNeedsPrepare As Boolean
'       * CellCtxMnuHideDefault As Boolean
'       * CellCtxMenuType As Integer
'   - Expects modRooms.IsRoomSheet(wks As Worksheet) As Boolean
'   - Expects named range constants:
'       * NAME_LIST_ROOM_IDS, NAME_LIST_OBJECTS, NAME_LIST_ACTORS
'
' ============================================================================

' ================================
' Public API – Types
' ================================
Public Enum CellCtxMnu
    CCM_Default      ' Default menu behavior
    CCM_Rooms        ' Context menu for cells validated against room IDs
    CCM_Objects      ' Context menu for cells validated against objects list
    CCM_Actors       ' Context menu for cells validated against actors list
End Enum

' ================================
' Private module state
' ================================
Private m_isCtxCacheInitialized As Boolean               ' Cache guard, True once m_cellCtxControls/m_cellCtxControls reflect the current "Cell" menu
Private m_controlCountSignature As Long                  ' Signature of live controls count to detect Excel-driven menu mutations
Private m_cellCtxControls() As CommandBarControl         ' Cached controls of the "Cell" CommandBar, 1-based, index aligned with m_cellCtxCaptions()
Private m_cellCtxCaptions()  As String                   ' Cached captions for the corresponding controls, used for case-insensitive matching
Private m_ctxMenuWhitelist As Variant                    ' Whitelist of caption substrings to re-show after hiding built-in entries

' ================================
' Public API – Initialization & state
' ================================
' -----------------------------------------------------------------------------------
' Procedure : InitCellCtxMenu
' Purpose   : Initializes the cell context menu handling, builds the cache of the
'             current "Cell" CommandBar entries, hides built-in items, and re-shows
'             a whitelisted subset by caption.
'
' Params    : (none)
' Returns   : (none)
' Notes     :
'   - Populates m_cellCtxControls/m_cellCtxControls via BuildCellCtxMenuCache.
'   - Hides built-in entries, then re-shows captions in m_ctxMenuWhitelist.
'   - Stores a control-count signature (m_controlCountSignature) to detect later changes.
'   - Resets clsState.CellCtxMnuNeedsPrepare to False.
' -----------------------------------------------------------------------------------
Public Sub InitCellCtxMenu()
    On Error GoTo ErrHandler
    
    Dim itemCaption As Variant
    
    BuildCellCtxMenuCache
    HideAllBuiltInCellCtxMenuEntries

    m_ctxMenuWhitelist = Array("K&opieren", "&Copy", "Kom&mentar einfügen", "Insert Co&mment", "Neuer Kommentar", "New Co&mment", "Neue Notiz", "&New Note", "K&ommentare ein-/ausblenden", "&Kommentar löschen")

    For Each itemCaption In m_ctxMenuWhitelist
        ShowCellCtxByCachedCaption itemCaption
    Next
    
    m_controlCountSignature = Application.CommandBars("Cell").Controls.Count
    clsState.CellCtxMnuNeedsPrepare = False
CleanExit:
    Exit Sub
ErrHandler:
    MsgBox "Error (" & Err.Number & "): " & Err.Description, vbCritical, "InitCellCtxMenu"
    LogError "InitCellCtxMenu", Err.Number, Erl
    Resume CleanExit
End Sub

' -----------------------------------------------------------------------------------
' Procedure : EnsureCellCtxMenuReady
' Purpose   : Ensures the cell context menu is prepared for the current state,
'             optionally rebuilding the cache and re-applying visibility rules if
'             Excel has modified the menu for the current context.
'
' Params    : (none)
' Returns   : (none)
' Behavior  :
'   - Exits early if clsState.CellCtxMnuNeedsPrepare is False.
'   - If the live control count differs from m_controlCountSignature, rebuilds cache and
'     re-applies default visibility and whitelist.
'   - If clsState.CellCtxMnuHideDefault is True, hides built-in items again.
'
' Notes     :
'   - Updates m_controlCountSignature to the current control count.
'   - Resets clsState.CellCtxMnuNeedsPrepare to False at the end.
' -----------------------------------------------------------------------------------
Public Sub EnsureCellCtxMenuReady()
    On Error GoTo ErrHandler
    
    If Not clsState.CellCtxMnuNeedsPrepare Then Exit Sub

    Dim cellCommandBar As CommandBar
    Dim itemCaption As Variant
    Set cellCommandBar = Application.CommandBars("Cell")

    ' Menu has been modified by Excel for the current context,
    ' then rebuild context menu cache
    If cellCommandBar.Controls.Count <> m_controlCountSignature Then
        m_isCtxCacheInitialized = False
        BuildCellCtxMenuCache
        'Hide everything built-in (RibbonX buttons remain visible)
        HideAllBuiltInCellCtxMenuEntries
        ' Show standards again (by caption substring, case-insensitive)
        For Each itemCaption In m_ctxMenuWhitelist
            ShowCellCtxByCachedCaption itemCaption
        Next
        m_controlCountSignature = cellCommandBar.Controls.Count
    ElseIf clsState.CellCtxMnuHideDefault Then
        HideAllBuiltInCellCtxMenuEntries
        clsState.CellCtxMnuHideDefault = False
    End If

    clsState.CellCtxMnuNeedsPrepare = False
CleanExit:
    Exit Sub
ErrHandler:
    MsgBox "Error (" & Err.Number & "): " & Err.Description, vbCritical, "EnsureCellCtxMenuReady"
    LogError "EnsureCellCtxMenuReady", Err.Number, Erl
    Resume CleanExit
End Sub

' -----------------------------------------------------------------------------------
' Function  : EvaluateCellCtxMenu
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
Public Function EvaluateCellCtxMenu(wks As Worksheet, Target As Range) As Integer
    On Error GoTo ErrHandler
     
    Dim validationTarget As Validation
    clsState.CellCtxMenuType = CCM_Default
    
    If modRooms.IsRoomSheet(wks) Then
        On Error Resume Next
        Set validationTarget = Target.Validation
        If Not validationTarget Is Nothing Then
            If validationTarget.Type = xlValidateList Then
                If Left(validationTarget.Formula1, 1) = "=" Then
                    Dim nameRef As String
                    nameRef = Mid(validationTarget.Formula1, 2) ' remove '='
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
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo ErrHandler
    End If
    
    EvaluateCellCtxMenu = clsState.CellCtxMenuType
    
    If clsState.CellCtxMenuType = CCM_Default Then
        ' Update cache for default case
        If Not m_isCtxCacheInitialized Then BuildCellCtxMenuCache
        
        If Application.CommandBars("Cell").Controls.Count <> m_controlCountSignature Then
            m_isCtxCacheInitialized = False
            BuildCellCtxMenuCache
            m_controlCountSignature = Application.CommandBars("Cell").Controls.Count
        End If
        ShowAllCachedCellCtx
    Else
        clsState.CellCtxMnuHideDefault = True
    End If
    
CleanExit:
    Exit Function
ErrHandler:
    MsgBox "Error (" & Err.Number & "): " & Err.Description, vbCritical, "EvaluateCellCtxMnu"
    LogError "EvaluateCellCtxMenu", Err.Number, Erl
    Resume CleanExit
End Function

' ================================
' Public API – Cache-based show helpers
' ================================
' -----------------------------------------------------------------------------------
' Procedure : ShowCellCtxByCachedCaption
' Purpose   : Shows cached context menu entries whose stored captions contain the
'             given substring (case-insensitive).
'
' Parameters:
'   part [String] - Caption substring to match (case-insensitive)
' Returns  : (none)
' Notes:
'   - Operates on the cached captions m_cellCtxCaptions() and corresponding controls m_cellCtxControls().
'   - Uses On Error Resume Next for robustness.
' -----------------------------------------------------------------------------------
Public Sub ShowCellCtxByCachedCaption(ByVal part As String)
    On Error GoTo ErrHandler

    Dim Idx As Long, strPrt As String
    If (Not Not m_cellCtxControls) = 0 Then Exit Sub
    On Error Resume Next
    For Idx = 1 To UBound(m_cellCtxControls)
        If InStr(1, m_cellCtxCaptions(Idx), part, vbTextCompare) > 0 Then m_cellCtxControls(Idx).Visible = True
    Next
    
CleanExit:
    Exit Sub
ErrHandler:
    MsgBox "Error (" & Err.Number & "): " & Err.Description, vbCritical, "ShowCellCtxByCachedCaption"
    Resume CleanExit
End Sub

' -----------------------------------------------------------------------------------
' Procedure : ShowAllCachedCellCtx
' Purpose   : Makes all cached context menu entries visible again.
'
' Params    : (none)
' Returns   : (none)
' Notes:
'   - Operates on the cached control array m_cellCtxControls().
'   - Uses On Error Resume Next for robustness.
' -----------------------------------------------------------------------------------
Public Sub ShowAllCachedCellCtx()
    Dim Idx As Long
    On Error Resume Next
    For Idx = 1 To UBound(m_cellCtxControls)
        m_cellCtxControls(Idx).Visible = True
    Next
End Sub

' ================================
' Private implementation
' ================================
' -----------------------------------------------------------------------------------
' Procedure : BuildCellCtxMenuCache
' Purpose   : Builds (or rebuilds) an in-memory cache of the current "Cell"
'             CommandBar controls and their captions, for later show/hide operations.
'
' Params    : (none)
' Returns   : (none)
'
' Behavior  :
'   - Reads Application.CommandBars("Cell").Controls into m_cellCtxControls(), stores captions in m_cellCtxCaptions().
'   - Skips rebuilding if m_isCtxCacheInitialized is already True.
'
' Notes     :
'   - Sets m_isCtxCacheInitialized = True after successful caching.
' -----------------------------------------------------------------------------------
Private Sub BuildCellCtxMenuCache()
    Dim cellCommandBar As CommandBar, Idx As Long
    If Not m_isCtxCacheInitialized Then
        Set cellCommandBar = Application.CommandBars("Cell")

        ReDim m_cellCtxControls(1 To cellCommandBar.Controls.Count)
        ReDim m_cellCtxCaptions(1 To cellCommandBar.Controls.Count)

        For Idx = 1 To cellCommandBar.Controls.Count
            Set m_cellCtxControls(Idx) = cellCommandBar.Controls(Idx)
            'If cellCommandBar.Controls(Idx).BuiltIn Then Debug.Print cellCommandBar.Controls(Idx).Caption
            m_cellCtxCaptions(Idx) = m_cellCtxControls(Idx).Caption
        Next
    End If
    m_isCtxCacheInitialized = True
End Sub

' -----------------------------------------------------------------------------------
' Procedure : HideAllBuiltInCellCtxMenuEntries
' Purpose   : Hides all built-in context menu entries using the cached control list.
'
' Params    : (none)
' Returns   : (none)
'
' Notes:
'   - Requires m_cellCtxControls() to be populated by BuildCellCtxMenuCache.
'   - Uses On Error Resume Next to be resilient to stale references.
' -----------------------------------------------------------------------------------
Private Sub HideAllBuiltInCellCtxMenuEntries()
    Dim Idx As Long
    On Error Resume Next
    For Idx = 1 To UBound(m_cellCtxControls)
        If m_cellCtxControls(Idx).BuiltIn Then m_cellCtxControls(Idx).Visible = False
    Next
End Sub

' -----------------------------------------------------------------------------------
' Procedure : HideAllCellCtxMenuEntries
' Purpose   : Hides all current entries of the live "Cell" CommandBar (not the cache).
'
' Params    : (none)
' Returns   : (none)
'
' Notes:
'   - Iterates Application.CommandBars("Cell").Controls directly.
'   - Uses On Error Resume Next for robustness.
' -----------------------------------------------------------------------------------
Private Sub HideAllCellCtxMenuEntries()
    Dim CmdBarCtrl As CommandBarControl
    On Error Resume Next
    For Each CmdBarCtrl In Application.CommandBars("Cell").Controls
        CmdBarCtrl.Visible = False
    Next CmdBarCtrl
End Sub

' -----------------------------------------------------------------------------------
' Procedure : HideCellCtxByCaption
' Purpose   : Hides live context menu entries whose current captions contain the
'             given substring (case-insensitive).
'
' Parameters:
'   part [String] - Caption substring to match (case-insensitive)
' Returns   : (none)
' Notes:
'   - Operates on the live CommandBar controls (not the cache).
'   - Uses On Error Resume Next for robustness.
' -----------------------------------------------------------------------------------
Private Sub HideCellCtxByCaption(ByVal part As String)
    Dim CmdBarCtrl As CommandBarControl
    On Error Resume Next
    For Each CmdBarCtrl In Application.CommandBars("Cell").Controls
        If InStr(1, CmdBarCtrl.Caption, part, vbTextCompare) > 0 Then CmdBarCtrl.Visible = False
        'Debug.Print CmdBarCtrl.Caption
    Next CmdBarCtrl
End Sub

' -----------------------------------------------------------------------------------
' Procedure : ShowCellCtxByCaption
' Purpose   : Shows live context menu entries whose current captions contain the
'             given substring (case-insensitive).
'
' Parameters:
'   part [String] - Caption substring to match (case-insensitive)
' Returns   : (none)
' Notes:
'   - Operates on the live CommandBar controls (not the cache).
'   - Uses On Error Resume Next for robustness.
' -----------------------------------------------------------------------------------
Private Sub ShowCellCtxByCaption(ByVal part As String)
    Dim CmdBarCtrl As CommandBarControl
    On Error Resume Next
    For Each CmdBarCtrl In Application.CommandBars("Cell").Controls
        If InStr(1, CmdBarCtrl.Caption, part, vbTextCompare) > 0 Then CmdBarCtrl.Visible = True
    Next CmdBarCtrl
End Sub

