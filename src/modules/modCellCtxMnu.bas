Attribute VB_Name = "modCellCtxMnu"
' ============================================================================
' Module: modCellContextMenu
' Purpose: Encapsulates building, caching, and showing/hiding of the Excel
'          "Cell" CommandBar context menu, including default entries handling
'          and whitelist-based re-showing of selected items.
'
' Public API:
'   - InitializeCellCtxMenu()
'   - EnsureCellCtxMenuReady()
'   - EvaluateCellCtxMenu(wks As Worksheet, Target As Range) As Integer
'   - ResetToDefaultCtxMenu()
'   - ShowCellCtxByCachedCaption(part As String)
'   - ShowAllCachedCellCtx()
'   - CleanupCellCtxMenu()
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
' Notes:
'
' ============================================================================

Option Explicit
Option Private Module

' ================================
' Public API – Types
' ================================
Public Enum CellCtxMnu
    CCM_Default      ' Default menu behavior
    CCM_Rooms        ' Context menu for cells that intersect Room ID/Alias cells
    CCM_Puzzles = 2  ' Puzzle ID cell
    CCM_Items = 3    ' Item ID/Name cell
    CCM_Actors = 4   ' Actor ID/Name cell
    CCM_Hotspot = 5  ' Hotspot ID/Name cell
    CCM_Flags = 6    ' Flag ID cell
    CCM_Dependencies = 7 ' Puzzle DependsOn and Requires cells
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
' Procedure : InitializeCellCtxMenu
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
Public Sub InitializeCellCtxMenu()
    On Error GoTo ErrHandler
    
    Dim itemCaption As Variant
    
    ' Reset to clean state
    CommandBars("Cell").Reset
    
    ' Build cache from the clean/full menu
    BuildCellCtxMenuCache

    ' Define whitelist
    m_ctxMenuWhitelist = Array("K&opieren", "&Copy", "Kom&mentar einfügen", "Insert Co&mment", "Neuer Kommentar", "New Co&mment", "Neue Notiz", "&New Note", "K&ommentare ein-/ausblenden", "&Kommentar löschen")

    ' Store signature of the FULL menu
    m_controlCountSignature = Application.CommandBars("Cell").Controls.count
    
    ' State is default (full menu visible)
    clsState.CellCtxMenuType = CCM_Default
    clsState.CellCtxMnuNeedsPrepare = False
    clsState.CellCtxMnuHideDefault = False
    
CleanExit:
    Exit Sub
ErrHandler:
    modErr.ReportError "InitializeCellCtxMenu", Err.Number, Erl, caption:=modMain.AppProjectName
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
    If cellCommandBar.Controls.count <> m_controlCountSignature Then
        m_isCtxCacheInitialized = False
        BuildCellCtxMenuCache
        'Hide everything built-in (RibbonX buttons remain visible)
        HideAllBuiltInCellCtxMenuEntries
        ' Show standards again (by caption substring, case-insensitive)
        For Each itemCaption In m_ctxMenuWhitelist
            ShowCellCtxByCachedCaption itemCaption
        Next
        m_controlCountSignature = cellCommandBar.Controls.count
    ElseIf clsState.CellCtxMnuHideDefault Then
        HideAllBuiltInCellCtxMenuEntries
        clsState.CellCtxMnuHideDefault = False
    End If

    clsState.CellCtxMnuNeedsPrepare = False
CleanExit:
    Exit Sub
ErrHandler:
    modErr.ReportError "EnsureCellCtxMenuReady", Err.Number, Erl, caption:=modMain.AppProjectName
    Resume CleanExit
End Sub

' -----------------------------------------------------------------------------------
' Function  : EvaluateCellCtxMenu
' Purpose   : Determines the context menu type for the given cell based on the
'             worksheet name prefix and whether the cell intersects with specific
'             named ranges.
'
' Parameters:
'   sheet     [Worksheet] - Worksheet containing the target cell
'   cell      [Range]     - Target cell whose context is evaluated
'
' Returns:
'   Integer - One of the CellCtxMnu enumeration values
'             (CCM_Default, CCM_Rooms, ...).
'
' Behavior  :
'   - If the sheet is a Room sheet and the cell intersects with
'     a named range (NAME_RANGE_...),
'     maps to the respective menu type.
'   - For Default: ensures cache exists and matches the live menu, then shows all cached items.
'   - For non-default: sets clsState.CellCtxMnuHideDefault to True to hide built-ins on prepare.
'
' Notes:
'   - Uses On Error Resume Next around Named Range access.
'   - Updates global clsState.CellCtxMenuType and returns it.
'   - Order of checks matters: more specific checks should come first.
' -----------------------------------------------------------------------------------
Public Function EvaluateCellCtxMenu(sheet As Worksheet, cell As Range) As Integer
    On Error GoTo ErrHandler
     
    Dim namedRng As Range
    clsState.CellCtxMenuType = CCM_Default
    
    If modRooms.IsRoomSheet(sheet) Then
    
        ' === Check Room References (Doors To) ===
        
        ' Check if cell intersects with NAME_RANGE_DOORS_TO_ROOM_ID
        On Error Resume Next
        Set namedRng = sheet.Names(NAME_RANGE_DOORS_TO_ROOM_ID).RefersToRange
        On Error GoTo ErrHandler
        If Not namedRng Is Nothing Then
            If Not Intersect(cell, namedRng) Is Nothing Then
                clsState.CellCtxMenuType = CCM_Rooms
                GoTo FoundMatch
            End If
        End If
        Set namedRng = Nothing
        
        ' Check if cell intersects with NAME_RANGE_DOORS_TO_ROOM_ALIAS
        On Error Resume Next
        Set namedRng = sheet.Names(NAME_RANGE_DOORS_TO_ROOM_ALIAS).RefersToRange
        On Error GoTo ErrHandler
        If Not namedRng Is Nothing Then
            If Not Intersect(cell, namedRng) Is Nothing Then
                clsState.CellCtxMenuType = CCM_Rooms
                GoTo FoundMatch
            End If
        End If
        Set namedRng = Nothing
        
        ' === Check Puzzle-related ranges ===
        
        ' Check Puzzle ID column
        On Error Resume Next
        Set namedRng = sheet.Names(NAME_RANGE_PUZZLES_PUZZLE_ID).RefersToRange
        On Error GoTo ErrHandler
        If Not namedRng Is Nothing Then
            If Not Intersect(cell, namedRng) Is Nothing Then
                clsState.CellCtxMenuType = CCM_Puzzles
                GoTo FoundMatch
            End If
        End If
        Set namedRng = Nothing
        
        ' Check Puzzle DependsOn column (Dependencies)
        On Error Resume Next
        Set namedRng = sheet.Names(NAME_RANGE_PUZZLES_DEPENDS_ON).RefersToRange
        On Error GoTo ErrHandler
        If Not namedRng Is Nothing Then
            If Not Intersect(cell, namedRng) Is Nothing Then
                clsState.CellCtxMenuType = CCM_Dependencies
                GoTo FoundMatch
            End If
        End If
        Set namedRng = Nothing
        
        ' Check Puzzle Requires column (Dependencies)
        On Error Resume Next
        Set namedRng = sheet.Names(NAME_RANGE_PUZZLES_REQUIRES).RefersToRange
        On Error GoTo ErrHandler
        If Not namedRng Is Nothing Then
            If Not Intersect(cell, namedRng) Is Nothing Then
                clsState.CellCtxMenuType = CCM_Dependencies
                GoTo FoundMatch
            End If
        End If
        Set namedRng = Nothing
        
        ' === Check Item References (PICKUPABLE_OBJECTS) ===
        
        'item ID
        On Error Resume Next
        Set namedRng = sheet.Names(NAME_RANGE_PICKUPABLE_OBJECTS_ITEM_ID).RefersToRange
        On Error GoTo ErrHandler
        If Not namedRng Is Nothing Then
            If Not Intersect(cell, namedRng) Is Nothing Then
                clsState.CellCtxMenuType = CCM_Items
                GoTo FoundMatch
            End If
        End If
        Set namedRng = Nothing
        
        'Item Name
        On Error Resume Next
        Set namedRng = sheet.Names(NAME_RANGE_PICKUPABLE_OBJECTS_NAME).RefersToRange
        On Error GoTo ErrHandler
        If Not namedRng Is Nothing Then
            If Not Intersect(cell, namedRng) Is Nothing Then
                clsState.CellCtxMenuType = CCM_Items
                GoTo FoundMatch
            End If
        End If
        Set namedRng = Nothing
        
        ' === Check Actor References ===
        
        'Actor ID
        On Error Resume Next
        Set namedRng = sheet.Names(NAME_RANGE_ACTORS_ACTOR_ID).RefersToRange
        On Error GoTo ErrHandler
        If Not namedRng Is Nothing Then
            If Not Intersect(cell, namedRng) Is Nothing Then
                clsState.CellCtxMenuType = CCM_Actors
                GoTo FoundMatch
            End If
        End If
        Set namedRng = Nothing
        
        'Actor Name
        On Error Resume Next
        Set namedRng = sheet.Names(NAME_RANGE_ACTORS_ACTOR_NAME).RefersToRange
        On Error GoTo ErrHandler
        If Not namedRng Is Nothing Then
            If Not Intersect(cell, namedRng) Is Nothing Then
                clsState.CellCtxMenuType = CCM_Actors
                GoTo FoundMatch
            End If
        End If
        Set namedRng = Nothing
        
        ' === Check Hotspot References ===
        
        'Hotspot ID
        On Error Resume Next
        Set namedRng = sheet.Names(NAME_RANGE_TOUCHABLE_OBJECTS_HOTSPOT_ID).RefersToRange
        On Error GoTo ErrHandler
        If Not namedRng Is Nothing Then
            If Not Intersect(cell, namedRng) Is Nothing Then
                clsState.CellCtxMenuType = CCM_Hotspot
                GoTo FoundMatch
            End If
        End If
        Set namedRng = Nothing
        
        'Hotspot Name
        On Error Resume Next
        Set namedRng = sheet.Names(NAME_RANGE_TOUCHABLE_OBJECTS_HOTSPOT_NAME).RefersToRange
        On Error GoTo ErrHandler
        If Not namedRng Is Nothing Then
            If Not Intersect(cell, namedRng) Is Nothing Then
                clsState.CellCtxMenuType = CCM_Hotspot
                GoTo FoundMatch
            End If
        End If
        Set namedRng = Nothing
        
        ' === Check Flag ID ===
        
        On Error Resume Next
        Set namedRng = sheet.Names(NAME_RANGE_FLAGS_FLAG_ID).RefersToRange
        On Error GoTo ErrHandler
        If Not namedRng Is Nothing Then
            If Not Intersect(cell, namedRng) Is Nothing Then
                clsState.CellCtxMenuType = CCM_Flags
                GoTo FoundMatch
            End If
        End If
        Set namedRng = Nothing
                
FoundMatch:
    End If
    
    EvaluateCellCtxMenu = clsState.CellCtxMenuType
    
    If clsState.CellCtxMenuType = CCM_Default Then
        ' Update cache for default case
        If Not m_isCtxCacheInitialized Then BuildCellCtxMenuCache
        
        If Application.CommandBars("Cell").Controls.count <> m_controlCountSignature Then
            m_isCtxCacheInitialized = False
            BuildCellCtxMenuCache
            m_controlCountSignature = Application.CommandBars("Cell").Controls.count
        End If
        ShowAllCachedCellCtx
        clsState.CellCtxMnuNeedsPrepare = False
    Else
        clsState.CellCtxMnuHideDefault = True
    End If
    
CleanExit:
    Exit Function
ErrHandler:
    modErr.ReportError "EvaluateCellCtxMenu", Err.Number, Erl, caption:=modMain.AppProjectName
    Resume CleanExit
End Function

' -----------------------------------------------------------------------------------
' Procedure : ResetToDefaultCtxMenu
' Purpose   : Short-circuit method to restore default context menu state
' Parameters: (none)
' Returns   : (none)
' Notes     : Only resets if current menu type is not already default
' -----------------------------------------------------------------------------------
Public Sub ResetToDefaultCtxMenu()
    On Error GoTo ErrHandler
    
    ' Skip if already in default state AND menu wasn't manually hidden
    If clsState.CellCtxMenuType = CCM_Default And Not clsState.CellCtxMnuHideDefault Then Exit Sub
    
    ' Reset any custom context menu modifications
    If Not m_isCtxCacheInitialized Then BuildCellCtxMenuCache
    
    If Application.CommandBars("Cell").Controls.count <> m_controlCountSignature Then
        m_isCtxCacheInitialized = False
        BuildCellCtxMenuCache
        m_controlCountSignature = Application.CommandBars("Cell").Controls.count
    End If
    
    ShowAllCachedCellCtx
    
    ' Reset menu type
    clsState.CellCtxMenuType = CCM_Default
    clsState.CellCtxMnuNeedsPrepare = False
    clsState.CellCtxMnuHideDefault = False
        
    ' Clear dynamic button states
    clsState.InvalidateControl RIBBON_CTX_MNU_BTN_1
    clsState.InvalidateControl RIBBON_CTX_MNU_BTN_2
            
CleanExit:
    Exit Sub
ErrHandler:
    modErr.ReportError "ResetToDefaultCtxMenu", Err.Number, Erl
    On Error Resume Next
    ' Hard reset for ctx menu
    CommandBars("Cell").Reset
    
    ' Cache invalidieren nach Reset!
    m_isCtxCacheInitialized = False
    m_controlCountSignature = 0
    
    ' State zurücksetzen
    clsState.CellCtxMenuType = CCM_Default
    clsState.CellCtxMnuNeedsPrepare = False
    clsState.CellCtxMnuHideDefault = False
    On Error GoTo 0
End Sub

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

    Dim idx As Long, strPrt As String
    If (Not Not m_cellCtxControls) = 0 Then Exit Sub
    On Error Resume Next
    For idx = 1 To UBound(m_cellCtxControls)
        If InStr(1, m_cellCtxCaptions(idx), part, vbTextCompare) > 0 Then m_cellCtxControls(idx).Visible = True
    Next
    
CleanExit:
    Exit Sub
ErrHandler:
    modErr.ReportError "ShowCellCtxByCachedCaption", Err.Number, Erl, caption:=modMain.AppProjectName
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
    Dim idx As Long
    On Error Resume Next
    For idx = 1 To UBound(m_cellCtxControls)
        m_cellCtxControls(idx).Visible = True
    Next
End Sub

' ===== Private Helpers =============================================================
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
    Dim cellCommandBar As CommandBar, idx As Long
    If Not m_isCtxCacheInitialized Then
        Set cellCommandBar = Application.CommandBars("Cell")

        ReDim m_cellCtxControls(1 To cellCommandBar.Controls.count)
        ReDim m_cellCtxCaptions(1 To cellCommandBar.Controls.count)

        For idx = 1 To cellCommandBar.Controls.count
            Set m_cellCtxControls(idx) = cellCommandBar.Controls(idx)
            'If cellCommandBar.Controls(Idx).BuiltIn Then Debug.Print cellCommandBar.Controls(Idx).Caption
            m_cellCtxCaptions(idx) = m_cellCtxControls(idx).caption
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
    Dim idx As Long
    On Error Resume Next
    For idx = 1 To UBound(m_cellCtxControls)
        If m_cellCtxControls(idx).BuiltIn Then m_cellCtxControls(idx).Visible = False
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
        If InStr(1, CmdBarCtrl.caption, part, vbTextCompare) > 0 Then CmdBarCtrl.Visible = False
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
        If InStr(1, CmdBarCtrl.caption, part, vbTextCompare) > 0 Then CmdBarCtrl.Visible = True
    Next CmdBarCtrl
End Sub

' -----------------------------------------------------------------------------------
' Procedure : CleanupCellCtxMenu
' Purpose   : Cleans up module state, releases object references, and restores the
'             default "Cell" context menu.
'
' Params    : (none)
' Returns   : (none)
'
' Behavior  :
'   - Resets the "Cell" CommandBar to default Excel state.
'   - Releases all cached CommandBarControl object references.
'   - Clears all module-level arrays and flags.
'   - Resets related clsState flags to default values.
'
' Notes:
'   - Should be called during add-in shutdown or before re-initialization.
'   - Uses On Error Resume Next to ensure cleanup completes even if errors occur.
' -----------------------------------------------------------------------------------
Public Sub CleanupCellCtxMenu()
    On Error Resume Next
    
    Dim idx As Long
    
    ' Release all cached CommandBarControl object references
    If (Not Not m_cellCtxControls) <> 0 Then
        For idx = LBound(m_cellCtxControls) To UBound(m_cellCtxControls)
            Set m_cellCtxControls(idx) = Nothing
        Next idx
    End If
    
    ' Clear arrays
    Erase m_cellCtxControls
    Erase m_cellCtxCaptions
    m_ctxMenuWhitelist = Empty
    
    ' Reset module-level state variables
    m_isCtxCacheInitialized = False
    m_controlCountSignature = 0
    
    ' Reset the "Cell" CommandBar to Excel default
    Application.CommandBars("Cell").Reset
    
    ' Reset global state flags in clsState
    clsState.CellCtxMenuType = CCM_Default
    clsState.CellCtxMnuNeedsPrepare = False
    clsState.CellCtxMnuHideDefault = False
    
    ' Clear dynamic button states
    clsState.InvalidateControl RIBBON_CTX_MNU_BTN_1
    clsState.InvalidateControl RIBBON_CTX_MNU_BTN_2
    
    On Error GoTo 0
End Sub

