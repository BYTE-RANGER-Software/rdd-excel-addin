VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFormDropManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' -----------------------------------------------------------------------------------
' Class     : clsFormDropManager
' Purpose   : Business logic for "wandering" Form DropDown pair (arrow + 2 dropdowns).
'             One clsFormDrop manager per worksheet, reused and moved on selection.
'
' Public API:
'   - Init                         : Initialize manager with callback host.
'   - HandleSelectionChange        : Central selection handler (shows arrow / hides).
'   - HandleSheetDeactivate        : Hide on sheet switch.
'
' Dependencies:
'   - clsFormDrop (main class)
'   - modFormDropRouter (routes Form control events)
'   - modTags, modRanges, modErr
'
' Notes     :
'   - Requires reference "Microsoft Scripting Runtime" (Scripting.Dictionary).
'   - This module owns the per-worksheet clsFormDrop instances.
'   - Uses Application.Run to invoke callbacks on the provided callback host.
' -----------------------------------------------------------------------------------
Option Explicit

' === Private state =================================================================

Private m_isInitialized As Boolean

' Internal registry: one clsFormDrop manager per worksheet (keyed by CodeName).
Private m_mngrMap As Scripting.Dictionary     ' key = Worksheet.CodeName, value = clsFormDrop
Private m_formDropRegistry As Scripting.Dictionary 'clsFormDrop instance registry: instance id -> class instance (clsFormDrop).

Private m_onCatCallback As String  ' Public Module that holds callback method for category
Private m_onSubCallback As String  ' Public Module that holds callback method for subitems

' Anchor config cache: key = anchor Name (Name.Name), value = config dictionary
' with CatRange / CatLabels / SubRanges.
Private m_anchorConfigCache As Scripting.Dictionary

' === Public API ====================================================================


' -----------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Purpose   : Handles initialization of a new class instance.
'
' Parameters:
'   (none)
'
' Returns   : (none)
'
' Notes     :
'   - Actual logical initialization is performed in Init.
'   - keep empty for default instance
' -----------------------------------------------------------------------------------
Private Sub Class_Initialize()
    ' Leave empty, use Init() for initialization.
End Sub

' -----------------------------------------------------------------------------------
' Procedure : Init
' Purpose   : Initialize feature: create instance map.
'
' Parameters:
'   onCatCallback  [String] - Optional; Full callback names (Module.Procedure) for category
'   onSubCallback  [String] - Optional; Full callback names (Module.Procedure) for subitems
'
' Returns   : (none)
'
' Notes     : Safe to call multiple times.
' -----------------------------------------------------------------------------------
Public Sub Init(Optional ByVal onCatCallback As String = vbNullString, Optional ByVal onSubCallback As String = vbNullString)
                
    On Error GoTo ErrHandler
    
    If Not m_isInitialized Then
        If m_mngrMap Is Nothing Then Set m_mngrMap = New Scripting.Dictionary
        
        m_onCatCallback = onCatCallback
        m_onSubCallback = onSubCallback
        
        If m_formDropRegistry Is Nothing Then
            Set m_formDropRegistry = New Scripting.Dictionary
        End If
        
        ' give router access to instance registry (reference)
        Set modFormDropRouter.g_formDropRegistryDict = m_formDropRegistry
                
        If m_anchorConfigCache Is Nothing Then
            Set m_anchorConfigCache = New Scripting.Dictionary
            m_anchorConfigCache.CompareMode = TextCompare
        End If
        
        m_isInitialized = True
    End If
    
CleanExit:
    Exit Sub
ErrHandler:
    modErr.ReportError "clsFormDropManager.Init", Err.Number, Erl
    Resume CleanExit
End Sub

' -----------------------------------------------------------------------------------
' Procedure : Class_Terminate
' Purpose   : Handles cleanup for a class instance when it is destroyed;
'             destroy all managers and clear map.
'
' Parameters:
'   (none)
'
' Returns   : (none)
'
' Notes     :
'   - Safeguards against missing registry or missing key.
' -----------------------------------------------------------------------------------
Private Sub Class_Terminate()
    On Error GoTo ErrHandler
        
    If m_isInitialized Then
        Dim key As Variant
        If Not m_mngrMap Is Nothing Then
            For Each key In m_mngrMap.Keys
                On Error Resume Next
                Dim FormDropInst As clsFormDrop
                Set FormDropInst = m_mngrMap(key)
                    
                If Not FormDropInst Is Nothing Then
                    FormDropInst.Dispose
                    Set FormDropInst = Nothing
                End If
                
                On Error GoTo ErrHandler
            Next key
            m_mngrMap.RemoveAll
            Set m_mngrMap = Nothing
        End If
        
        ' clean up instance registry
        If Not m_formDropRegistry Is Nothing Then
            m_formDropRegistry.RemoveAll
            Set m_formDropRegistry = Nothing
        End If
        
        ' remove router access
        Set modFormDropRouter.g_formDropRegistryDict = Nothing
                
        m_onCatCallback = vbNullString
        m_onSubCallback = vbNullString
        
        If Not m_anchorConfigCache Is Nothing Then
            m_anchorConfigCache.RemoveAll
            Set m_anchorConfigCache = Nothing
        End If
        
        m_isInitialized = False
    End If
    
CleanExit:
    Exit Sub
ErrHandler:
    modErr.ReportError "clsFormDropManager.Class_Terminate", Err.Number, Erl
    Resume CleanExit
End Sub

' -----------------------------------------------------------------------------------
' Procedure : HandleSelectionChange
' Purpose   : Form Drop handler for App.SheetSelectionChange.
'             Shows arrow only when the selected cell is an anchor (named *!DD_Anchor_*).;
'             hides otherwise. Dropdowns open on arrow click.
'
' Parameters:
'   eventSheet       [Worksheet] - Worksheet raising the event.
'   targetRange      [Range]     - New selection.
'
' Returns   : (none)
'
' Notes     :
'   - Requires that lists were set in clsFormDrop; otherwise ShowAt will raise an error.
' -----------------------------------------------------------------------------------
Public Sub HandleSelectionChange(ByVal eventSheet As Worksheet, ByVal targetRange As Range)
    If Not EnsureInitialized Then Exit Sub
    
    On Error GoTo ErrHandler
    
    Dim isSlaveAnchor As Boolean
    Dim targetCell As Range
    Dim cfgCell As Range
    
    If Not ShouldHandleSheet(eventSheet) Then Exit Sub

    Dim formDropMngr As clsFormDrop: Set formDropMngr = GetMngr(eventSheet)

    If targetRange Is Nothing Or targetRange.CountLarge > 1 Then
        formDropMngr.HideDropDowns True
        Exit Sub
    End If
    
    Set targetCell = targetRange.Cells(1, 1)
     
    If IsAnchorCell(eventSheet, targetCell, isSlaveAnchor) Then
        If isSlaveAnchor Then
            Set cfgCell = GetMasterAnchorCell(eventSheet, targetCell)
            If cfgCell Is Nothing Then
                formDropMngr.HideDropDowns True
                GoTo CleanExit
            End If
        Else
            Set cfgCell = targetCell
        End If
        
        Call ConfigListsFromAnchorMeta(eventSheet, cfgCell, formDropMngr)
        formDropMngr.ShowAt eventSheet, targetCell
    Else
        formDropMngr.HideDropDowns True
    End If
    
CleanExit:
    Exit Sub
ErrHandler:
    modErr.ReportError "clsFormDropManager.HandleSelectionChange", Err.Number, Erl
    Resume CleanExit
End Sub

' -----------------------------------------------------------------------------------
' Procedure : HandleSheetDeactivate
' Purpose   : Hide controls when leaving a sheet.
'             Form Drop Handler for App_SheetDeactivate
'
' Parameters:
'   eventSheet       [Worksheet] - Worksheet being deactivated.
'
' Returns   : (none)
' -----------------------------------------------------------------------------------
Public Sub HandleSheetDeactivate(ByVal eventSheet As Worksheet)
    If Not EnsureInitialized Then Exit Sub
    
    On Error GoTo ErrHandler
    
    If m_mngrMap Is Nothing Then Exit Sub
    
    Dim key As String: key = eventSheet.CodeName
    Dim fd As clsFormDrop
    If m_mngrMap.Exists(key) Then
        Set fd = m_mngrMap(key)
        fd.HideDropDowns
    End If
    
CleanExit:
    Exit Sub
ErrHandler:
    modErr.ReportError "clsFormDropManager.HandleSheetDeactivate", Err.Number, Erl
    Resume CleanExit
End Sub

' -----------------------------------------------------------------------------------
' Procedure : ClearAnchorCache
' Purpose   : Clears all cached anchor configurations. Use when anchor metadata
'             (Name.Comment) has been changed and should be re-parsed.
'
' Parameters:
'   (none)
'
' Returns   : (none)
' -----------------------------------------------------------------------------------
Public Sub ClearAnchorCache()
    If m_anchorConfigCache Is Nothing Then Exit Sub
    m_anchorConfigCache.RemoveAll
End Sub

' -----------------------------------------------------------------------------------
' Procedure : PreFillAnchorCache
' Purpose   : Pre-populates the anchor metadata cache for the given set of
'             worksheets by scanning all Names in the owning workbook and
'             resolving configs for matching anchors.
'
' Parameters:
'   targetSheets   [Variant] - 1-D array of Worksheet objects.
'
' Returns   : (none)
'
' Notes     :
'   - Safe to call multiple times; existing cache entries are reused.
'   - Only Names whose RefersToRange is on one of the provided sheets and whose
'     Name matches FD_ANCHOR_NAME_PATTERN are processed.
' -----------------------------------------------------------------------------------
Public Sub PreFillAnchorCache(ByVal targetSheets As Variant)
    If Not EnsureInitialized Then Exit Sub
        
    On Error GoTo ErrHandler
    
        If m_anchorConfigCache Is Nothing Then
        Set m_anchorConfigCache = New Scripting.Dictionary
        m_anchorConfigCache.CompareMode = TextCompare
    End If
    
    If Not IsArray(targetSheets) Then Exit Sub
            
    Dim i As Long
    For i = LBound(targetSheets) To UBound(targetSheets)
        Dim hostSheet As Worksheet
        Set hostSheet = targetSheets(i)
        
        Dim anchorName As Name
        For Each anchorName In hostSheet.Names
            Dim refCell As Range
            On Error Resume Next
            Set refCell = anchorName.RefersToRange
            On Error GoTo ErrHandler
            
            If Not refCell Is Nothing Then
                ' Only process names that match the anchor pattern
                If anchorName.Name Like FD_ANCHOR_NAME_PATTERN Then
                    Call GetOrCreateAnchorConfig(anchorName)
                End If
            End If
            
            Set refCell = Nothing
        Next anchorName
    Next i
    
CleanExit:
    Exit Sub
ErrHandler:
    modErr.ReportError "clsFormDropManager.PreFillAnchorCache", Err.Number, Erl
    Resume CleanExit
End Sub



' === Private helpers ===============================================================

' -----------------------------------------------------------------------------------
' Function  : ShouldHandleSheet
' Purpose   : Determine whether the given sheet participates in the feature.
'
' Parameters:
'   hostSheet       [Worksheet] - Candidate sheet.
'
' Returns   : [Boolean] - True if the sheet should be handled.
'
' Notes     :
'       - Skips when dispatcher tag is not present.
'       - Skips dispatcher sheet itself and sheets tagged with SHEET_DISPATCHER.
' -----------------------------------------------------------------------------------
Private Function ShouldHandleSheet(ByVal hostSheet As Worksheet) As Boolean
    On Error GoTo ErrHandler
    
    Dim hostWorkbook As Workbook: Set hostWorkbook = hostSheet.Parent
    
    If Not modTags.SheetWithTagExists(hostWorkbook, SHEET_DISPATCHER) Then Exit Function
    If hostSheet.CodeName = SHEET_DISPATCHER Then Exit Function
    If modTags.HasSheetTag(hostSheet, SHEET_DISPATCHER) Then Exit Function
    
    ShouldHandleSheet = True
CleanExit:
    Exit Function
ErrHandler:
    ShouldHandleSheet = False
    Resume CleanExit
End Function

' -----------------------------------------------------------------------------------
' Function  : GetMngr
' Purpose   : Get the per-sheet clsFormDrop manager.
'
' Parameters:
'   hostSheet       [Worksheet] - Host sheet.
'
' Returns   : [clsFormDrop] - Configured manager instance.
'
' Notes     :
'   - Sets arrow options and list sources once on first use.
' -----------------------------------------------------------------------------------

Private Function GetMngr(ByVal hostSheet As Worksheet) As clsFormDrop
    On Error GoTo ErrHandler
    
    If m_mngrMap Is Nothing Then Err.Raise ERR_OBJ_IS_NOTHING, "clsFormDropManager.GetMngr", _
                  "FormDrop Manager Map is nothing / not initialized"

    Dim key As String: key = hostSheet.CodeName
    
    If Not m_mngrMap.Exists(key) Then
        Dim formDrop As New clsFormDrop
        formDrop.Init hostSheet.Parent, m_formDropRegistry
        
        ' Arrow config
        formDrop.SetArrowEnabled True
        formDrop.SetArrowStyle 2, 10, 10
        formDrop.SetPlacement True
                
        'Set callbacks for drop-downs
        If Len(m_onCatCallback) Or Len(m_onSubCallback) Then
            formDrop.SetCallbacks m_onCatCallback, m_onSubCallback
        End If
                              
        m_mngrMap.Add key, formDrop
    End If

    Set GetMngr = m_mngrMap(key)
    
CleanExit:
    Exit Function
ErrHandler:
    modErr.ReportError "clsFormDropManager.GetMngr", Err.Number, Erl
    Resume CleanExit
End Function

' -----------------------------------------------------------------------------------
' Function  : IsAnchorCell
' Purpose   : True if the cell has a name matching FD_ANCHOR_NAME_PATTERN.
'
' Parameters:
'   hostSheet    [Worksheet] - Host sheet.
'   Cell         [Range]     - Single cell to test.
'   isSlave      [Boolean]   - Optionl, ByRef: True if the anchor is a slave anchor (with index
'                              suffix like ".1."); False if master or not an anchor.
'
'
' Returns   : [Boolean]      - True if the cell is an anchor cell; otherwise False.
' -----------------------------------------------------------------------------------
Private Function IsAnchorCell(ByVal hostSheet As Worksheet, ByVal cell As Range, Optional ByRef isSlave As Boolean = False) As Boolean
    Dim nm As Name
    
    isSlave = False
    
    Set nm = modRanges.GetCellNameByPattern(cell, FD_ANCHOR_NAME_PATTERN)
    If nm Is Nothing Then
        IsAnchorCell = False
        Exit Function
    End If
    
    IsAnchorCell = True
    
    If nm.Name Like FD_SLAVE_ANCHOR_NAME_PATTERN Then
        isSlave = True
    End If
   
    
End Function

' -----------------------------------------------------------------------------------
' Function  : GetMasterAnchorCell
' Purpose   : Returns the master anchor cell for a given anchor cell (master or slave).
'
' Parameters:
'   hostSheet    [Worksheet] - Host sheet.
'   cell         [Range]     - Anchor cell (master or slave).
'
' Returns   : [Range]        - The master anchor cell, or Nothing if no anchor found.
' -----------------------------------------------------------------------------------
Private Function GetMasterAnchorCell( _
        ByVal hostSheet As Worksheet, _
        ByVal cell As Range) As Range
    
    Dim nm As Name
    Dim masterNm As Name
    Dim fullName As String
    Dim masterName As String
    Dim lastDot As Long
    Dim prevDot As Long
    
    ' Get the defined name for this cell
    Set nm = modRanges.GetCellNameByPattern(cell, FD_ANCHOR_NAME_PATTERN)
    If nm Is Nothing Then
        Exit Function
    End If
    
    fullName = nm.Name                     ' e.g. "R001!DD_Anchor_Target.1."
    
    ' Is it a slave anchor (".<index>." at the end)?
    If fullName Like FD_SLAVE_ANCHOR_NAME_PATTERN Then
        lastDot = InStrRev(fullName, ".")          ' trailing "."
        If lastDot > 0 Then
            prevDot = InStrRev(fullName, ".", lastDot - 1)
            If prevDot > 0 Then
                ' Cut off ".<index>."
                masterName = Left$(fullName, prevDot - 1)
            Else
                masterName = fullName
            End If
        Else
            masterName = fullName
        End If
    Else
        ' Already master
        masterName = fullName
    End If
    
    ' If the name did not change, we are already on the master anchor
    If StrComp(masterName, nm.Name, vbTextCompare) = 0 Then
        Set GetMasterAnchorCell = cell
        Exit Function
    End If
    
    ' Resolve the master name to a range via the workbook's Names collection
    On Error Resume Next
    Set masterNm = nm.Parent.Names(masterName)
    On Error GoTo 0
    
    If Not masterNm Is Nothing Then
        Set GetMasterAnchorCell = masterNm.RefersToRange
    End If
End Function


' -----------------------------------------------------------------------------------
' Function  : GetOrCreateAnchorConfig
' Purpose   : Returns a cached anchor configuration for the given Name or builds
'             and caches it from the Name.Comment metadata.
'
' Parameters:
'   anchorName     [Name]     - Name referring to the anchor cell.
'
' Returns   : Scripting.Dictionary - Config dictionary with keys:
'                 "CatRange"  [Range or Nothing]
'                 "CatLabels" [Variant array of String or Empty]
'                 "SubRanges" [Variant array of Range]
'             or Nothing if metadata is invalid or missing.
'
' Notes     :
'   - Config is built once per anchorName.Name and reused on subsequent calls.
'   - Only two result shapes are supported:
'       * CatRange + SubRanges
'       * CatLabels + SubRanges
'   - Supports two variants of metadata:
'       * Variant A: Worksheet Names (cat, subs)
'       * Variant B: ListObject tables (cattbl, catcol, substbl, subscols)
'         with optional sheet names (catsheet, subsheet) to specify where
'         the tables are located (defaults to anchor sheet if not specified)
' -----------------------------------------------------------------------------------
Private Function GetOrCreateAnchorConfig(ByVal anchorName As Name) As Scripting.Dictionary
    
    On Error GoTo ErrHandler
    
    If m_anchorConfigCache Is Nothing Then
        Set m_anchorConfigCache = New Scripting.Dictionary
        m_anchorConfigCache.CompareMode = TextCompare
    End If
    
    Dim cacheKey As String
    cacheKey = anchorName.Name
    
    ' Use cached config when available
    On Error Resume Next
    Set GetOrCreateAnchorConfig = m_anchorConfigCache(cacheKey)
    On Error GoTo 0

    If Not GetOrCreateAnchorConfig Is Nothing Then Exit Function
    
    ' --- Build configuration from metadata ---
    Dim meta As String
    On Error Resume Next
    meta = CStr(anchorName.Comment)
    On Error GoTo ErrHandler
    
    If Len(meta) = 0 Then Exit Function
    If StrComp(Left$(Trim$(meta), Len(FD_META_PREFIX)), FD_META_PREFIX, vbTextCompare) <> 0 Then Exit Function
    
    meta = Trim$(Mid$(meta, Len(FD_META_PREFIX) + 1))
    
    Dim kv As Scripting.Dictionary
    Set kv = ParseMetaKeyValues(meta)
    If kv Is Nothing Then Exit Function
    
    Dim anchorSheet As Worksheet
    Set anchorSheet = anchorName.RefersToRange.Parent
    
    Dim cfg As Scripting.Dictionary
    Set cfg = New Scripting.Dictionary
    cfg.CompareMode = TextCompare
        
    ' --- Variant A: worksheet Names ---
    If kv.Exists(FD_META_KEY_SUBS) Then
        Dim subsArr As Variant
        subsArr = modUtil.SplitTrim(CStr(kv(FD_META_KEY_SUBS)), FD_META_LIST_SEP)
        
        Dim i As Long
        Dim subRanges() As Variant
        ReDim subRanges(LBound(subsArr) To UBound(subsArr))
        
        For i = LBound(subsArr) To UBound(subsArr)
            Dim nr As Range
            On Error Resume Next
            Set nr = anchorSheet.Names(CStr(subsArr(i))).RefersToRange
            On Error GoTo ErrHandler
            If nr Is Nothing Then Exit Function
            Set subRanges(i) = nr
        Next i
        
        cfg("SubRanges") = subRanges
        
        If kv.Exists(FD_META_KEY_CAT) Then
            Dim catRng As Range
            On Error Resume Next
            Set catRng = anchorSheet.Names(CStr(kv(FD_META_KEY_CAT))).RefersToRange
            On Error GoTo ErrHandler
            If catRng Is Nothing Then Exit Function
            Set cfg("CatRange") = catRng
        Else
            Set cfg("CatRange") = Nothing
            cfg("CatLabels") = subsArr
        End If
        
        Set m_anchorConfigCache(cacheKey) = cfg
        Set GetOrCreateAnchorConfig = cfg
        Exit Function
    End If
    
    ' --- Variant B: table & columns ---
    If kv.Exists(FD_META_KEY_SUBS_TBL) And kv.Exists(FD_META_KEY_SUBS_COLS) Then
        
        ' Determine the sheet where substbl is located
        Dim subsSheet As Worksheet
        If kv.Exists(FD_META_KEY_SUBS_SHEET) Then
            ' Optional: use specified sheet for substbl
            On Error Resume Next
            Set subsSheet = anchorSheet.Parent.Worksheets(CStr(kv(FD_META_KEY_SUBS_SHEET)))
            On Error GoTo ErrHandler
            If subsSheet Is Nothing Then Exit Function
        Else
            ' Default: use anchor sheet
            Set subsSheet = anchorSheet
        End If
        
        Dim loSubs As ListObject
        Set loSubs = modRanges.GetTable(subsSheet, CStr(kv(FD_META_KEY_SUBS_TBL)))
        If loSubs Is Nothing Then Exit Function
        
        Dim cols As Variant
        cols = modUtil.SplitTrim(CStr(kv(FD_META_KEY_SUBS_COLS)), FD_META_LIST_SEP)
        
        Dim subRanges2() As Variant
        ReDim subRanges2(LBound(cols) To UBound(cols))
        
        Dim i2 As Long
        For i2 = LBound(cols) To UBound(cols)
            Dim rr As Range
            On Error Resume Next
            Set rr = loSubs.ListColumns(CStr(cols(i2))).DataBodyRange
            On Error GoTo ErrHandler
            If rr Is Nothing Then Exit Function
            Set subRanges2(i2) = rr
        Next i2
        
        cfg("SubRanges") = subRanges2
        
        If kv.Exists(FD_META_KEY_CAT_TBL) And kv.Exists(FD_META_KEY_CAT_COL) Then
        
            ' Determine the sheet where cattbl is located
            Dim catSheet As Worksheet
            If kv.Exists(FD_META_KEY_CAT_SHEET) Then
                ' Optional: use specified sheet for cattbl
                On Error Resume Next
                Set catSheet = anchorSheet.Parent.Worksheets(CStr(kv(FD_META_KEY_CAT_SHEET)))
                On Error GoTo ErrHandler
                If catSheet Is Nothing Then Exit Function
            Else
                ' Default: use anchor sheet
                Set catSheet = anchorSheet
            End If
            
            Dim loCat As ListObject
            Dim catRng2 As Range
            Set loCat = modRanges.GetTable(subsSheet, CStr(kv(FD_META_KEY_CAT_TBL)))
            If loCat Is Nothing Then Exit Function
            On Error Resume Next
            Set catRng2 = loCat.ListColumns(CStr(kv(FD_META_KEY_CAT_COL))).DataBodyRange
            On Error GoTo ErrHandler
            If catRng2 Is Nothing Then Exit Function
            Set cfg("CatRange") = catRng2
        Else
            Set cfg("CatRange") = Nothing
            cfg("CatLabels") = cols
        End If
        
        Set m_anchorConfigCache(cacheKey) = cfg
        Set GetOrCreateAnchorConfig = cfg
        Exit Function
    End If
    
CleanExit:
    Exit Function
ErrHandler:
    modErr.ReportError "clsFormDropManager.GetOrCreateAnchorConfig", Err.Number, Erl
    Resume CleanExit
End Function

' -----------------------------------------------------------------------------------
' Function  : ConfigListsFromAnchorMeta
' Purpose   : Tries to configure category/sub lists for an anchor anchorCell based on the
'             Name.Comment metadata.
'
' Parameters:
'   hostSheet    [Worksheet]  - Host sheet for lookup.
'   anchorCell   [Range]      - Anchor anchorCell carrying a Name.
'   formDrop     [clsFormDrop]- Form Drop Manager to configure.
'
' Returns   : Boolean - True if configuration was applied; False if no/invalid metadata.
'
' Notes     :
'   - Supports two variants:
'       * Variant A: workbook Names (cat + subs) OR subs only as category labels.
'       * Variant B: ListObject (table & columns).
' -----------------------------------------------------------------------------------
Private Function ConfigListsFromAnchorMeta( _
    ByVal hostSheet As Worksheet, _
    ByVal anchorCell As Range, _
    ByVal formDrop As clsFormDrop) As Boolean
    
    Dim nm As Name
    Set nm = modRanges.GetCellNameByPattern(anchorCell, FD_ANCHOR_NAME_PATTERN)
    If nm Is Nothing Then Exit Function
        
    Dim cfg As Scripting.Dictionary
    Set cfg = GetOrCreateAnchorConfig(nm)
    If cfg Is Nothing Then Exit Function
    
    Dim catRng As Range
    Dim catLabels As Variant
    Dim subRanges As Variant
    
    If cfg.Exists("CatRange") Then
        Set catRng = cfg("CatRange")
    End If
    If cfg.Exists("CatLabels") Then
        catLabels = cfg("CatLabels")
    End If
    subRanges = cfg("SubRanges")
    
    If Not catRng Is Nothing Then
        formDrop.SetListsFromNamedRanges catRng, subRanges
    Else
        formDrop.SetListsFromLabelsAndRanges catLabels, subRanges
    End If
    
    ConfigListsFromAnchorMeta = True
End Function


' -----------------------------------------------------------------------------------
' Function  : ParseMetaKeyValues
' Purpose   : Parses "key=value; key=value" style metadata into a late-bound dictionary
'             with case-insensitive keys.
'
' Parameters:
'   metaString   [String] - Raw metadata string.
'
' Returns   : Object - Scripting.Dictionary containing parsed key/value pairs.
'
' Notes     :
'   - Uses FD_META_PAIR_SEP for splitting pairs.
' -----------------------------------------------------------------------------------
Private Function ParseMetaKeyValues(ByVal metaString As String) As Scripting.Dictionary
    Dim metaDict As Scripting.Dictionary: Set metaDict = New Scripting.Dictionary
    metaDict.CompareMode = 1
    Dim parts As Variant, i As Long
    parts = Split(metaString, FD_META_PAIR_SEP)
    For i = LBound(parts) To UBound(parts)
        Dim part As String: part = Trim$(parts(i))
        If Len(part) > 0 Then
            Dim eqPos As Long: eqPos = InStr(1, part, "=", vbTextCompare)
            If eqPos > 1 Then
                Dim key As String, Value As String
                key = Trim$(Left$(part, eqPos - 1))
                Value = Trim$(Mid$(part, eqPos + 1))
                If Len(key) > 0 Then metaDict(key) = Value
            End If
        End If
    Next
    Set ParseMetaKeyValues = metaDict
End Function

' -----------------------------------------------------------------------------------
' Function  : EnsureInitialized
' Purpose   : Ensures that Init has been called before using the instance.
'
' Parameters:
'   (none)
'
' Returns   : Boolean - True when initialized; raises error otherwise.
'
' Notes     :
'   - Raises ERR_NOT_INITIALIZED when the instance is not initialized.
' -----------------------------------------------------------------------------------
Private Function EnsureInitialized() As Boolean
    EnsureInitialized = m_isInitialized
    If Not EnsureInitialized Then
        Err.Raise ERR_NOT_INITIALIZED, "clsFormDropManager", "Call .Init first."
    End If
End Function
