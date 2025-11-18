Attribute VB_Name = "modFormDropManager"
' -----------------------------------------------------------------------------------
' Module    : modFormDropManager
' Purpose   : Business logic for "wandering" Form DropDown pair (arrow + 2 dropdowns).
'             One clsFormDrop manager per worksheet, reused and moved on selection.
'
' Public API:
'   - FD_InitMngrs()                         : Prepare feature (per-app lifetime).
'   - FD_DisposeMngrs()                      : Destroy all per-sheet managers.
'   - FD_HandleSelectionChange(Sh, Target)   : Central selection handler (shows arrow / hides).
'   - FD_HandleSheetDeactivate(Sh)           : Hide on sheet switch.
'
' Dependencies:
'   - clsFormDrop (main class)
'   - modFormDropRouter (routes Form control events)
'   - modTags
'   - modRanges
'   - modErr
'
' Notes     :
'   - Requires reference "Microsoft Scripting Runtime" (Scripting.Dictionary).
'   - This module owns the per-worksheet clsFormDrop instances.
' -----------------------------------------------------------------------------------
Option Explicit
Option Private Module

' ===== Private State ================================================================
' Internal registry: one clsFormDrop manager per worksheet (keyed by CodeName).

Private m_mngrMap As Scripting.Dictionary   ' key = Worksheet.CodeName, value = clsFormDrop

' ===== Public API ===================================================================
' Public entry points used by the application event layer (clsAppEvents).

' -----------------------------------------------------------------------------------
' Procedure : FD_InitMngrs
' Purpose   : Initialize feature: create instance map.
'
' Parameters:
'   (none)
'
' Returns   : (none)
'
' Notes     : Safe to call multiple times.
' -----------------------------------------------------------------------------------
Public Sub FD_InitMngrs()
On Error GoTo ErrHandler
    If m_mngrMap Is Nothing Then Set m_mngrMap = New Scripting.Dictionary
CleanExit:
    Exit Sub
ErrHandler:
    modErr.ReportError "FD_InitMngrs", Err.Number, Erl
    Resume CleanExit
End Sub

' -----------------------------------------------------------------------------------
' Procedure : FD_DisposeMngrs
' Purpose   : Destroy all managers and clear map.
'
' Parameters:
'   (none)
'
' Returns   : (none)
'
' Notes     : Safe to call multiple times.
' -----------------------------------------------------------------------------------
Public Sub FD_DisposeMngrs()
On Error GoTo ErrHandler
    Dim key As Variant
    If Not m_mngrMap Is Nothing Then
        For Each key In m_mngrMap.Keys
            On Error Resume Next
            m_mngrMap(key).Destroy
            On Error GoTo ErrHandler
        Next key
        m_mngrMap.RemoveAll
    End If
CleanExit:
    Exit Sub
ErrHandler:
    modErr.ReportError "FD_DisposeMngrs", Err.Number, Erl
    Resume CleanExit
End Sub

' -----------------------------------------------------------------------------------
' Procedure : FD_HandleSelectionChange
' Purpose   : Form Drop handler for App.SheetSelectionChange.
'             Shows arrow only when the selected cell is an anchor (named DD_Anchor_*).; hides otherwise. Dropdowns open on arrow click.
'
' Parameters:
'   eventSheet       [Object] - Worksheet raising the event.
'   targetRange      [Range]  - New selection.
'
' Returns   : (none)
'
' Notes     :
'   - Requires that lists were set in clsFormDrop; otherwise ShowAt raises (as desired).
' -----------------------------------------------------------------------------------
Public Sub FD_HandleSelectionChange(ByVal eventSheet As Worksheet, ByVal targetRange As Range)
On Error GoTo ErrHandler

    If Not FD_ShouldHandleSheet(eventSheet) Then Exit Sub

    Dim mgr As clsFormDrop: Set mgr = FD_GetOrCreateMngr(eventSheet)

    If targetRange Is Nothing Or targetRange.CountLarge > 1 Then
        mgr.HideDropDowns
        Exit Sub
    End If

    If FD_IsAnchorCell(eventSheet, targetRange.Cells(1, 1)) Then
        Call FD_ConfigListsFromAnchorMeta(eventSheet, targetRange.Cells(1, 1), mgr)
        mgr.ShowAt eventSheet, targetRange.Cells(1, 1)   ' shows arrow only; dropdown opens on arrow click
    Else
        mgr.HideDropDowns
    End If

CleanExit:
    Exit Sub
ErrHandler:
    modErr.ReportError "FD_HandleSelectionChange", Err.Number, Erl
    Resume CleanExit
End Sub

' -----------------------------------------------------------------------------------
' Procedure : FD_HandleSheetDeactivate
' Purpose   : Hide controls when leaving a sheet. Form Drop Handler for App_SheetDeactivate
'
' Parameters:
'   eventSheet       [Object] - Worksheet being deactivated.
'
' Returns   : (none)
' -----------------------------------------------------------------------------------
Public Sub FD_HandleSheetDeactivate(ByVal eventSheet As Worksheet)
On Error GoTo ErrHandler
    If m_mngrMap Is Nothing Then Exit Sub
    Dim key As String: key = eventSheet.CodeName
    If m_mngrMap.Exists(key) Then m_mngrMap(key).Hide
CleanExit:
    Exit Sub
ErrHandler:
    modErr.ReportError "FD_HandleSheetDeactivate", Err.Number, Erl
    Resume CleanExit
End Sub

' ===== Private Helpers ==============================================================
' Internal helper functions to determine eligibility, create managers and
' configure list behavior based on anchor metadata.

' -----------------------------------------------------------------------------------
' Function  : FD_ShouldHandleSheet
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
Private Function FD_ShouldHandleSheet(ByVal hostSheet As Worksheet) As Boolean
On Error GoTo ErrHandler

    Dim hostWorkbook As Workbook
    Set hostWorkbook = hostSheet.Parent
    If Not modTags.SheetWithTagExists(hostWorkbook, SHEET_DISPATCHER) Then GoTo CleanExit
    If hostSheet.CodeName = SHEET_DISPATCHER Then GoTo CleanExit
    If modTags.HasSheetTag(hostSheet, SHEET_DISPATCHER) Then GoTo CleanExit

    FD_ShouldHandleSheet = True
CleanExit:
    Exit Function
ErrHandler:
    ' On any error: be conservative and do not handle
    FD_ShouldHandleSheet = False
    Resume CleanExit
End Function


' -----------------------------------------------------------------------------------
' Function  : FD_GetOrCreateMngr
' Purpose   : Get or create the per-sheet clsFormDrop manager.
'
' Parameters:
'   hostSheet       [Worksheet] - Host sheet.
'
' Returns   : [clsFormDrop] - Configured manager instance.
'
' Notes     :
'   - Sets arrow options and list sources once on first use.
' -----------------------------------------------------------------------------------
Private Function FD_GetOrCreateMngr(ByVal hostSheet As Worksheet) As clsFormDrop
On Error GoTo ErrHandler
    If m_mngrMap Is Nothing Then FD_InitMngrs

    Dim key As String: key = hostSheet.CodeName
    If Not m_mngrMap.Exists(key) Then
        Dim formDrop As New clsFormDrop
        formDrop.Init hostSheet.Parent, modFormDropRouter.g_formDropRegistryDict

        ' Arrow config
        formDrop.SetArrowEnabled True
        formDrop.SetArrowStyle 2, 10, 10
        formDrop.SetPlacement True

        ' Provide default list sources
        formDrop.SetListsFromTable modRanges.GetTable(hostSheet, NAME_DATA_TABLE), LISTS_HEADER_ITEM_ID, LISTS_HEADER_OBJECTS
                              
        m_mngrMap.Add key, formDrop
    End If

    Set FD_GetOrCreateMngr = m_mngrMap(key)
CleanExit:
    Exit Function
ErrHandler:
    modErr.ReportError "FD_GetOrCreateMngr", Err.Number, Erl
    Resume CleanExit
End Function

' -----------------------------------------------------------------------------------
' Function  : FD_IsAnchorCell
' Purpose   : True if the cell has a name matching FD_ANCHOR_NAME_PATTERN.
'
' Parameters:
'   hostSheet    [Worksheet] - Host sheet.
'   cell         [Range]     - Single cell to test.
'
' Returns   : [Boolean]      - True if the cell is an anchor cell; otherwise False.
' -----------------------------------------------------------------------------------
Private Function FD_IsAnchorCell(ByVal hostSheet As Worksheet, ByVal cell As Range) As Boolean
    Dim nm As Name

    Set nm = modRanges.GetCellNameByPattern(cell, FD_ANCHOR_NAME_PATTERN)
    If Not nm Is Nothing Then
        FD_IsAnchorCell = True
        Exit Function
    End If
End Function


' -----------------------------------------------------------------------------------
' Function  : FD_ConfigListsFromAnchorMeta
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
Private Function FD_ConfigListsFromAnchorMeta(ByVal hostSheet As Worksheet, ByVal anchorCell As Range, formDrop As clsFormDrop) As Boolean
    
    Dim nm As Name
    Set nm = modRanges.GetCellNameByPattern(anchorCell, FD_ANCHOR_NAME_PATTERN)
    If nm Is Nothing Then Exit Function
    
    Dim meta As String
    On Error Resume Next
    meta = CStr(nm.Comment)
    On Error GoTo 0
    If Len(meta) = 0 Then Exit Function

    ' Must start with "FD:"
    If StrComp(Left$(Trim$(meta), Len(FD_META_PREFIX)), FD_META_PREFIX, vbTextCompare) <> 0 Then Exit Function

    meta = Trim$(Mid$(meta, Len(FD_META_PREFIX) + 1))
    Dim kv As Scripting.Dictionary: Set kv = FD_ParseMetaKeyValues(meta)
    If kv Is Nothing Then Exit Function

    Dim wb As Workbook: Set wb = hostSheet.Parent

    ' --- Variant A: workbook Names (cat + subs) OR fallback to subs as category labels ---
    If kv.Exists(FD_META_KEY_SUBS) Then
        Dim subsArr As Variant
        subsArr = modUtil.SplitTrim(CStr(kv(FD_META_KEY_SUBS)), FD_META_LIST_SEP)

        Dim i As Long
        Dim subRanges() As Variant
        ReDim subRanges(LBound(subsArr) To UBound(subsArr))
        For i = LBound(subsArr) To UBound(subsArr)
            Dim nr As Range
            On Error Resume Next
            Set nr = wb.Names(CStr(subsArr(i))).RefersToRange
            On Error GoTo 0
            If nr Is Nothing Then Exit Function
            Set subRanges(i) = nr
        Next i

        If kv.Exists(FD_META_KEY_CAT) Then
            ' Normalfall: cat= vorhanden
            Dim catRng As Range
            On Error Resume Next
            Set catRng = wb.Names(CStr(kv(FD_META_KEY_CAT))).RefersToRange
            On Error GoTo 0
            If catRng Is Nothing Then Exit Function
            formDrop.SetListsFromNamedRanges catRng, subRanges
        Else
            ' Fallback: keine cat= ? subs-Namen als Kategorienamen verwenden
            formDrop.SetListsFromLabelsAndRanges subsArr, subRanges
        End If

        FD_ConfigListsFromAnchorMeta = True
        Exit Function
    End If

    ' --- Variant B: table & columns ---
    If kv.Exists(FD_META_KEY_SUBS_TBL) And kv.Exists(FD_META_KEY_SUBS_COLS) Then
        Dim loSubs As ListObject
        Set loSubs = modRanges.GetTable(hostSheet, CStr(kv(FD_META_KEY_SUBS_TBL)))
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
            On Error GoTo 0
            If rr Is Nothing Then Exit Function
            Set subRanges2(i2) = rr
        Next i2

        If kv.Exists(FD_META_KEY_CAT_TBL) And kv.Exists(FD_META_KEY_CAT_COL) Then
            ' Normalfall: catTbl + catCol vorhanden
            Dim loCat As ListObject, catRng2 As Range
            Set loCat = modRanges.GetTable(hostSheet, CStr(kv(FD_META_KEY_CAT_TBL)))
            If loCat Is Nothing Then Exit Function
            On Error Resume Next
            Set catRng2 = loCat.ListColumns(CStr(kv(FD_META_KEY_CAT_COL))).DataBodyRange
            On Error GoTo 0
            If catRng2 Is Nothing Then Exit Function
            formDrop.SetListsFromNamedRanges catRng2, subRanges2
        Else
            ' Fallback: keine catTbl/catCol ? Spaltennamen als Kategorienamen verwenden
            formDrop.SetListsFromLabelsAndRanges cols, subRanges2
        End If

        FD_ConfigListsFromAnchorMeta = True
        Exit Function
    End If
End Function

' -----------------------------------------------------------------------------------
' Function  : FD_ParseMetaKeyValues
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
Private Function FD_ParseMetaKeyValues(ByVal metaString As String) As Scripting.Dictionary
    Dim metaDict As Scripting.Dictionary: Set metaDict = New Scripting.Dictionary
    metaDict.CompareMode = 1 ' TextCompare
    Dim parts As Variant, i As Long
    parts = Split(metaString, FD_META_PAIR_SEP)
    For i = LBound(parts) To UBound(parts)
        Dim part As String: part = Trim$(parts(i))
        If Len(part) > 0 Then
            Dim eqPos As Long: eqPos = InStr(1, part, "=", vbTextCompare)
            If eqPos > 1 Then
                Dim key As String, value As String
                key = Trim$(Left$(part, eqPos - 1))
                value = Trim$(Mid$(part, eqPos + 1))
                If Len(key) > 0 Then metaDict(key) = value
            End If
        End If
    Next
    Set FD_ParseMetaKeyValues = metaDict
End Function



