Attribute VB_Name = "modFormDropManager"
' -----------------------------------------------------------------------------------
' Module    : modFormDropManager
' Purpose   : Business logic for "wandering" Form DropDown pair (arrow + 2 dropdowns).
'             One clsFormDrop manager per worksheet, reused and moved on selection.
'
' Public API:
'   - FD_Init()                     : Prepare feature (per-app lifetime).
'   - FD_Cleanup()                  : Destroy all per-sheet managers.
'   - FD_OnSelectionChange(Sh, Target) : Central selection handler (shows arrow / hides).
'   - FD_OnSheetDeactivate(Sh)      : Hide on sheet switch.
'
' Dependencies:
'   - clsFormDrop (main class)
'   - modFormDropRouter (routes Form control events)
'   - modTags
'
' Notes     :
'   - Requires reference "Microsoft Scripting Runtime" (Scripting.Dictionary).
'   - This module owns the per-worksheet clsFormDrop instances.
' -----------------------------------------------------------------------------------
Option Explicit
Option Private Module

Private m_mgrMap As Scripting.Dictionary   ' key = Worksheet.CodeName, value = clsFormDrop

' -----------------------------------------------------------------------------------
' Procedure : FD_Init
' Purpose   : Initialize feature: create instance map.
'
' Parameters:
'   (none)
'
' Returns   : (none)
'
' Notes     : Safe to call multiple times.
' -----------------------------------------------------------------------------------
Public Sub FD_Init()
On Error GoTo ErrHandler
    If m_mgrMap Is Nothing Then Set m_mgrMap = New Scripting.Dictionary
CleanExit:
    Exit Sub
ErrHandler:
    modErr.ReportError "FD_Init", Err.Number, Erl
    Resume CleanExit
End Sub

' -----------------------------------------------------------------------------------
' Procedure : FD_Cleanup
' Purpose   : Destroy all managers and clear map.
'
' Parameters:
'   (none)
'
' Returns   : (none)
'
' Notes     : Safe to call multiple times.
' -----------------------------------------------------------------------------------
Public Sub FD_Cleanup()
On Error GoTo ErrHandler
    Dim k As Variant
    If Not m_mgrMap Is Nothing Then
        For Each k In m_mgrMap.Keys
            On Error Resume Next
            m_mgrMap(k).Destroy
            On Error GoTo ErrHandler
        Next k
        m_mgrMap.RemoveAll
    End If
CleanExit:
    Exit Sub
ErrHandler:
    modErr.ReportError "FD_Cleanup", Err.Number, Erl
    Resume CleanExit
End Sub

' -----------------------------------------------------------------------------------
' Procedure : FD_OnSelectionChange
' Purpose   : Central handler for App.SheetSelectionChange.
'             Shows arrow only when the selected cell is an anchor (named DD_Anchor_*).; hides otherwise. Dropdowns open on arrow click.
'
' Parameters:
'   Sh       [Object] - Worksheet raising the event.
'   Target   [Range]  - New selection.
'
' Returns   : (none)
'
' Notes     :
'   - Requires that lists were set in clsFormDrop; otherwise ShowAt raises (as desired).
' -----------------------------------------------------------------------------------
Public Sub FD_OnSelectionChange(ByVal Sh As Object, ByVal Target As Range)
On Error GoTo ErrHandler

    Dim ws As Worksheet: Set ws = Sh
    If Not FD_ShouldHandleSheet(ws) Then Exit Sub

    Dim mgr As clsFormDrop: Set mgr = FD_EnsureMgr(ws)

    If Target Is Nothing Or Target.CountLarge > 1 Then
        mgr.HideDropDowns
        Exit Sub
    End If

    If FD_IsAnchorCell(ws, Target.Cells(1, 1)) Then
        Call FD_TryConfigureListsFromAnchor(ws, Target.Cells(1, 1), mgr)
        mgr.ShowAt ws, Target.Cells(1, 1)   ' shows arrow only; dropdown opens on arrow click
    Else
        mgr.HideDropDowns
    End If

CleanExit:
    Exit Sub
ErrHandler:
    modErr.ReportError "FD_OnSelectionChange", Err.Number, Erl
    Resume CleanExit
End Sub

' -----------------------------------------------------------------------------------
' Procedure : FD_OnSheetDeactivate
' Purpose   : Hide controls when leaving a sheet.
'
' Parameters:
'   Sh       [Object] - Worksheet being deactivated.
'
' Returns   : (none)
' -----------------------------------------------------------------------------------
Public Sub FD_OnSheetDeactivate(ByVal Sh As Object)
On Error GoTo ErrHandler
    If m_mgrMap Is Nothing Then Exit Sub
    Dim key As String: key = Sh.CodeName
    If m_mgrMap.Exists(key) Then m_mgrMap(key).Hide
CleanExit:
    Exit Sub
ErrHandler:
    modErr.ReportError "FD_OnSheetDeactivate", Err.Number, Erl
    Resume CleanExit
End Sub

' -----------------------------------------------------------------------------------
' Function  : FD_GetAnchors
' Purpose   : Returns a union range of all cells on worksheet that are named DD_Anchor_*.
'
' Parameters:
'   ws  [Worksheet]
'
' Returns   : [Range] - Union of anchor cells; Nothing if none.
'
' Note:
' - only sheet-level names are searched
' -----------------------------------------------------------------------------------
Private Function FD_GetAnchors(ByVal ws As Worksheet) As Range
On Error GoTo ErrHandler
    Dim nm As Name, r As Range, acc As Range

    ' sheet-level names
    For Each nm In ws.Names
        If nm.Name Like FD_ANCHOR_NAME_PATTERN Then
            If Not nm.RefersToRange Is Nothing Then
                Set r = nm.RefersToRange
                If r.Parent Is ws Then
                    If r.CountLarge = 1 Then
                        If acc Is Nothing Then Set acc = r Else Set acc = Application.Union(acc, r)
                    End If
                End If
            End If
        End If
    Next nm

    Set FD_GetAnchors = acc
CleanExit:
    Exit Function
ErrHandler:
    Set FD_GetAnchors = Nothing
    Resume CleanExit
End Function

' -----------------------------------------------------------------------------------
' Function  : FD_ShouldHandleSheet
' Purpose   : Determine whether the given sheet participates in the feature.
'
' Parameters:
'   ws       [Worksheet] - Candidate sheet.
'
' Returns   : [Boolean] - True if the sheet should be handled.
'
' Notes     :
'   - Mirrors your gating in clsAppEvents: skip add-in workbook/sheet dispatcher.
' -----------------------------------------------------------------------------------
Private Function FD_ShouldHandleSheet(ByVal ws As Worksheet) As Boolean
On Error GoTo ErrHandler
    Dim wb As Workbook: Set wb = ws.Parent
    If wb Is RDDAddInWkBk Then GoTo CleanExit
    If wb.IsAddin Then GoTo CleanExit
    If Not modTags.SheetWithTagExists(wb, SHEET_DISPATCHER) Then GoTo CleanExit
    If ws.CodeName = SHEET_DISPATCHER Then GoTo CleanExit
    If modTags.HasSheetTag(ws, SHEET_DISPATCHER) Then GoTo CleanExit

    FD_ShouldHandleSheet = True
CleanExit:
    Exit Function
ErrHandler:
    ' On any error: be conservative and do not handle
    FD_ShouldHandleSheet = False
    Resume CleanExit
End Function


' -----------------------------------------------------------------------------------
' Function  : FD_EnsureMgr
' Purpose   : Get or create the per-sheet clsFormDrop manager.
'
' Parameters:
'   ws       [Worksheet] - Host sheet.
'
' Returns   : [clsFormDrop] - Configured manager instance.
'
' Notes     :
'   - Sets arrow options and list sources once on first use.
' -----------------------------------------------------------------------------------
Private Function FD_EnsureMgr(ByVal ws As Worksheet) As clsFormDrop
On Error GoTo ErrHandler
    If m_mgrMap Is Nothing Then FD_Init

    Dim key As String: key = ws.CodeName
    If Not m_mgrMap.Exists(key) Then
        Dim mgr As New clsFormDrop
        mgr.Init ws.Parent, modFormDropRouter.g_formDropRegistryDict

        ' Arrow config
        mgr.SetArrowEnabled True
        mgr.SetArrowStyle 2, 10, 10
        mgr.SetPlacement True

        ' Provide list sources
        mgr.SetListsFromNames FD_CAT_RANGE_NAME, _
                              FD_SUB_ITEMS_RANGE_NAME, _
                              FD_SUB_OBJECTS_RANGE_NAME, _
                              FD_SUB_HOTSPOTS_RANGE_NAME, _
                              FD_SUB_ACTORS_RANGE_NAME

        m_mgrMap.Add key, mgr
    End If

    Set FD_EnsureMgr = m_mgrMap(key)
CleanExit:
    Exit Function
ErrHandler:
    modErr.ReportError "FD_EnsureMgr", Err.Number, Erl
    Resume CleanExit
End Function

' -----------------------------------------------------------------------------------
' Function  : FD_IsAnchorCell
' Purpose   : True if the cell has a name matching FD_ANCHOR_NAME_PATTERN.
'
' Parameters:
'   ws    [Worksheet] - Host sheet.
'   cell  [Range]     - Single cell to test.
'
' Returns   : [Boolean]
' -----------------------------------------------------------------------------------
Private Function FD_IsAnchorCell(ByVal ws As Worksheet, ByVal cell As Range) As Boolean
    Dim nm As Name

    Set nm = modRanges.GetCellNameByPattern(cell, FD_ANCHOR_NAME_PATTERN)
    If Not nm Is Nothing Then
        FD_IsAnchorCell = True
        Exit Function
    End If
End Function


' Try to configure lists for an anchor cell from the Name.Comment.
' Returns True if configuration was applied; False if no/invalid metadata.
Private Function FD_TryConfigureListsFromAnchor(ByVal ws As Worksheet, ByVal cell As Range, mgr As clsFormDrop) As Boolean
    
    Dim nm As Name
    Set nm = modRanges.GetCellNameByPattern(cell, FD_ANCHOR_NAME_PATTERN)
    If nm Is Nothing Then Exit Function
    
    Dim meta As String
    On Error Resume Next
    meta = CStr(nm.Comment)
    On Error GoTo 0
    If Len(meta) = 0 Then Exit Function

    ' Must start with "FD:"
    If StrComp(Left$(Trim$(meta), Len(FD_META_PREFIX)), FD_META_PREFIX, vbTextCompare) <> 0 Then Exit Function

    meta = Trim$(Mid$(meta, Len(FD_META_PREFIX) + 1))
    Dim kv As Object: Set kv = FD_ParseMetaKeyValues(meta)
    If kv Is Nothing Then Exit Function

    Dim wb As Workbook: Set wb = ws.Parent

    ' --- Variant A: workbook Names (cat + subs) OR fallback to subs as category labels ---
    If kv.Exists(FD_META_KEY_SUBS) Then
        Dim subsArr As Variant
        subsArr = FD_SplitTrim(CStr(kv(FD_META_KEY_SUBS)), FD_META_LIST_SEP)

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
            mgr.SetListsFromNamedRanges catRng, subRanges
        Else
            ' Fallback: keine cat= ? subs-Namen als Kategorienamen verwenden
            mgr.SetListsFromLabelsAndRanges subsArr, subRanges
        End If

        FD_TryConfigureListsFromAnchor = True
        Exit Function
    End If

    ' --- Variant B: table & columns ---
    If kv.Exists(FD_META_KEY_SUBS_TBL) And kv.Exists(FD_META_KEY_SUBS_COLS) Then
        Dim loSubs As ListObject
        Set loSubs = FD_GetTable(ws, CStr(kv(FD_META_KEY_SUBS_TBL)))
        If loSubs Is Nothing Then Exit Function

        Dim cols As Variant
        cols = FD_SplitTrim(CStr(kv(FD_META_KEY_SUBS_COLS)), FD_META_LIST_SEP)

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
            Set loCat = FD_GetTable(ws, CStr(kv(FD_META_KEY_CAT_TBL)))
            If loCat Is Nothing Then Exit Function
            On Error Resume Next
            Set catRng2 = loCat.ListColumns(CStr(kv(FD_META_KEY_CAT_COL))).DataBodyRange
            On Error GoTo 0
            If catRng2 Is Nothing Then Exit Function
            mgr.SetListsFromNamedRanges catRng2, subRanges2
        Else
            ' Fallback: keine catTbl/catCol ? Spaltennamen als Kategorienamen verwenden
            mgr.SetListsFromLabelsAndRanges cols, subRanges2
        End If

        FD_TryConfigureListsFromAnchor = True
        Exit Function
    End If
End Function

' Parse "key=value; key=value" into a late-bound dictionary (case-insensitive keys)
Private Function FD_ParseMetaKeyValues(ByVal s As String) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1 ' TextCompare
    Dim parts As Variant, i As Long
    parts = Split(s, FD_META_PAIR_SEP)
    For i = LBound(parts) To UBound(parts)
        Dim p As String: p = Trim$(parts(i))
        If Len(p) > 0 Then
            Dim eqPos As Long: eqPos = InStr(1, p, "=", vbTextCompare)
            If eqPos > 1 Then
                Dim k As String, v As String
                k = Trim$(Left$(p, eqPos - 1))
                v = Trim$(Mid$(p, eqPos + 1))
                If Len(k) > 0 Then d(k) = v
            End If
        End If
    Next
    Set FD_ParseMetaKeyValues = d
End Function

' Split by sep and trim entries; returns a Variant array (LB=0)
Private Function FD_SplitTrim(ByVal s As String, ByVal sep As String) As Variant
    Dim arr As Variant: arr = Split(s, sep)
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        arr(i) = Trim$(CStr(arr(i)))
    Next
    FD_SplitTrim = arr
End Function

' Get a ListObject by name and worksheet
Private Function FD_GetTable(ByVal ws As Worksheet, ByVal tableName As String) As ListObject
    On Error Resume Next
    Set FD_GetTable = ws.ListObjects(tableName)
    On Error GoTo 0
End Function



