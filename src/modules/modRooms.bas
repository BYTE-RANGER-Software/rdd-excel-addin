Attribute VB_Name = "modRooms"
' -----------------------------------------------------------------------------------
' Module    : modRooms
' Purpose   : Create, initialize and manage "Room" worksheets; aggregate list data.
'
' Public API:
'   - AddRoom                 : Create a new room sheet from template and initialize it.
'   - IsRoomSheet             : Check whether a sheet is a room sheet (by tag).
'   - RemoveRoom              : Delete a room sheet after reference checks.
'   - GetNextRoomIndex        : Compute next numeric room index.
'   - HasRoomID            : Determine whether a room with specific ID exists.
'   - UpdateLists             : Append/merge room/object/scene items into Lists table.
'   - SyncLists               : Rebuild Lists columns (clear + write).
'   - GetFormattedRoomID      : Build a formatted room ID from an index.
'   - ApplyParallaxRangeCover : Ensure/show/hide the parallax cover according to a dropdown.
'
' Dependencies: modProps, modUtil, modTags, modRanges, modSheets, modLists, modRangeCover, modConst
' Notes     :
'   - Public API section exposes operations used by UI/other modules.
'   - Private Helpers contain internal utilities for this module only.
' -----------------------------------------------------------------------------------
Option Explicit
Option Private Module

Public Enum ListUpdateMode
    LUM_Append = 0 ' only Add new Datas
    LUM_Sync = 1   ' Rewrite all Datas new
    LUM_OnlyRoomRelated = 2 ' Rewrite only Room related Data
End Enum

' ===== Public API ==================================================================

' -----------------------------------------------------------------------------------
' Function  : AddRoom
' Purpose   : Clones required templates, ensures helper sheets exist, creates a new
'             Room sheet with the next sequential ID, initializes it, and activates it.
'
' Parameters:
'   targetBook            [Workbook] - Target workbook.
'   newName               [String]   - Name for the new Room sheet (already computed).
'   roomIdx               [Long]     - Numeric index used for the Room ID formatting.
'   updateAggregations    [Boolean]  - Optional. If True, UpdateLists is called after creation.
'
' Returns   : [Worksheet] - The newly created Room worksheet.
'
' Notes     :
'   - Ensures dispatcher and lists sheets by copying them from RDDAddInWkBk if missing.
'   - Creates a new sheet from SHEET_ROOM_TEMPLATE and tags it with ROOM_SHEET_ID_TAG_NAME.
'   - Calls SetupRoom to wire controls/values; toggles HideOpMode during operations.
' -----------------------------------------------------------------------------------
Public Function AddRoom(ByVal targetBook As Workbook, ByRef newName As String, ByRef roomIdx As Long, Optional ByVal updateAggregations As Boolean = True) As Worksheet
    On Error GoTo ErrHandler

    Dim tmplSheet As Worksheet
    Dim newRoomSheet As Worksheet
    
    modUtil.HideOpMode True
            
    If Not modSheets.SheetCodeNameExists(modConst.SHEET_DISPATCHER, targetBook) And Not modTags.SheetWithTagExists(targetBook, SHEET_DISPATCHER) Then
    
        Set tmplSheet = RDDAddInWkBk.Worksheets(modConst.SHEET_DISPATCHER)
        tmplSheet.Visible = xlSheetVisible
        
        tmplSheet.Copy After:=targetBook.Sheets(targetBook.Sheets.Count)
        Set newRoomSheet = ActiveSheet 'targetBook.Sheets(targetBook.Sheets.Count)
        
        newRoomSheet.Visible = xlSheetHidden
        newRoomSheet.Name = SHEET_DISPATCHER
        modProps.ClearAllCustomProperties newRoomSheet
        modTags.TagSheet newRoomSheet, SHEET_DISPATCHER
        
        Set tmplSheet = Nothing
        Set newRoomSheet = Nothing
    End If
            
    Set tmplSheet = RDDAddInWkBk.Worksheets(modConst.SHEET_ROOM_TEMPLATE)
    tmplSheet.Visible = xlSheetVisible
    
    tmplSheet.Copy After:=targetBook.Sheets(targetBook.Sheets.Count)
    Set newRoomSheet = ActiveSheet ' targetBook.Sheets(targetBook.Sheets.Count)
    
    newRoomSheet.Name = newName
    
    modProps.ClearAllCustomProperties newRoomSheet
    modTags.TagSheet newRoomSheet, ROOM_SHEET_ID_TAG_NAME, GetFormattedRoomID(roomIdx)
    
    SetupRoom newRoomSheet, roomIdx
    
    Set AddRoom = newRoomSheet
        
    If updateAggregations Then
        UpdateLists targetBook, LUM_OnlyRoomRelated
    End If
    
CleanExit:
    modUtil.HideOpMode False
    Exit Function
    
ErrHandler:
    modErr.ReportError "modRooms.AddRoom", Err.Number, Erl, caption:=modMain.AppProjectName
    Resume CleanExit
End Function

' -----------------------------------------------------------------------------------
' Function  : IsRoomSheet
' Purpose   : Checks whether the given sheet is a Room sheet (identified by tag).
'
' Parameters:
'   targetSheet       [Worksheet]         - Sheet to test.
'   returnRoomId      [String]            - (Optional, ByRef) Receives the room ID if it is a room sheet.
'
' Returns   : [Boolean] - True if the sheet is a Room sheet.
' -----------------------------------------------------------------------------------
Public Function IsRoomSheet(ByRef targetSheet As Worksheet, Optional ByRef returnRoomId As String = vbNullString) As Boolean
    Dim tagValue As String
    If modTags.HasSheetTag(targetSheet, ROOM_SHEET_ID_TAG_NAME, returnRoomId) Then
        IsRoomSheet = True
    End If
End Function

' -----------------------------------------------------------------------------------
' Function  : RemoveRoom
' Purpose   : Deletes the given Room sheet after verifying that no other Room sheet
'             references it; updates aggregated lists.
'
' Parameters:
'   targetSheet         [Worksheet]   - Room sheet to delete (must not be Nothing).
'   updateAggregations  [Boolean]     - If True, UpdateLists is called after deletion.
'   returnRefSheets     [Collection]  - (Optional, ByRef) receives referencing sheet names.
'
' Returns   : Boolean - True on success; False is not used (errors are raised/logged).
'
' Notes     :
'   - Checks references via GetAllSheetNamesUsingRoomID.
'   - Uses HideOpMode during deletion; hands back referencing sheets if found.
' -----------------------------------------------------------------------------------
Public Function RemoveRoom(ByVal targetSheet As Worksheet, _
    Optional ByVal updateAggregations As Boolean = True, _
    Optional ByRef returnRefSheets As Collection) As Boolean
    
    On Error GoTo ErrHandler

    Dim parentBook As Workbook: Set parentBook = targetSheet.Parent
    Dim roomID As String

    If targetSheet Is Nothing Then
        Err.Raise ERR_ARG_NULL_TARGETSHEET, "modRooms.RemoveRoom", "Argument 'targetSheet' must not be Nothing."
    End If
    
    If Not IsRoomSheet(targetSheet, roomID) Then
        Err.Raise ERR_NOT_A_ROOM_SHEET, "modRooms.RemoveRoom", "The provided sheet is not a Room sheet."
    End If
        
    ' Check references to the active room sheet in all other room sheets
    Dim usedByCol As Collection
    Set usedByCol = GetAllSheetNamesUsingRoomID(roomID, parentBook, targetSheet)
    If Not usedByCol Is Nothing Then
        If usedByCol.Count > 0 Then
            ' Hand result back to caller for UI, then raise a error
            Set returnRefSheets = usedByCol
            Err.Raise ERR_ROOM_SHEET_REFERENCED, _
                "modRooms.RemoveRoom", _
                "Room sheet cannot be deleted because it is referenced by other Room sheets."
        End If
    End If
        
    modUtil.HideOpMode True
    targetSheet.Delete
    Set targetSheet = Nothing
    
    If updateAggregations Then
        UpdateLists targetSheet.Parent, LUM_Sync
    End If
    
CleanExit:
    modUtil.HideOpMode False
    Exit Function
    
ErrHandler:
    modErr.ReportError "RemoveRoom", Err.Number, Erl, caption:=modMain.AppProjectName
    Resume CleanExit
End Function

' -----------------------------------------------------------------------------------
' Procedure : UpdateRoomReferences
' Purpose   : Updates all references to a room ID and alias throughout the workbook.
' -----------------------------------------------------------------------------------
Public Function UpdateRoomReferences(ByVal targetBook As Workbook, _
                                ByVal oldRoomID As String, _
                                ByVal oldRoomAlias As String, _
                                ByVal newRoomID As String, _
                                ByVal newRoomAlias As String) As Long
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim roomID As String
    Dim doorsToRoomIDRange As Range
    Dim doorsToRoomAliasRange As Range
    Dim cell As Range
    Dim updatedCount As Long
    
    For Each ws In targetBook.Worksheets
        If modRooms.IsRoomSheet(ws, roomID) Then
            
            ' Update Room ID and Room Alias Cell
            On Error Resume Next
            Set doorsToRoomIDRange = ws.Range(NAME_RANGE_DOORS_TO_ROOM_ID)
            Set doorsToRoomAliasRange = ws.Range(NAME_RANGE_DOORS_TO_ROOM_ALIAS)
            On Error GoTo ErrHandler
            
            ' Update Framed Range "DOORS TO..."
            If Not doorsToRoomIDRange Is Nothing Then
                For Each cell In doorsToRoomIDRange.Cells
                    If Not IsEmpty(cell.value) Then
                        If StrComp(CStr(cell.value), oldRoomID, vbTextCompare) = 0 Then
                            cell.value = newRoomID
                            updatedCount = updatedCount + 1
                        End If
                    End If
                Next cell
            End If
            
            If Not doorsToRoomAliasRange Is Nothing Then
                For Each cell In doorsToRoomAliasRange.Cells
                    If Not IsEmpty(cell.value) Then
                        If StrComp(CStr(cell.value), oldRoomAlias, vbTextCompare) = 0 Then
                            cell.value = newRoomAlias
                        End If
                    End If
                Next cell
            End If
            
        End If
    Next ws
    
    Call UpdateDispatcherTable(targetBook, oldRoomID, oldRoomAlias, newRoomID, newRoomAlias)
    
    UpdateRoomReferences = updatedCount
    
    Exit Function
    
ErrHandler:
    modErr.ReportError "modMain.UpdateRoomReferences", Err.Number, Erl, caption:=AppProjectName
End Function

' -----------------------------------------------------------------------------------
' Function  : GetNextRoomIndex
' Purpose   : Computes the next free numeric index by scanning existing Room sheets
'             and returning (max index + 1).
'
' Parameters:
'   targetBook      [Workbook] - Workbook to scan
'
' Returns   : Long - Next available Room index.
'
' Notes     :
'   - Detects Room sheets via ROOM_SHEET_ID_TAG_NAME and parses the ID.
'
' -----------------------------------------------------------------------------------
Public Function GetNextRoomIndex(ByVal targetBook As Workbook) As Long
    On Error GoTo ErrHandler
    
    Dim sheet As Worksheet, numIdx As Long, maxIdx As Long
    Dim tagValue As String
    Dim resultIdx As Long
    
    For Each sheet In targetBook.Worksheets
        If modTags.HasSheetTag(sheet, ROOM_SHEET_ID_TAG_NAME, tagValue) Then
            numIdx = val(Mid$(CStr(tagValue), Len(ROOM_SHEET_ID_TAG_VAL_PRE) + 1))
            If numIdx > maxIdx Then maxIdx = numIdx
        End If
    Next sheet
    
    resultIdx = maxIdx + 1

CleanExit:
    GetNextRoomIndex = resultIdx
    Exit Function

ErrHandler:
    modErr.ReportError "GetNextRoomIndex", Err.Number, Erl, caption:=modMain.AppProjectName
    Resume CleanExit
End Function

' -----------------------------------------------------------------------------------
' Function  : HasRoomID
' Purpose   : Determine whether the workbook contains a Room sheet with the given ID.
'
' Parameters:
'   targetBook       [Workbook]              - Workbook to scan.
'   roomId           [String]                - Room ID to search for (e.g., "R001").
'   returnSheet      [Worksheet]             - (Optional ByRef) Receives the first matching sheet if found.
'
' Returns   : Boolean - True if found; otherwise False.
'
' Notes     :
' -----------------------------------------------------------------------------------
Public Function HasRoomID(ByVal targetBook As Workbook, ByVal roomID As String, Optional ByRef returnSheet As Worksheet = Nothing) As Boolean
    Dim sheet As Worksheet
    Dim tagValue As String
    Dim isFound As Boolean
    
    On Error GoTo ErrHandler
    
    For Each sheet In targetBook.Worksheets
        If modTags.HasSheetTag(sheet, ROOM_SHEET_ID_TAG_NAME, tagValue) Then
            If StrComp(roomID, tagValue, vbBinaryCompare) = 0 Then
                Set returnSheet = sheet
                isFound = True
                GoTo CleanExit
            End If
        End If
    Next sheet
    
CleanExit:
    HasRoomID = isFound
    Exit Function

ErrHandler:
    modErr.ReportError "HasRoomID", Err.Number, Erl, caption:=modMain.AppProjectName
    Resume CleanExit
End Function

' -----------------------------------------------------------------------------------
' Function  : HasRoomAlias
' Purpose   : Determine whether the workbook contains a Room sheet with the given Alias.
'
' Parameters:
'   targetBook       [Workbook]              - Workbook to scan.
'   roomAlis         [String]                - Room Alias to search for (e.g., "r_TempleEntrance").
'   returnSheet      [Worksheet]             - (Optional ByRef) Receives the first matching sheet if found.
'
' Returns   : Boolean - True if found; otherwise False.
'
' Notes     :
' -----------------------------------------------------------------------------------
Public Function HasRoomAlias(ByVal targetBook As Workbook, ByVal roomAlias As String, Optional ByRef returnSheet As Worksheet = Nothing) As Boolean
    Dim sheet As Worksheet
    Dim tagValue As String
    Dim isFound As Boolean
    Dim cell As Range
    
    On Error GoTo ErrHandler
    
    For Each sheet In targetBook.Worksheets
        If modTags.HasSheetTag(sheet, ROOM_SHEET_ID_TAG_NAME) Then
            On Error Resume Next
            Set cell = sheet.Range(modConst.NAME_CELL_ROOM_ALIAS)
            On Error GoTo ErrHandler
            
            If Not cell Is Nothing Then
                If StrComp(roomAlias, cell.value, vbBinaryCompare) = 0 Then
                    Set returnSheet = sheet
                    isFound = True
                    GoTo CleanExit
                End If
            End If
        End If
    Next sheet
    
CleanExit:
    HasRoomAlias = isFound
    Exit Function

ErrHandler:
    modErr.ReportError "HasRoomAlias", Err.Number, Erl, caption:=modMain.AppProjectName
    Resume CleanExit
End Function

' -----------------------------------------------------------------------------------
' Procedure : UpdateLists
' Purpose   : Central function for all Room Datalist update operations.
'             Collects data and writes it differently depending on the mode.
'
' Parameters:
'   targetBook  [Workbook]        - Target Workbook
'   mode        [ListUpdateMode]  - Update-Modus (Append/Sync/NewRoomOnly)
'
' Notes     :
' -----------------------------------------------------------------------------------
Public Sub UpdateLists(targetBook As Workbook, ByVal mode As ListUpdateMode)
    On Error GoTo ErrHandler
    
    ' Enable silent mode
    modUtil.HideOpMode True, affectScreen:=False, affectEvents:=False
    
    ' Set workbook
    If targetBook Is Nothing Then Set targetBook = ActiveWorkbook
    
    ' Determine Lists Sheet and DataTable
    Dim listsSheet As Worksheet
    Dim dataList As ListObject
    If Not GetListsSheetAndTable(targetBook, listsSheet, dataList) Then
        Exit Sub
    End If
    
    ' Initialize dictionaries
    Dim roomsDict As Scripting.Dictionary
    Dim scenesDict As Scripting.Dictionary
    Dim itemObjectsDict As Scripting.Dictionary
    Dim stateObjectsDict As Scripting.Dictionary
    Dim hotspotObjectsDict As Scripting.Dictionary
    
    ' Collect data from all room sheets
    Dim collectObjects As Boolean
    collectObjects = (mode <> LUM_OnlyRoomRelated)  ' Do not collect property data for NewRoomOnly
    
    CollectRoomData targetBook, roomsDict, scenesDict, itemObjectsDict, _
                    stateObjectsDict, hotspotObjectsDict, collectObjects
    
    ' Process differently depending on mode
    Select Case mode
        Case LUM_Sync
            ' SYNC mode: Delete all columns and rewrite them
            WriteSyncMode dataList, roomsDict, scenesDict, itemObjectsDict, _
                         stateObjectsDict, hotspotObjectsDict
                         
        Case LUM_Append, LUM_OnlyRoomRelated
            ' APPEND mode: Add only new entries
            WriteAppendMode dataList, roomsDict, scenesDict, itemObjectsDict, _
                           stateObjectsDict, hotspotObjectsDict, collectObjects
    End Select

CleanExit:
    modUtil.HideOpMode False
    Exit Sub
    
ErrHandler:
    modUtil.HideOpMode False
    modErr.ReportError "ProcessListsUpdate", Err.Number, Erl, caption:=modMain.AppProjectName
    Resume CleanExit
End Sub

' -----------------------------------------------------------------------------------
' Function  : GetFormattedRoomID
' Purpose   : Build a formatted Room ID from the numeric index using a prefix.
'
' Parameters:
'   roomIdx [Long] - Numeric index.
'
' Returns   : String - e.g., R001 (depends on ROOM_SHEET_ID_TAG_VAL_PRE).
'
' Notes     :
' -----------------------------------------------------------------------------------
Public Function GetFormattedRoomID(ByVal roomIdx As Long) As String
    GetFormattedRoomID = ROOM_SHEET_ID_TAG_VAL_PRE & Format(roomIdx, "000")
End Function

' -----------------------------------------------------------------------------------
' Procedure : ApplyParallaxRangeCover
' Purpose   : Ensures and toggles a named range "cover" according to a dropdown state.
'
' Parameters:
'   targetSheet [Worksheet] - Target Room worksheet.
'
' Returns   : (none)
' -----------------------------------------------------------------------------------
Public Sub ApplyParallaxRangeCover(targetSheet As Worksheet)
    modRangeCover.EnsureRangeCover targetSheet, targetSheet.Range(NAME_RANGE_ADD_PARALLAX_SET_WITH_HEADER), ROOM_ADD_PARALLAX_SET_COVER_NAME
    modRangeCover.ApplyCoverVisibilityByDropdown targetSheet, NAME_CELL_PARALLAX, ROOM_ADD_PARALLAX_SET_HIDE_TOKEN, ROOM_ADD_PARALLAX_SET_COVER_NAME
End Sub

' ===== Private Helpers =============================================================

' -----------------------------------------------------------------------------------
' Function  : SetupRoom
' Purpose   : Initializes a newly created Room sheet: sets RoomID, removes stale
'             name links, and wires the "insert room picture" button action.
'
' Parameters:
'   targetSheet       [Worksheet]  - Target Room worksheet to initialize.
'   roomIdx    [Long]       - Numeric index used to format the Room ID cell value.
'
' Returns   : none
'
' Notes     :
' -----------------------------------------------------------------------------------
Private Sub SetupRoom(targetSheet As Worksheet, ByVal roomIdx As Long)
    On Error GoTo ErrHandler
    Dim insertBtnShape As Shape
    Dim dispatcherSheet As Worksheet
    Dim dataRange As Range
    
    'Set 'RoomID' named cell on the template
    targetSheet.Range(modConst.NAME_CELL_ROOM_ID).value = GetFormattedRoomID(roomIdx)
    'Set 'RoomAlias" named Cell on the Template
    targetSheet.Range(modConst.NAME_CELL_ROOM_ALIAS).value = "r_" & GetCleanRoomAlias(targetSheet.Name)
    
    ' remove wrong links
    Dim nm As Name
    For Each nm In targetSheet.Parent.Names
        If InStr(nm.RefersTo, "[" & RDDAddInWkBk.Name & "]") > 0 Then
            nm.Delete
        End If
    Next
    
    ' update button link
    Set dispatcherSheet = modSheets.GetSheetByCodeName(modConst.SHEET_DISPATCHER)
    Set insertBtnShape = targetSheet.Shapes(modConst.BTN_INSERT_ROOM_PICTURE)
    insertBtnShape.OnAction = modConst.MACRO_BTN_INSERT_PICTURE
        
    Exit Sub
    
ErrHandler:
    modErr.ReportError "modRooms.SetupRoom", Err.Number, Erl, caption:=modMain.AppProjectName
    
End Sub

' -----------------------------------------------------------------------------------
' Procedure : ApplyListValidation
' Purpose   : Adds a list data validation to a target range using a named list.
'
' Parameters:
'   target  [Range]  - Target range for the validation.
'   nameRef [String] - Name of the named list (without "=").
'   title   [String] - Input title.
'   msg     [String] - Input message.
'
' Returns   : (none)
' -----------------------------------------------------------------------------------
Private Sub ApplyListValidation(ByVal Target As Range, ByVal nameRef As String, _
                                ByVal title As String, ByVal msg As String)
    With Target.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=" & nameRef
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = title
        .InputMessage = msg
        .ErrorTitle = title & " invalid"
        .ErrorMessage = "Pick a value from the list."
    End With
End Sub

' -----------------------------------------------------------------------------------
' Function  : GetAllsSheetsUsingRoomID
' Purpose   : Find all Room sheets (from a provided dictionary) that reference the
'             given room id either inside the "DOORS TO..." framed range or in the
'             "RoomID" column of the "PUZZLES" area.
'
' Parameters:
'   roomId           [String]  - Room sheet id to search for.
'   targetBook       [Object]  - target Workbook.
'   sheetToExclude        [Worksheet]- (Optional) Sheet to exclude from results.
'
' Returns   : Collection - Sheet names that reference roomId.
'
' Notes     :
' -----------------------------------------------------------------------------------
Private Function GetAllSheetNamesUsingRoomID(ByVal roomID As String, _
                                            ByVal targetBook As Workbook, _
                                            ByVal sheetToExclude As Worksheet) As Collection
    Dim col As New Collection
    Dim targetSheet As Worksheet
    Dim currentRoomID As String
    Dim cell As Range
    
    For Each targetSheet In targetBook.Worksheets
        ' Skip the sheet to be deleted/excluded
        If targetSheet.Name = sheetToExclude.Name Then GoTo NextSheet
        
        ' Only check Room sheets
        If Not IsRoomSheet(targetSheet, currentRoomID) Then GoTo NextSheet
        
        ' Check for references in "DOORS TO..." area
        On Error Resume Next
        Set cell = targetSheet.Range(NAME_RANGE_DOORS_TO_ROOM_ID)
        On Error GoTo 0
        
        If Not cell Is Nothing Then
            If modRanges.RangeHasValue(cell, roomID, True, False) Then
                col.Add targetSheet.Name
                ' Continue checking other sheets for complete list
            End If
        End If
                
NextSheet:
        Set cell = Nothing
    Next targetSheet
    
    Set GetAllSheetNamesUsingRoomID = col
End Function

' -----------------------------------------------------------------------------------
' Function  : GetCleanRoomAlias
' Purpose   : Produces a simplified alias by removing specific punctuation from a name.
'
' Parameters:
'   sourceName [String] - Source name to normalize.
'
' Returns   : String - Cleaned alias string.
' -----------------------------------------------------------------------------------
Private Function GetCleanRoomAlias(ByVal sourceName As String) As String
    Dim removeCharArray() As Variant
    Dim i As Long

    ' Liste der zu entfernenden Zeichen
    removeCharArray = Array(" ", "-", ".", "(", ")", ":", "/", "'")

    ' Alle Zeichen durch leeren String ersetzen
    For i = LBound(removeCharArray) To UBound(removeCharArray)
        sourceName = Replace(sourceName, removeCharArray(i), "")
    Next i

    GetCleanRoomAlias = sourceName
End Function

' -----------------------------------------------------------------------------------
' Function  : GetListsSheetAndTable
' Purpose   : Helper function for determining the list sheet and the DataTable.
'
' Parameters:
'   listSheet           [Worksheet]            - Returned List Sheet
'   dataList            [ListObject]           - Returned Data Table
'
' Returns   : Boolean - True if successful, False if an error occurs
' -----------------------------------------------------------------------------------
Private Function GetListsSheetAndTable(ByVal targetBook As Workbook, _
                                      ByRef listsSheet As Worksheet, _
                                      ByRef dataList As ListObject) As Boolean
    On Error GoTo ErrHandler
    
    ' Determine Lists Sheet
    Set listsSheet = modSheets.GetSheetByCodeName(SHEET_DISPATCHER)
    If listsSheet Is Nothing Then
        Set listsSheet = modTags.GetSheetByTag(targetBook, SHEET_DISPATCHER)
    End If
    
    If listsSheet Is Nothing Then
        Err.Raise ERR_MISSING_DISPATCHER, "GetListsSheetAndTable", _
                  "Lists sheet (Dispatcher) not found in workbook."
        Exit Function
    End If
    
    ' Determine DataTable
    On Error Resume Next
    Set dataList = listsSheet.ListObjects(NAME_DATA_TABLE)
    On Error GoTo ErrHandler
    
    If dataList Is Nothing Then
        Err.Raise ERR_MISSING_DATA_TABLE, "GetListsSheetAndTable", _
                  "Data table '" & NAME_DATA_TABLE & "' not found in Lists sheet."
        Exit Function
    End If
    
    GetListsSheetAndTable = True
    Exit Function
    
ErrHandler:
    modErr.ReportError "GetListsSheetAndTable", Err.Number, Erl, caption:=modMain.AppProjectName
    GetListsSheetAndTable = False
End Function

' -----------------------------------------------------------------------------------
' Procedure : CollectRoomData
' Purpose   : Collect all relevant data from room sheets.
'
' Parameters:
'   targetBook          [Workbook]              - Workbook to be scanned
'   roomsDict           [Dictionary]            - Receives room IDs/aliases
'   scenesDict          [Dictionary]            - Receives scene IDs
'   itemObjectsDict     [Dictionary]            - Receives item objects
'   stateObjectsDict    [Dictionary]            - Receives state objects
'   hotspotObjectsDict  [Dictionary]            - Receives hotspot objects
'   collectObjects      [Boolean]               - If False, no object data is collected.
' -----------------------------------------------------------------------------------
Private Sub CollectRoomData(ByVal targetBook As Workbook, _
                           ByRef roomsDict As Scripting.Dictionary, _
                           ByRef scenesDict As Scripting.Dictionary, _
                           ByRef itemObjectsDict As Scripting.Dictionary, _
                           ByRef stateObjectsDict As Scripting.Dictionary, _
                           ByRef hotspotObjectsDict As Scripting.Dictionary, _
                           ByVal collectObjects As Boolean)
    
    ' Initialize dictionaries
    Set roomsDict = New Scripting.Dictionary
    Set scenesDict = New Scripting.Dictionary
    Set itemObjectsDict = New Scripting.Dictionary
    Set stateObjectsDict = New Scripting.Dictionary
    Set hotspotObjectsDict = New Scripting.Dictionary
    
    ' Header arrays for lookup
    Dim roomAliasHeaders As Variant
    Dim sceneIdHeaders As Variant
    roomAliasHeaders = Array("Room Alias", NAME_CELL_ROOM_ALIAS)
    sceneIdHeaders = Array("Scene ID", NAME_CELL_SCENE_ID)
    
    ' Go through all room sheets
    Dim targetSheet As Worksheet
    Dim roomID As String
    Dim roomAlias As String
    Dim sceneId As String
    
    For Each targetSheet In targetBook.Worksheets
        If modRooms.IsRoomSheet(targetSheet, roomID) Then
            If Len(roomID) > 0 Then
                
                ' Collect room ID and room alias
                roomAlias = modLists.GetNamedOrHeaderValue(targetSheet, NAME_CELL_ROOM_ALIAS, roomAliasHeaders)
                roomsDict(roomID) = roomAlias
                
                ' Collect scene IDs
                sceneId = modLists.GetNamedOrHeaderValue(targetSheet, NAME_CELL_SCENE_ID, sceneIdHeaders)
                If Len(sceneId) > 0 Then scenesDict(sceneId) = True
                
                ' Collect object data (only if desired)
                If collectObjects Then
                    modLists.CollectNamedRangePairs targetSheet, NAME_RANGE_PICKUPABLE_OBJECTS_ITEM_ID, _
                                                    NAME_RANGE_PICKUPABLE_OBJECTS_NAME, itemObjectsDict
                    modLists.CollectNamedRangePairs targetSheet, NAME_RANGE_MULTI_STATE_OBJECTS_STATE_ID, _
                                                    NAME_RANGE_MULTI_STATE_OBJECTS_OBJECT_NAME, stateObjectsDict
                    modLists.CollectNamedRangePairs targetSheet, NAME_RANGE_TOUCHABLE_OBJECTS_HOTSPOT_ID, _
                                                    NAME_RANGE_TOUCHABLE_OBJECTS_HOTSPOT_NAME, hotspotObjectsDict
                End If
            End If
        End If
    Next targetSheet
End Sub

' -----------------------------------------------------------------------------------
' Procedure : WriteSyncMode
' Purpose   : SYNC mode: Deletes all columns and rewrites everything.
' -----------------------------------------------------------------------------------
Private Sub WriteSyncMode(ByVal dataList As ListObject, _
                         ByVal roomsDict As Scripting.Dictionary, _
                         ByVal scenesDict As Scripting.Dictionary, _
                         ByVal itemObjectsDict As Scripting.Dictionary, _
                         ByVal stateObjectsDict As Scripting.Dictionary, _
                         ByVal hotspotObjectsDict As Scripting.Dictionary)
    
    ' clear all columns
    modLists.ClearTableColumn dataList, LISTS_HEADER_ROOM_ID
    modLists.ClearTableColumn dataList, LISTS_HEADER_ROOM_ALIAS
    modLists.ClearTableColumn dataList, LISTS_HEADER_SCENE_ID
    modLists.ClearTableColumn dataList, LISTS_HEADER_ITEM_ID
    modLists.ClearTableColumn dataList, LISTS_HEADER_STATE_OBJECT_ID
    modLists.ClearTableColumn dataList, LISTS_HEADER_HOTSPOT_ID
    modLists.ClearTableColumn dataList, LISTS_HEADER_ITEM_NAME
    modLists.ClearTableColumn dataList, LISTS_HEADER_STATE_OBJECT_NAME
    modLists.ClearTableColumn dataList, LISTS_HEADER_HOTSPOT_NAME
    
    ' Rewrite all data
    modLists.WriteDictionaryToTableColumns dataList, LISTS_HEADER_ROOM_ID, roomsDict, LISTS_HEADER_ROOM_ALIAS
    modLists.WriteDictionaryToTableColumns dataList, LISTS_HEADER_SCENE_ID, scenesDict
    modLists.WriteDictionaryToTableColumns dataList, LISTS_HEADER_ITEM_ID, itemObjectsDict, LISTS_HEADER_ITEM_NAME
    modLists.WriteDictionaryToTableColumns dataList, LISTS_HEADER_STATE_OBJECT_ID, stateObjectsDict, LISTS_HEADER_STATE_OBJECT_NAME
    modLists.WriteDictionaryToTableColumns dataList, LISTS_HEADER_HOTSPOT_ID, hotspotObjectsDict, LISTS_HEADER_HOTSPOT_NAME
End Sub

' -----------------------------------------------------------------------------------
' Procedure : WriteAppendMode
' Purpose   : APPEND mode: Only deletes and renews Room/Scene columns,
'             to others columns only adds new objects.
' -----------------------------------------------------------------------------------
Private Sub WriteAppendMode(ByVal dataList As ListObject, _
                           ByVal roomsDict As Scripting.Dictionary, _
                           ByVal scenesDict As Scripting.Dictionary, _
                           ByVal itemObjectsDict As Scripting.Dictionary, _
                           ByVal stateObjectsDict As Scripting.Dictionary, _
                           ByVal hotspotObjectsDict As Scripting.Dictionary, _
                           ByVal processObjects As Boolean)
    
    ' Always rewrite room/scene columns
    modLists.ClearTableColumn dataList, LISTS_HEADER_ROOM_ID
    modLists.ClearTableColumn dataList, LISTS_HEADER_ROOM_ALIAS
    modLists.ClearTableColumn dataList, LISTS_HEADER_SCENE_ID
    
    modLists.WriteDictionaryToTableColumns dataList, LISTS_HEADER_ROOM_ID, roomsDict, LISTS_HEADER_ROOM_ALIAS
    modLists.WriteDictionaryToTableColumns dataList, LISTS_HEADER_SCENE_ID, scenesDict
    
    ' Only process object data if desired
    If processObjects Then
        'Collect existing object pairs from table
        Dim existingItemPairsDict As Scripting.Dictionary
        Dim existingStatePairsDict As Scripting.Dictionary
        Dim existingHotspotPairsDict As Scripting.Dictionary
        Dim existingKeysDict As Scripting.Dictionary
        
        Set existingItemPairsDict = New Scripting.Dictionary
        Set existingStatePairsDict = New Scripting.Dictionary
        Set existingHotspotPairsDict = New Scripting.Dictionary
        Set existingKeysDict = New Scripting.Dictionary
        
        modLists.CollectTableColumnPairs dataList, LISTS_HEADER_ITEM_ID, LISTS_HEADER_ITEM_NAME, existingItemPairsDict
        modLists.CollectTableColumnPairs dataList, LISTS_HEADER_STATE_OBJECT_ID, LISTS_HEADER_STATE_OBJECT_NAME, existingStatePairsDict
        modLists.CollectTableColumnPairs dataList, LISTS_HEADER_HOTSPOT_ID, LISTS_HEADER_HOTSPOT_NAME, existingHotspotPairsDict
        
        ' Add only missing objects
        modLists.AppendMissingDictSetToTableColumns dataList, LISTS_HEADER_ITEM_ID, existingKeysDict, _
                                                    itemObjectsDict, LISTS_HEADER_ITEM_NAME
        modLists.AppendMissingDictSetToTableColumns dataList, LISTS_HEADER_STATE_OBJECT_ID, existingKeysDict, _
                                                    stateObjectsDict, LISTS_HEADER_STATE_OBJECT_NAME
        modLists.AppendMissingDictSetToTableColumns dataList, LISTS_HEADER_HOTSPOT_ID, existingKeysDict, _
                                                    hotspotObjectsDict, LISTS_HEADER_HOTSPOT_NAME
    End If
End Sub

' -----------------------------------------------------------------------------------
' Procedure : UpdateDispatcherTable
' Purpose   : Updates Room ID and Room Alias in the DropDownLists table.
' -----------------------------------------------------------------------------------
Private Sub UpdateDispatcherTable(ByVal targetBook As Workbook, _
                                 ByVal oldRoomID As String, _
                                 ByVal oldRoomAlias As String, _
                                 ByVal newRoomID As String, _
                                 ByVal newRoomAlias As String)
    On Error GoTo ErrHandler
    
    Dim dispatcherSheet As Worksheet
    Dim dataTable As ListObject
    Dim roomIDCol As ListColumn
    Dim roomAliasCol As ListColumn
    Dim cell As Range
    
    For Each dispatcherSheet In targetBook.Worksheets
        If modTags.HasSheetTag(dispatcherSheet, SHEET_DISPATCHER) Then
            Exit For
        End If
    Next dispatcherSheet
    
    If dispatcherSheet Is Nothing Then Exit Sub
    
    On Error Resume Next
    Set dataTable = dispatcherSheet.ListObjects(NAME_DATA_TABLE)
    On Error GoTo ErrHandler
    
    If dataTable Is Nothing Then Exit Sub
    
    On Error Resume Next
    Set roomIDCol = dataTable.ListColumns(LISTS_HEADER_ROOM_ID)
    On Error GoTo ErrHandler
    
    If Not roomIDCol Is Nothing Then
        If Not roomIDCol.DataBodyRange Is Nothing Then
            For Each cell In roomIDCol.DataBodyRange.Cells
                If Not IsEmpty(cell.value) Then
                    If StrComp(CStr(cell.value), oldRoomID, vbTextCompare) = 0 Then
                        cell.value = newRoomID
                    End If
                End If
            Next cell
        End If
    End If
    
    On Error Resume Next
    Set roomAliasCol = dataTable.ListColumns(LISTS_HEADER_ROOM_ALIAS)
    On Error GoTo ErrHandler
    
    If Not roomAliasCol Is Nothing Then
        If Not roomAliasCol.DataBodyRange Is Nothing Then
            For Each cell In roomAliasCol.DataBodyRange.Cells
                If Not IsEmpty(cell.value) Then
                    If StrComp(CStr(cell.value), oldRoomAlias, vbTextCompare) = 0 Then
                        cell.value = newRoomAlias
                    End If
                End If
            Next cell
        End If
    End If
    
    Exit Sub
    
ErrHandler:
    modErr.ReportError "modMain.UpdateDispatcherTable", Err.Number, Erl, caption:=AppProjectName
End Sub

