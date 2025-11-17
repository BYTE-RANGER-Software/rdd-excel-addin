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
'   - HasRoomSheet            : Determine whether a room with specific ID exists.
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
        newRoomSheet.Name = "DO_NOT_DELETE"
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
        UpdateLists targetBook
    End If
    
CleanExit:
    modUtil.HideOpMode False
    Exit Function
    
ErrHandler:
    modErr.ReportError "AddRoom", Err.Number, Erl, caption:=modMain.AppProjectName
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
'   - Collects Room sheets via modSheets.BuildDictFromSheetsByTag (excludes target sheet).
'   - Checks references via GetAllSheetNamesUsingRoomID.
'   - Uses HideOpMode during deletion; hands back referencing sheets if found.
' -----------------------------------------------------------------------------------
Public Function RemoveRoom(ByVal targetSheet As Worksheet, _
    Optional ByVal updateAggregations As Boolean = True, _
    Optional ByRef returnRefSheets As Collection) As Boolean
    
    On Error GoTo ErrHandler

    Dim parentBook As Workbook: Set parentBook = targetSheet.Parent
    Dim roomId As String

    If targetSheet Is Nothing Then
        Err.Raise ERR_ARG_NULL_TARGETSHEET, "modRooms.RemoveRoom", "Argument 'targetSheet' must not be Nothing."
    End If
    
    If Not IsRoomSheet(targetSheet, roomId) Then
        Err.Raise ERR_NOT_A_ROOM_SHEET, "modRooms.RemoveRoom", "The provided sheet is not a Room sheet."
    End If
    
    ' Collect all Room-sheets except the active one
    Dim roomsDict As Scripting.Dictionary
    Set roomsDict = modSheets.BuildDictFromSheetsByTag(parentBook, ROOM_SHEET_ID_TAG_NAME, sheetToExclude:=targetSheet)
    
    ' Check references to the active room sheet in all other room sheets
    Dim usedByCol As Collection
    Set usedByCol = GetAllSheetNamesUsingRoomID(roomId, roomsDict)
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
        UpdateLists targetSheet.Parent
    End If
    
CleanExit:
    modUtil.HideOpMode False
    Exit Function
    
ErrHandler:
    modErr.ReportError "RemoveRoom", Err.Number, Erl, caption:=modMain.AppProjectName
    Resume CleanExit
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
' Function  : HasRoomSheet
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
Public Function HasRoomSheet(ByVal targetBook As Workbook, ByVal roomId As String, Optional ByRef returnSheet As Worksheet = Nothing) As Boolean
    Dim sheet As Worksheet
    Dim tagValue As String
    Dim isFound As Boolean
    
    On Error GoTo ErrHandler
    
    For Each sheet In targetBook.Worksheets
        If modTags.HasSheetTag(sheet, ROOM_SHEET_ID_TAG_NAME, tagValue) Then
            If StrComp(roomId, tagValue, vbBinaryCompare) = 0 Then
                Set returnSheet = sheet
                isFound = True
                GoTo CleanExit
            End If
        End If
    Next sheet
    
CleanExit:
    HasRoomSheet = isFound
    Exit Function

ErrHandler:
    modErr.ReportError "HasRoomSheet", Err.Number, Erl, caption:=modMain.AppProjectName
    Resume CleanExit
End Function

' -----------------------------------------------------------------------------------
' Function  : UpdateLists
' Purpose   : Aggregates Room IDs, Objects, and Scene IDs from all Room sheets into
'             the Lists sheet and updates the corresponding named ranges.
'
' Parameters:
'   targetBook [Workbook]     - (optional) Target Workbook that contains data lists, default ActiveWorkbook
' Returns   : (none)
'
' Notes     :
'   - Room ID via IsRoomSheet (tag), Room Alias via NAME_CELL_ROOM_ALIAS.
'   - Collects object names via named ranges (pickupable/multistate/touchable).
' -----------------------------------------------------------------------------------
Public Sub UpdateLists(Optional targetBook As Workbook)
    On Error GoTo ErrHandler
    Dim roomsDict As Scripting.Dictionary: Set roomsDict = New Scripting.Dictionary
    Dim objectsDict As Scripting.Dictionary:  Set objectsDict = New Scripting.Dictionary
    Dim scenesDict As Scripting.Dictionary: Set scenesDict = New Scripting.Dictionary
    
    Dim existingKeysDict As Scripting.Dictionary: Set existingKeysDict = New Scripting.Dictionary
    Dim targetSheet As Worksheet
    Dim roomId As String
    Dim dataList As ListObject
    
    If targetBook Is Nothing Then Set targetBook = ActiveWorkbook
    
    For Each targetSheet In targetBook.Worksheets
        If modRooms.IsRoomSheet(targetSheet, roomId) Then
            ' Room ID
            If Len(roomId) = 0 Then GoTo SkipWksIteration
            If Len(roomId) > 0 Then roomsDict(roomId) = True
            
            ' Room Alias
            Dim roomAlias As String: roomAlias = modLists.GetNamedOrHeaderValue(targetSheet, NAME_CELL_ROOM_ALIAS, Array("Room Alias", NAME_CELL_ROOM_ALIAS))
            If Len(roomId) > 0 Then roomsDict(roomId) = roomAlias
            
            ' Scene ID
            Dim sceneId As String: sceneId = modLists.GetNamedOrHeaderValue(targetSheet, NAME_CELL_SCENE_ID, Array("Scene ID", NAME_CELL_SCENE_ID))
            If Len(sceneId) > 0 Then scenesDict(sceneId) = True
            
            modLists.CollectNamedRangeValues targetSheet, NAME_RANGE_PICKUPABLE_OBJECTS_ITEM_ID, objectsDict
            modLists.CollectNamedRangeValues targetSheet, NAME_RANGE_MULTI_STATE_OBJECTS_STATE_ID, objectsDict
            modLists.CollectNamedRangeValues targetSheet, NAME_RANGE_TOUCHABLE_OBJECTS_HOTSPOT_ID, objectsDict
            
        End If
SkipWksIteration:
    Next targetSheet
    
    Dim listsSheet As Worksheet: Set listsSheet = modSheets.GetSheetByCodeName(SHEET_DISPATCHER)
    
    If listsSheet Is Nothing Then: Set listsSheet = modTags.GetSheetByTag(targetBook, SHEET_DISPATCHER)
    
    If Not listsSheet Is Nothing Then
        
        ' Room IDs
        Set dataList = listsSheet.ListObjects(NAME_DATA_TABLE)
        modLists.ClearTableColumn dataList, LISTS_HEADER_ROOM_ID
        modLists.ClearTableColumn dataList, LISTS_HEADER_ROOM_ALIAS
        
        ' Write Room IDs & Room Alias sorted, must always be rewritten, as it is related to the room pages
        modLists.WriteDictSetToTableColumn dataList, LISTS_HEADER_ROOM_ID, roomsDict, True
           
        ' Append only missing Object names
        modLists.CollectTableColumnValues dataList, LISTS_HEADER_OBJECTS, existingKeysDict
        modLists.AppendMissingDictKeysToTableColumn dataList, LISTS_HEADER_OBJECTS, existingKeysDict, objectsDict
        
        ' Append only missing Scene IDs
        modLists.CollectTableColumnValues dataList, LISTS_HEADER_SCENE_ID, existingKeysDict
        modLists.AppendMissingDictKeysToTableColumn dataList, LISTS_HEADER_SCENE_ID, existingKeysDict, scenesDict

    End If
    Exit Sub
    
ErrHandler:
    modErr.ReportError "UpdateLists", Err.Number, Erl, caption:=modMain.AppProjectName
End Sub

' -----------------------------------------------------------------------------------
' Function  : SyncLists
' Purpose   : Aggregates Room IDs, Objects, and Scene IDs from all Room sheets and
'             writes them into the Lists sheet. Clears the three target columns,
'             writes headers, outputs sorted values, and updates named ranges.
'
' Parameters: (none)
'
' Returns   : (none)
'
' Notes     :
'   - Scans Room sheets via IsRoomSheet (not by name prefix).
'   - Room Alias is written alongside Room ID (dictionary value).
' -----------------------------------------------------------------------------------
Public Sub SyncLists()
    On Error GoTo ErrHandler
    Dim roomsDict As Scripting.Dictionary:   Set roomsDict = New Scripting.Dictionary
    Dim objectsDict As Scripting.Dictionary: Set objectsDict = New Scripting.Dictionary
    Dim scenesDict As Scripting.Dictionary:  Set scenesDict = New Scripting.Dictionary
    
    Dim targetSheet As Worksheet
    Dim activeBook As Workbook
    Dim roomId As String
    
    Set activeBook = ActiveWorkbook
    ' collect datas
    For Each targetSheet In activeBook.Worksheets
        If modRooms.IsRoomSheet(targetSheet, roomId) Then
            If Len(roomId) = 0 Then GoTo SkipWksIteration
            If Len(roomId) > 0 Then roomsDict(roomId) = True
            
            ' Room Alias
            Dim roomAlias As String: roomAlias = modLists.GetNamedOrHeaderValue(targetSheet, NAME_CELL_ROOM_ALIAS, Array("Room Alias", NAME_CELL_ROOM_ALIAS))
            If Len(roomId) > 0 Then roomsDict(roomId) = roomAlias
            
            Dim sceneId As String
            sceneId = modLists.GetNamedOrHeaderValue(targetSheet, NAME_CELL_SCENE_ID, Array("Scene ID", NAME_CELL_SCENE_ID))
            If Len(sceneId) > 0 Then scenesDict(sceneId) = True
            
            modLists.CollectNamedRangeValues targetSheet, NAME_RANGE_PICKUPABLE_OBJECTS_ITEM_ID, objectsDict
            modLists.CollectNamedRangeValues targetSheet, NAME_RANGE_MULTI_STATE_OBJECTS_STATE_ID, objectsDict
            modLists.CollectNamedRangeValues targetSheet, NAME_RANGE_TOUCHABLE_OBJECTS_HOTSPOT_ID, objectsDict
        End If
SkipWksIteration:
    Next targetSheet
    
    Dim listsSheet As Worksheet: Set listsSheet = modSheets.GetSheetByCodeName(SHEET_DISPATCHER)
    
    If listsSheet Is Nothing Then: Set listsSheet = modTags.GetSheetByTag(activeBook, SHEET_DISPATCHER)
    
    If Not listsSheet Is Nothing Then
        ' Clear target columns
        listsSheet.Columns(LISTS_COL_ROOM_ID).Clear     ' Room IDs
        listsSheet.Columns(LISTS_COL_ROOM_ALIAS).Clear
        listsSheet.Columns(LISTS_COL_SCENE_ID).Clear    ' Scene IDs
        listsSheet.Columns(LISTS_COL_OBJECTS).Clear     ' Objects
        
        ' Headers
        listsSheet.Cells(1, LISTS_COL_ROOM_ID).value = LISTS_HEADER_ROOM_ID
        listsSheet.Cells(1, LISTS_COL_ROOM_ALIAS).value = LISTS_HEADER_ROOM_ALIAS
        listsSheet.Cells(1, LISTS_COL_OBJECTS).value = LISTS_HEADER_OBJECTS
        listsSheet.Cells(1, LISTS_COL_SCENE_ID).value = LISTS_HEADER_SCENE_ID
        listsSheet.Range("A1:ZZ1").Font.Bold = True
    
        ' Write sorted values
        modLists.WriteDictSetToColumn listsSheet, roomsDict, 2, LISTS_COL_ROOM_ID, True
        modLists.WriteDictSetToColumn listsSheet, objectsDict, 2, LISTS_COL_OBJECTS
        modLists.WriteDictSetToColumn listsSheet, scenesDict, 2, LISTS_COL_SCENE_ID
        
        ' Update named ranges
        modLists.UpdateNamedListRange NAME_LIST_ROOM_IDS, listsSheet, LISTS_COL_ROOM_ID
        modLists.UpdateNamedListRange NAME_LIST_OBJECTS, listsSheet, LISTS_COL_OBJECTS
        modLists.UpdateNamedListRange NAME_LIST_SCENE_IDS, listsSheet, LISTS_COL_SCENE_ID
    End If
    
    Exit Sub
    
ErrHandler:
    modErr.ReportError "SyncLists", Err.Number, Erl, caption:=modMain.AppProjectName
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
        
    ' Add data validations
    
    ' Type
    Set dataRange = targetSheet.Range(NAME_RANGE_PUZZLES_ACTION)
    ApplyListValidation dataRange, NAME_LIST_PUZZLE_TYPES, "Type", "Choose a type from the list."

    
    
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
'   roomSheetsDict   [Object]  - Dictionary of Room worksheets to inspect (values = Worksheet).
'
' Returns   : Collection - Sheet names that reference roomId.
'
' Notes     :
' -----------------------------------------------------------------------------------
Private Function GetAllSheetNamesUsingRoomID(ByVal roomId As String, ByVal roomSheetsDict As Scripting.Dictionary) As Collection
    ' Returns a collection of sheet names that reference roomId
    Dim col As New Collection
    Dim dictKey As Variant, targetSheet As Worksheet
    Dim cellRange As Range
    
    For Each dictKey In roomSheetsDict.Keys
        Set targetSheet = roomSheetsDict(dictKey)
        
        ' in Framed area "DOORS TO..."
        Set cellRange = targetSheet.Range(NAME_RANGE_DOORS_TO_ROOM_ID) 'FindFramedRangeByHeading(targetSheet, "DOORS TO", False)
        If Not cellRange Is Nothing Then
            If modRanges.RangeHasValue(cellRange, roomId, True, False) Then
                col.Add targetSheet.Name
                GoTo NextSheet
            End If
        End If
                
NextSheet:
        Set cellRange = Nothing
    Next
    
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
