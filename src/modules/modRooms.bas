Attribute VB_Name = "modRooms"
' -----------------------------------------------------------------------------------
' Module    : modRooms
' Purpose   : Create, initialize and manage "Room" worksheets; aggregate list data.
'
' Public API:
'   - ProtectAllRoomSheets    :
'   - AddRoom                 : Create a new room sheet from template and initialize it.
'   - IsRoomSheet             : Check whether a sheet is a room sheet (by tag).
'   - RemoveRoom              : Delete a room sheet after reference checks.
'   - GetNextRoomIndex        : Compute next numeric room index.
'   - HasRoomID               : Determine whether a room with specific ID exists.
'   - UpdateRoomdataLists
'   - UpdateScenesLists
'   - UpdateItemsLists
'   - UpdateStateObjectsLists
'   - UpdateHotspotsLists
'   - UpdateGeneralSettingsLists
'   - UpdateActorsLists
'   - UpdateSoundsLists
'   - UpdateSpecialFXLists
'   - UpdateFlagsLists
'   - UpdateAllLists
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
End Enum

' ===== Public API ==================================================================

' -----------------------------------------------------------------------------------
' Procedure : ProtectAllRoomSheets
' Purpose   : Protects all Room sheets with UserInterfaceOnly:=True to allow VBA
'             modifications while keeping user restrictions active.
'
' Parameters:
'   targetBook   [Workbook] - Optional; target workbook (default: ActiveWorkbook)
'
' Returns   : (none)
'
' Notes     :
'   - Must be called on Workbook_Open as UserInterfaceOnly is session-specific.
'   - Unprotects first to ensure clean state, then applies UserInterfaceOnly protection.
'   - Only affects sheets identified by modRooms.IsRoomSheet().
' -----------------------------------------------------------------------------------
Public Function ProtectAllRoomSheets(Optional targetBook As Workbook = Nothing) As Long
    On Error GoTo ErrHandler
    
    If targetBook Is Nothing Then Set targetBook = ActiveWorkbook
    
    Dim roomSheet As Worksheet
    Dim protectedCount As Long
    
    For Each roomSheet In targetBook.Worksheets
        If modRooms.IsRoomSheet(roomSheet) Then
            roomSheet.Protect password:=ROOM_SHEET_PWD, UserInterfaceOnly:=True
            protectedCount = protectedCount + 1
        End If
    Next roomSheet
    
    ProtectAllRoomSheets = protectedCount
CleanExit:
    Exit Function
    
ErrHandler:
    modErr.ReportError "modRooms.ProtectAllRoomSheets", Err.Number, Erl, caption:=modMain.AppProjectName
End Function

' -----------------------------------------------------------------------------------
' Function  : AddRoom
' Purpose   : Clones required templates, ensures helper sheets exist, creates a new
'             Room sheet with the next sequential ID, initializes it, and activates it.
'
' Parameters:
'   targetBook            [Workbook] - Target workbook.
'   roomName              [String]   - Name for the new Room sheet.
'   roomIdx               [Long]     - Numeric index used for the Room ID formatting.
'   roomNo                [Long]     - AGS room number
'   sceneID               [String]   - Scene identifier/name (optional, can be empty)
'   updateAggregations    [Boolean]  - Optional. If True, UpdateLists is called after creation.
'
' Returns   : [Worksheet] - The newly created Room worksheet.
'
' Notes     :
'   - Ensures dispatcher and lists sheets by copying them from RDDAddInWkBk if missing.
'   - Creates a new sheet from SHEET_ROOM_TEMPLATE and tags it with ROOM_SHEET_ID_TAG_NAME.
'   - Calls SetupRoom to wire controls/values; toggles HideOpMode during operations.
'   - Scene ID can be shared across multiple rooms (no uniqueness validation)
' -----------------------------------------------------------------------------------
Public Function AddRoom(targetBook As Workbook, _
    roomName As String, _
    roomIdx As Long, _
    roomNo As Long, _
    Optional sceneID As String = "", _
    Optional ByVal updateAggregations As Boolean = True) As Worksheet
                        
    On Error GoTo ErrHandler
    
    If Len(Trim$(roomName)) = 0 Then
        Err.Raise ERR_INVALID_ROOM_NAME, "AddRoom", "Room name cannot be empty"
    End If

    Dim tmplSheet As Worksheet
    Dim newRoomSheet As Worksheet
    
    modUtil.HideOpMode True
            
    If Not modSheets.SheetCodeNameExists(modConst.SHEET_DISPATCHER, targetBook) And Not modTags.SheetWithTagExists(targetBook, SHEET_DISPATCHER) Then
    
        Set tmplSheet = RDDAddInWkBk.Worksheets(modConst.SHEET_DISPATCHER)
        tmplSheet.Visible = xlSheetVisible
        
        tmplSheet.Copy After:=targetBook.Sheets(targetBook.Sheets.count)
        Set newRoomSheet = ActiveSheet
        
        newRoomSheet.Visible = xlSheetHidden
        newRoomSheet.name = SHEET_DISPATCHER
        modProps.ClearAllCustomProperties newRoomSheet
        modTags.TagSheet newRoomSheet, SHEET_DISPATCHER
        
        Set tmplSheet = Nothing
        Set newRoomSheet = Nothing
    End If
            
    Set tmplSheet = RDDAddInWkBk.Worksheets(modConst.SHEET_ROOM_TEMPLATE)
    tmplSheet.Visible = xlSheetVisible
    
    tmplSheet.Copy After:=targetBook.Sheets(targetBook.Sheets.count)
    Set newRoomSheet = ActiveSheet
    
    newRoomSheet.Protect password:=ROOM_SHEET_PWD, UserInterfaceOnly:=True
    newRoomSheet.name = modSheets.GetValidUniqueSheetName(roomName, targetBook)
    
    modProps.ClearAllCustomProperties newRoomSheet
    modTags.TagSheet newRoomSheet, ROOM_SHEET_ID_TAG_NAME, GetFormattedRoomID(roomIdx)
    
    SetupRoom newRoomSheet, roomName, roomIdx, roomNo, sceneID
    
    Set AddRoom = newRoomSheet
        
    If updateAggregations Then
        UpdateRoomsMetadataLists targetBook, newRoomSheet
        UpdateScenesMetadataLists targetBook, newRoomSheet
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
        If usedByCol.count > 0 Then
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
        SynchronizeAllLists parentBook
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
    
    Call UpdateRoomMetadataInDispatcherTable(targetBook, oldRoomID, oldRoomAlias, newRoomID, newRoomAlias)
    
    UpdateRoomReferences = updatedCount
    
    Exit Function
    
ErrHandler:
    modErr.ReportError "modRooms.UpdateRoomReferences", Err.Number, Erl, caption:=AppProjectName
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
            numIdx = Val(Mid$(CStr(tagValue), Len(ROOM_SHEET_ID_PREFIX) + 1))
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
    
    isFound = False
    Set returnSheet = Nothing
    
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
    Dim isFound As Boolean
    Dim cell As Range
    
    On Error GoTo ErrHandler
    
    isFound = False
    Set returnSheet = Nothing
    
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
' Function  : HasRoomNo
' Purpose   : Determine whether the workbook contains a Room sheet with the given AGS Number.
'
' Parameters:
'   targetBook       [Workbook]              - Workbook to scan.
'   roomNo           [Long]                  - Room No to search for (e.g., 100).
'   returnSheet      [Worksheet]             - (Optional ByRef) Receives the first matching sheet if found.
'
' Returns   : Boolean - True if found; otherwise False.
'
' Notes     :
' -----------------------------------------------------------------------------------
Public Function HasRoomNo(ByVal targetBook As Workbook, ByVal roomNo As Long, Optional ByRef returnSheet As Worksheet = Nothing) As Boolean
    Dim sheet As Worksheet
    Dim isFound As Boolean
    Dim cell As Range
    Dim cellValue As Variant
    
    On Error GoTo ErrHandler
    
    isFound = False
    Set returnSheet = Nothing
    
    For Each sheet In targetBook.Worksheets
        If modTags.HasSheetTag(sheet, ROOM_SHEET_ID_TAG_NAME) Then
            On Error Resume Next
            Set cell = sheet.Range(modConst.NAME_CELL_ROOM_NO)
            On Error GoTo ErrHandler
            
            If Not cell Is Nothing Then
                cellValue = cell.value
                If IsNumeric(cellValue) Then
                    If CLng(cellValue) = roomNo Then
                        Set returnSheet = sheet
                        isFound = True
                        GoTo CleanExit
                    End If
                End If
            End If
        End If
    Next sheet
    
CleanExit:
    HasRoomNo = isFound
    Exit Function

ErrHandler:
    modErr.ReportError "HasRoomNo", Err.Number, Erl, caption:=modMain.AppProjectName
    Resume CleanExit
End Function

' -----------------------------------------------------------------------------------
' Procedure : SynchronizeAllLists
' Purpose   : synchronize all list (complete rewrite).
'
' Parameters:
'   targetBook  [Workbook]        - Target Workbook
'
' Notes     :
' -----------------------------------------------------------------------------------
Public Sub SynchronizeAllLists(targetBook As Workbook)
    On Error GoTo ErrHandler
    
    ' Enable silent mode
    modUtil.HideOpMode True, affectScreen:=False, affectEvents:=False
    
    
    UpdateRoomsMetadataLists targetBook
    UpdateScenesMetadataLists targetBook
    UpdateGeneralSettingsLists targetBook
    UpdateActorsLists targetBook
    UpdateSoundsLists targetBook
    UpdateSpecialFXLists targetBook
    UpdateFlagsLists targetBook
    UpdateItemsLists targetBook
    UpdateStateObjectsLists targetBook
    UpdateHotspotsLists targetBook

CleanExit:
    modUtil.HideOpMode False
    Exit Sub
    
ErrHandler:
    modUtil.HideOpMode False
    modErr.ReportError "ProcessListsUpdate", Err.Number, Erl, caption:=modMain.AppProjectName
    Resume CleanExit
End Sub

' -----------------------------------------------------------------------------------
' Procedure : UpdateRoomsMetadataLists
' Purpose   : Updates Scene ID, Room ID, Room No, and Room Alias lists.
'
' Parameters:
'   targetBook    [Workbook]          - Target workbook
'   dataSrcSheet   [Worksheet]         - (Optional) If Sheet is provided, only this sheet is processed
'   mode          [ListUpdateMode]    - LUM_Sync = synchronise all Lists (rewrite) or LUM_Append = appen only new Data
'
' Returns   : (none)
'
' Notes     :
'   - DEFAULT: LUM_Sync (Room metadata are critical, should always be rewrite)
' -----------------------------------------------------------------------------------
Public Sub UpdateRoomsMetadataLists(ByVal targetBook As Workbook, _
    Optional ByVal dataSrcSheet As Worksheet = Nothing, _
    Optional ByVal mode As ListUpdateMode = LUM_Sync)
    On Error GoTo ErrHandler
    
    modUtil.HideOpMode True, affectScreen:=False, affectEvents:=False
    
    Dim listsSheet As Worksheet
    Dim dataList As ListObject
    If Not GetListsSheetAndTable(targetBook, listsSheet, dataList) Then Exit Sub
    
    Dim roomsDict As Scripting.Dictionary
    Set roomsDict = New Scripting.Dictionary
    
    ' Collect room metadata
    If dataSrcSheet Is Nothing Then
        'Collect from ALL sheets
        CollectRoomMetadata targetBook, roomsDict
    Else
        ' Collect only from dataSrcSheet
        CollectRoomMetadata targetBook, roomsDict, dataSrcSheet
    End If
    
    ' Write to table based on mode
    If mode = LUM_Sync Then
        ' SYNC: Clear + Rewrite (full refresh)
        
        modLists.ClearTableColumn dataList, LISTS_HEADER_ROOM_ID
        modLists.ClearTableColumn dataList, LISTS_HEADER_ROOM_NO
        modLists.ClearTableColumn dataList, LISTS_HEADER_ROOM_ALIAS
                
        modLists.WriteDictToTableColumns dataList, LISTS_HEADER_ROOM_ID, roomsDict, _
            LISTS_HEADER_ROOM_NO, LISTS_HEADER_ROOM_ALIAS
                
    Else
        ' APPEND: Add only new entries (Room IDs are keys, rarely use APPEND for rooms)
        Dim existingKeysDict As Scripting.Dictionary
        Set existingKeysDict = New Scripting.Dictionary
        
        ' Collect existing room IDs
        modLists.CollectTableColumnsToDict dataList, LISTS_HEADER_ROOM_ID, existingKeysDict, _
            LISTS_HEADER_ROOM_NO, LISTS_HEADER_ROOM_ALIAS
        
        ' Append missing entries
        modLists.AppendMissingDictSetToTableColumns dataList, LISTS_HEADER_ROOM_ID, _
            existingKeysDict, roomsDict, LISTS_HEADER_ROOM_NO, LISTS_HEADER_ROOM_ALIAS

    End If
    
    Set roomsDict = Nothing
    modUtil.HideOpMode False
    Exit Sub
    
ErrHandler:
    modUtil.HideOpMode False
    modErr.ReportError "modRooms.UpdateRoomMetadata", Err.Number, Erl, caption:=modMain.AppProjectName
End Sub

' -----------------------------------------------------------------------------------
' Procedure : UpdateScenesMetadataLists
' Purpose   : Updates Scene ID list.
'
' Parameters:
'   targetBook    [Workbook]          - Target workbook
'   dataSrcSheet   [Worksheet]         - (Optional) If Sheet is provided, only this sheet is processed
'   mode          [ListUpdateMode]    - LUM_Sync = synchronise all Lists (rewrite) or LUM_Append = appen only new Data
'
' Notes     :
' -----------------------------------------------------------------------------------
Public Sub UpdateScenesMetadataLists(ByVal targetBook As Workbook, _
    Optional ByVal dataSrcSheet As Worksheet = Nothing, _
    Optional ByVal mode As ListUpdateMode = LUM_Sync)
    On Error GoTo ErrHandler
    
    modUtil.HideOpMode True, affectScreen:=False, affectEvents:=False
    
    Dim listsSheet As Worksheet
    Dim dataList As ListObject
    If Not GetListsSheetAndTable(targetBook, listsSheet, dataList) Then Exit Sub
    
    Dim scenesDict As Scripting.Dictionary
    Set scenesDict = New Scripting.Dictionary
    
    ' Collect scene IDs
    If dataSrcSheet Is Nothing Then
        CollectSceneIDs targetBook, scenesDict
    Else
        CollectSceneIDs targetBook, scenesDict, dataSrcSheet
    End If
    
    ' Write based on mode
    If mode = LUM_Sync Then
        ' SYNC: Clear + Rewrite
        modLists.ClearTableColumn dataList, LISTS_HEADER_SCENE_ID
        modLists.WriteDictToTableColumns dataList, LISTS_HEADER_SCENE_ID, scenesDict
    Else
        ' APPEND: Add only new entries
        Dim existingKeysDict As Scripting.Dictionary
        Set existingKeysDict = New Scripting.Dictionary
        
        modLists.CollectTableColumnsToDict dataList, LISTS_HEADER_SCENE_ID, existingKeysDict
        modLists.AppendMissingDictSetToTableColumns dataList, LISTS_HEADER_SCENE_ID, _
            existingKeysDict, scenesDict
    End If
    
    modUtil.HideOpMode False
    Exit Sub
    
ErrHandler:
    modUtil.HideOpMode False
    modErr.ReportError "modRooms.UpdateScenesList", Err.Number, Erl, caption:=modMain.AppProjectName
End Sub

' -----------------------------------------------------------------------------------
' Procedure : UpdateGeneralSettingsLists
' Purpose   : Updates unique dimension values (Width, Height, UI Height) from rooms.
'
' Parameters:
'   targetBook    [Workbook]          - Target workbook
'   dataSrcSheet   [Worksheet]         - (Optional) If Sheet is provided, only this sheet is processed
'   mode          [ListUpdateMode]    - LUM_Sync = synchronise all Lists (rewrite) or LUM_Append = appen only new Data
'
' Notes     :
' -----------------------------------------------------------------------------------
Public Sub UpdateGeneralSettingsLists(ByVal targetBook As Workbook, _
    Optional ByVal dataSrcSheet As Worksheet = Nothing, _
    Optional ByVal mode As ListUpdateMode = LUM_Sync)
    On Error GoTo ErrHandler
    
    modUtil.HideOpMode True, affectScreen:=False, affectEvents:=False
    
    Dim listsSheet As Worksheet
    Dim dataList As ListObject
    If Not GetListsSheetAndTable(targetBook, listsSheet, dataList) Then Exit Sub
    
    Dim widthDict As Scripting.Dictionary
    Dim heightDict As Scripting.Dictionary
    Dim uiHeightDict As Scripting.Dictionary
    
    Set widthDict = New Scripting.Dictionary
    Set heightDict = New Scripting.Dictionary
    Set uiHeightDict = New Scripting.Dictionary
    
    ' Collect General Settings
    If dataSrcSheet Is Nothing Then
        CollectGeneralSettings targetBook, widthDict, heightDict, uiHeightDict
    Else
        CollectGeneralSettings targetBook, widthDict, heightDict, uiHeightDict, dataSrcSheet
    End If
    
    ' Write based on mode
    If mode = LUM_Sync Then
        ' SYNC: Clear + Rewrite
        
        modLists.ClearTableColumn dataList, LISTS_HEADER_WIDTH
        modLists.ClearTableColumn dataList, LISTS_HEADER_HEIGHT
        modLists.ClearTableColumn dataList, LISTS_HEADER_UI_HEIGHT
    
        modLists.WriteDictToTableColumns dataList, LISTS_HEADER_WIDTH, widthDict
        modLists.WriteDictToTableColumns dataList, LISTS_HEADER_HEIGHT, heightDict
        modLists.WriteDictToTableColumns dataList, LISTS_HEADER_UI_HEIGHT, uiHeightDict
    Else
        ' APPEND: Add only new entries
        Dim existingWidthKeysDict As Scripting.Dictionary
        Set existingWidthKeysDict = New Scripting.Dictionary
        Dim existingHeightKeysDict As Scripting.Dictionary
        Set existingHeightKeysDict = New Scripting.Dictionary
        Dim existingUIHeightKeysDict As Scripting.Dictionary
        Set existingUIHeightKeysDict = New Scripting.Dictionary
        
        modLists.CollectTableColumnsToDict dataList, LISTS_HEADER_WIDTH, existingWidthKeysDict
        modLists.CollectTableColumnsToDict dataList, LISTS_HEADER_HEIGHT, existingHeightKeysDict
        modLists.CollectTableColumnsToDict dataList, LISTS_HEADER_UI_HEIGHT, existingUIHeightKeysDict
        
        modLists.AppendMissingDictSetToTableColumns dataList, LISTS_HEADER_WIDTH, _
            existingWidthKeysDict, widthDict
        modLists.AppendMissingDictSetToTableColumns dataList, LISTS_HEADER_HEIGHT, _
            existingHeightKeysDict, heightDict
        modLists.AppendMissingDictSetToTableColumns dataList, LISTS_HEADER_UI_HEIGHT, _
            existingUIHeightKeysDict, uiHeightDict
    End If

    modUtil.HideOpMode False
    Exit Sub
    
ErrHandler:
    modUtil.HideOpMode False
    modErr.ReportError "modRooms.UpdateGeneralSettingsLists", Err.Number, Erl, caption:=modMain.AppProjectName
End Sub

' -----------------------------------------------------------------------------------
' Procedure : UpdateActorsLists
' Purpose   : Updates Actor ID + Name lists.
'
' Parameters:
'   targetBook    [Workbook]          - Target workbook
'   dataSrcSheet   [Worksheet]         - (Optional) If Sheet is provided, only this sheet is processed
'   mode          [ListUpdateMode]    - LUM_Sync = synchronise all Lists (rewrite) or LUM_Append = appen only new Data
'
' Notes     :
' -----------------------------------------------------------------------------------
Public Sub UpdateActorsLists(ByVal targetBook As Workbook, _
    Optional ByVal dataSrcSheet As Worksheet = Nothing, _
    Optional ByVal mode As ListUpdateMode = LUM_Sync)
    On Error GoTo ErrHandler
    
    modUtil.HideOpMode True, affectScreen:=False, affectEvents:=False
        
    Dim listsSheet As Worksheet
    Dim dataList As ListObject
    If Not GetListsSheetAndTable(targetBook, listsSheet, dataList) Then Exit Sub
    
    Dim actorsDict As Scripting.Dictionary
    Set actorsDict = New Scripting.Dictionary
    
    ' Collect actors
    If mode = LUM_Sync Or dataSrcSheet Is Nothing Then
        CollectActors targetBook, actorsDict
    Else
        CollectActors targetBook, actorsDict, dataSrcSheet
    End If
    
    ' Write based on mode
    If mode = LUM_Sync Then
        ' SYNC: Clear + Rewrite
        modLists.ClearTableColumn dataList, LISTS_HEADER_ACTOR_ID
        modLists.ClearTableColumn dataList, LISTS_HEADER_ACTOR_NAME
        modLists.WriteDictToTableColumns dataList, LISTS_HEADER_ACTOR_ID, actorsDict, LISTS_HEADER_ACTOR_NAME
    Else
        ' APPEND: Add only new entries
        Dim existingKeysDict As Scripting.Dictionary
        Set existingKeysDict = New Scripting.Dictionary
        
        modLists.CollectTableColumnsToDict dataList, LISTS_HEADER_ACTOR_ID, existingKeysDict, LISTS_HEADER_ACTOR_NAME
        modLists.AppendMissingDictSetToTableColumns dataList, LISTS_HEADER_ACTOR_ID, _
            existingKeysDict, actorsDict, LISTS_HEADER_ACTOR_NAME
    End If
    
CleanExit:
    modUtil.HideOpMode False
    Exit Sub
    
ErrHandler:
    modUtil.HideOpMode False
    modErr.ReportError "modRooms.UpdateActorsList", Err.Number, Erl, caption:=modMain.AppProjectName
    Exit Sub
End Sub


' -----------------------------------------------------------------------------------
' Procedure : UpdateSoundsLists
' Purpose   : Updates Sound ID + Description + Type lists.
'
' Parameters:
'   targetBook    [Workbook]          - Target workbook
'   dataSrcSheet   [Worksheet]         - (Optional) If Sheet is provided, only this sheet is processed
'   mode          [ListUpdateMode]    - LUM_Sync = synchronise all Lists (rewrite) or LUM_Append = appen only new Data
'
' Notes     :
' -----------------------------------------------------------------------------------
Public Sub UpdateSoundsLists(ByVal targetBook As Workbook, _
    Optional ByVal dataSrcSheet As Worksheet = Nothing, _
    Optional ByVal mode As ListUpdateMode = LUM_Sync)
    On Error GoTo ErrHandler
    
    modUtil.HideOpMode True, affectScreen:=False, affectEvents:=False
    
    Dim listsSheet As Worksheet
    Dim dataList As ListObject
    If Not GetListsSheetAndTable(targetBook, listsSheet, dataList) Then Exit Sub
    
    Dim soundsDict As Scripting.Dictionary
    Set soundsDict = New Scripting.Dictionary
    
    ' Collect sounds
    If mode = LUM_Sync Or dataSrcSheet Is Nothing Then
        CollectSounds targetBook, soundsDict
    Else
        CollectSounds targetBook, soundsDict, dataSrcSheet
    End If
    
    ' Write based on mode
    If mode = LUM_Sync Then
        ' SYNC: Clear + Rewrite
        modLists.ClearTableColumn dataList, LISTS_HEADER_SOUND_ID
        modLists.ClearTableColumn dataList, LISTS_HEADER_SOUND_NAME
        modLists.WriteDictToTableColumns dataList, LISTS_HEADER_SOUND_ID, soundsDict, _
            LISTS_HEADER_SOUND_NAME
    Else
        ' APPEND: Add only new entries
        Dim existingKeysDict As Scripting.Dictionary
        Set existingKeysDict = New Scripting.Dictionary
        
        modLists.CollectTableColumnsToDict dataList, LISTS_HEADER_SOUND_ID, existingKeysDict, _
            LISTS_HEADER_SOUND_NAME
        modLists.AppendMissingDictSetToTableColumns dataList, LISTS_HEADER_SOUND_ID, _
            existingKeysDict, soundsDict, LISTS_HEADER_SOUND_NAME
    End If
    
    modUtil.HideOpMode False
    Exit Sub
    
ErrHandler:
    modUtil.HideOpMode False
    modErr.ReportError "modRooms.UpdateSoundsList", Err.Number, Erl, caption:=modMain.AppProjectName
End Sub

' -----------------------------------------------------------------------------------
' Procedure : UpdateSpecialFXLists
' Purpose   : Updates Special FX (Animation ID + Description + Type) lists.
'
' Parameters:
'   targetBook    [Workbook]          - Target workbook
'   dataSrcSheet   [Worksheet]         - (Optional) If Sheet is provided, only this sheet is processed
'   mode          [ListUpdateMode]    - LUM_Sync = synchronise all Lists (rewrite) or LUM_Append = appen only new Data
'
' Notes     :
' -----------------------------------------------------------------------------------
Public Sub UpdateSpecialFXLists(ByVal targetBook As Workbook, _
    Optional ByVal dataSrcSheet As Worksheet = Nothing, _
    Optional ByVal mode As ListUpdateMode = LUM_Sync)
    On Error GoTo ErrHandler
    
    modUtil.HideOpMode True, affectScreen:=False, affectEvents:=False
    
    Dim listsSheet As Worksheet
    Dim dataList As ListObject
    If Not GetListsSheetAndTable(targetBook, listsSheet, dataList) Then Exit Sub
    
    Dim specialFXDict As Scripting.Dictionary
    Set specialFXDict = New Scripting.Dictionary
    
    ' Collect special FX
    If mode = LUM_Sync Or dataSrcSheet Is Nothing Then
        CollectSpecialFX targetBook, specialFXDict
    Else
        CollectSpecialFX targetBook, specialFXDict, dataSrcSheet
    End If
    
    ' Write based on mode
    If mode = LUM_Sync Then
        ' SYNC: Clear + Rewrite
        modLists.ClearTableColumn dataList, LISTS_HEADER_ANIMATION_ID
        modLists.ClearTableColumn dataList, LISTS_HEADER_ANIMATION_NAME
        modLists.WriteDictToTableColumns dataList, LISTS_HEADER_ANIMATION_ID, specialFXDict, _
            LISTS_HEADER_ANIMATION_NAME
    Else
        ' APPEND: Add only new entries
        Dim existingKeysDict As Scripting.Dictionary
        Set existingKeysDict = New Scripting.Dictionary
        
        modLists.CollectTableColumnsToDict dataList, LISTS_HEADER_ANIMATION_ID, existingKeysDict, _
            LISTS_HEADER_ANIMATION_NAME
        modLists.AppendMissingDictSetToTableColumns dataList, LISTS_HEADER_ANIMATION_ID, _
            existingKeysDict, specialFXDict, LISTS_HEADER_ANIMATION_NAME
    End If
    
    modUtil.HideOpMode False
    Exit Sub
    
ErrHandler:
    modUtil.HideOpMode False
    modErr.ReportError "modRooms.UpdateSpecialFXList", Err.Number, Erl, caption:=modMain.AppProjectName
End Sub

' -----------------------------------------------------------------------------------
' Procedure : UpdateFlagsLists
' Purpose   : Updates Flag ID + Description + Bool Type lists.
'
' Parameters:
'   targetBook    [Workbook]          - Target workbook
'   dataSrcSheet   [Worksheet]         - (Optional) If Sheet is provided, only this sheet is processed
'   mode          [ListUpdateMode]    - LUM_Sync = synchronise all Lists (rewrite) or LUM_Append = appen only new Data
'
' Notes     :
' -----------------------------------------------------------------------------------
Public Sub UpdateFlagsLists(ByVal targetBook As Workbook, _
    Optional ByVal dataSrcSheet As Worksheet = Nothing, _
    Optional ByVal mode As ListUpdateMode = LUM_Sync)
    On Error GoTo ErrHandler
    
    modUtil.HideOpMode True, affectScreen:=False, affectEvents:=False
    
    Dim listsSheet As Worksheet
    Dim dataList As ListObject
    If Not GetListsSheetAndTable(targetBook, listsSheet, dataList) Then Exit Sub
    
    Dim flagsDict As Scripting.Dictionary
    Set flagsDict = New Scripting.Dictionary
    Dim flagsTypeDict As Scripting.Dictionary
    Set flagsTypeDict = New Scripting.Dictionary
    
    ' Collect flags
    If mode = LUM_Sync Or dataSrcSheet Is Nothing Then
        CollectFlags targetBook, flagsDict, flagsTypeDict
    Else
        CollectFlags targetBook, flagsDict, flagsTypeDict, dataSrcSheet
    End If
    
    ' Write based on mode
    If mode = LUM_Sync Then
        ' SYNC: Clear + Rewrite
        modLists.ClearTableColumn dataList, LISTS_HEADER_FLAG_ID
        modLists.ClearTableColumn dataList, LISTS_HEADER_FLAG_NAME
        modLists.ClearTableColumn dataList, LISTS_HEADER_FLAG_TYPE
        modLists.WriteDictToTableColumns dataList, LISTS_HEADER_FLAG_ID, flagsDict, _
            LISTS_HEADER_FLAG_NAME
        modLists.WriteDictToTableColumns dataList, LISTS_HEADER_FLAG_TYPE, flagsTypeDict
    Else
        ' APPEND: Add only new entries
        Dim existingFlagsDict As Scripting.Dictionary
        Set existingFlagsDict = New Scripting.Dictionary
        Dim existingFlagsTypeDict As Scripting.Dictionary
        Set existingFlagsTypeDict = New Scripting.Dictionary
        
        modLists.CollectTableColumnsToDict dataList, LISTS_HEADER_FLAG_ID, existingFlagsDict, _
            LISTS_HEADER_FLAG_NAME
        modLists.CollectTableColumnsToDict dataList, LISTS_HEADER_FLAG_TYPE, existingFlagsTypeDict
        
        modLists.AppendMissingDictSetToTableColumns dataList, LISTS_HEADER_FLAG_ID, _
            existingFlagsDict, flagsDict, LISTS_HEADER_FLAG_NAME
        modLists.AppendMissingDictSetToTableColumns dataList, LISTS_HEADER_FLAG_TYPE, _
            existingFlagsTypeDict, flagsTypeDict
    End If
    
    modUtil.HideOpMode False
    Exit Sub
    
ErrHandler:
    modUtil.HideOpMode False
    modErr.ReportError "modRooms.UpdateFlagsList", Err.Number, Erl, caption:=modMain.AppProjectName
End Sub


' -----------------------------------------------------------------------------------
' Procedure : UpdateItemsLists
' Purpose   : Updates Pickupable Objects (Item ID + Name) lists.
'
' Parameters:
'   targetBook    [Workbook]          - Target workbook
'   dataSrcSheet   [Worksheet]         - (Optional) If Sheet is provided, only this sheet is processed
'   mode          [ListUpdateMode]    - LUM_Sync = synchronise all Lists (rewrite) or LUM_Append = appen only new Data
'
' Notes     :
' -----------------------------------------------------------------------------------
Public Sub UpdateItemsLists(ByVal targetBook As Workbook, _
    Optional ByVal dataSrcSheet As Worksheet = Nothing, _
    Optional ByVal mode As ListUpdateMode = LUM_Sync)
    On Error GoTo ErrHandler
    
    modUtil.HideOpMode True, affectScreen:=False, affectEvents:=False
    
    Dim listsSheet As Worksheet
    Dim dataList As ListObject
    If Not GetListsSheetAndTable(targetBook, listsSheet, dataList) Then Exit Sub
    
    Dim itemsDict As Scripting.Dictionary
    Set itemsDict = New Scripting.Dictionary
    
    ' Collect items
    If mode = LUM_Sync Or dataSrcSheet Is Nothing Then
        CollectItems targetBook, itemsDict
    Else
        CollectItems targetBook, itemsDict, dataSrcSheet
    End If
    
    ' Write based on mode
    If mode = LUM_Sync Then
        ' SYNC: Clear + Rewrite (full refresh)
        modLists.ClearTableColumn dataList, LISTS_HEADER_ITEM_ID
        modLists.ClearTableColumn dataList, LISTS_HEADER_ITEM_NAME
        modLists.WriteDictToTableColumns dataList, LISTS_HEADER_ITEM_ID, itemsDict, LISTS_HEADER_ITEM_NAME
    Else
        ' APPEND: Add only new entries
        Dim existingKeysDict As Scripting.Dictionary
        Set existingKeysDict = New Scripting.Dictionary
        
        modLists.CollectTableColumnsToDict dataList, LISTS_HEADER_ITEM_ID, existingKeysDict, LISTS_HEADER_ITEM_NAME
        modLists.AppendMissingDictSetToTableColumns dataList, LISTS_HEADER_ITEM_ID, _
            existingKeysDict, itemsDict, LISTS_HEADER_ITEM_NAME
    End If
    
    modUtil.HideOpMode False
    Exit Sub
    
ErrHandler:
    modUtil.HideOpMode False
    modErr.ReportError "modRooms.UpdateItemsList", Err.Number, Erl, caption:=modMain.AppProjectName
End Sub

' -----------------------------------------------------------------------------------
' Procedure : UpdateStateObjectsLists
' Purpose   : Updates Multi-State Objects (State ID + Name) lists.
'
' Parameters:
'   targetBook    [Workbook]          - Target workbook
'   dataSrcSheet   [Worksheet]         - (Optional) If Sheet is provided, only this sheet is processed
'   mode          [ListUpdateMode]    - LUM_Sync = synchronise all Lists (rewrite) or LUM_Append = appen only new Data
'
' Notes     :
' -----------------------------------------------------------------------------------
Public Sub UpdateStateObjectsLists(ByVal targetBook As Workbook, _
    Optional ByVal dataSrcSheet As Worksheet = Nothing, _
    Optional ByVal mode As ListUpdateMode = LUM_Sync)
    On Error GoTo ErrHandler
    
    modUtil.HideOpMode True, affectScreen:=False, affectEvents:=False
    
    Dim listsSheet As Worksheet
    Dim dataList As ListObject
    If Not GetListsSheetAndTable(targetBook, listsSheet, dataList) Then Exit Sub
    
    Dim objectsDict As Scripting.Dictionary
    Set objectsDict = New Scripting.Dictionary
    Dim objectsStateDict As Scripting.Dictionary
    Set objectsStateDict = New Scripting.Dictionary
    
    If mode = LUM_Sync Or dataSrcSheet Is Nothing Then
        CollectStateObjects targetBook, objectsDict, objectsStateDict
    Else
        CollectStateObjects targetBook, objectsDict, objectsStateDict, dataSrcSheet
    End If
    
    If mode = LUM_Sync Then
        modLists.ClearTableColumn dataList, LISTS_HEADER_STATE_OBJECT_ID
        modLists.ClearTableColumn dataList, LISTS_HEADER_STATE_OBJECT_NAME
        modLists.ClearTableColumn dataList, LISTS_HEADER_STATE_OBJECT_STATE
        modLists.WriteDictToTableColumns dataList, LISTS_HEADER_STATE_OBJECT_ID, objectsDict, LISTS_HEADER_STATE_OBJECT_NAME
        modLists.WriteDictToTableColumns dataList, LISTS_HEADER_STATE_OBJECT_STATE, objectsStateDict
    Else
        Dim existingObjectKeysDict  As Scripting.Dictionary
        Set existingObjectKeysDict = New Scripting.Dictionary
        Dim existingObjectStateKeysDict As Scripting.Dictionary
        Set existingObjectStateKeysDict = New Scripting.Dictionary
        
        modLists.CollectTableColumnsToDict dataList, LISTS_HEADER_STATE_OBJECT_ID, existingObjectKeysDict, LISTS_HEADER_STATE_OBJECT_NAME
        modLists.CollectTableColumnsToDict dataList, LISTS_HEADER_STATE_OBJECT_STATE, existingObjectStateKeysDict
        
        modLists.AppendMissingDictSetToTableColumns dataList, LISTS_HEADER_STATE_OBJECT_ID, _
            existingObjectKeysDict, objectsDict, LISTS_HEADER_STATE_OBJECT_NAME
            
        modLists.AppendMissingDictSetToTableColumns dataList, LISTS_HEADER_STATE_OBJECT_STATE, _
            existingObjectStateKeysDict, objectsStateDict
    End If
    
    modUtil.HideOpMode False
    Exit Sub
    
ErrHandler:
    modUtil.HideOpMode False
    modErr.ReportError "modRooms.UpdateStateObjectsList", Err.Number, Erl, caption:=modMain.AppProjectName
End Sub

' -----------------------------------------------------------------------------------
' Procedure : UpdateHotspotsLists
' Purpose   : Updates Touchable Objects (Hotspot ID + Name) lists.
'
' Parameters:
'   targetBook    [Workbook]          - Target workbook
'   dataSrcSheet   [Worksheet]         - (Optional) If Sheet is provided, only this sheet is processed
'   mode          [ListUpdateMode]    - LUM_Sync = synchronise all Lists (rewrite) or LUM_Append = appen only new Data
'
' Notes     :
' -----------------------------------------------------------------------------------
Public Sub UpdateHotspotsLists(ByVal targetBook As Workbook, _
    Optional ByVal dataSrcSheet As Worksheet = Nothing, _
    Optional ByVal mode As ListUpdateMode = LUM_Sync)
    On Error GoTo ErrHandler
    
    modUtil.HideOpMode True, affectScreen:=False, affectEvents:=False
    
    Dim listsSheet As Worksheet
    Dim dataList As ListObject
    If Not GetListsSheetAndTable(targetBook, listsSheet, dataList) Then Exit Sub
    
    Dim hotspotsDict As Scripting.Dictionary
    Set hotspotsDict = New Scripting.Dictionary
    
    If mode = LUM_Sync Or dataSrcSheet Is Nothing Then
        CollectHotspots targetBook, hotspotsDict
    Else
        CollectHotspots targetBook, hotspotsDict, dataSrcSheet
    End If
    
    If mode = LUM_Sync Then
        modLists.ClearTableColumn dataList, LISTS_HEADER_HOTSPOT_ID
        modLists.ClearTableColumn dataList, LISTS_HEADER_HOTSPOT_NAME
        modLists.WriteDictToTableColumns dataList, LISTS_HEADER_HOTSPOT_ID, hotspotsDict, LISTS_HEADER_HOTSPOT_NAME
    Else
        Dim existingKeysDict As Scripting.Dictionary
        Set existingKeysDict = New Scripting.Dictionary
        
        modLists.CollectTableColumnsToDict dataList, LISTS_HEADER_HOTSPOT_ID, existingKeysDict, LISTS_HEADER_HOTSPOT_NAME
        modLists.AppendMissingDictSetToTableColumns dataList, LISTS_HEADER_HOTSPOT_ID, _
            existingKeysDict, hotspotsDict, LISTS_HEADER_HOTSPOT_NAME
    End If
    
    modUtil.HideOpMode False
    Exit Sub
    
ErrHandler:
    modUtil.HideOpMode False
    modErr.ReportError "modRooms.UpdateHotspotsList", Err.Number, Erl, caption:=modMain.AppProjectName
End Sub

' -----------------------------------------------------------------------------------
' Procedure : ValidateRooms
' Purpose   : Validates all room sheets for duplicate IDs, missing references,
'             and potential cycles in room connections.
'
' Parameters:
'   targetBook  [Workbook]  - Workbook to validate
'
' Returns   : (Long) - returns number of found issues
'
' Notes     :
'   - Checks for duplicate Room IDs, Room Aliases, and Room Numbers
'   - Validates that all room references in "Doors To" exist
'   - Detects circular dependencies between rooms
'   - Creates or clears existing "Validation Results" sheet
' -----------------------------------------------------------------------------------
Public Function ValidateRooms(ByVal targetBook As Workbook) As Long
    On Error GoTo ErrHandler
    
    Dim validationSheet As Worksheet
    Dim issueRow As Long
    
    ' Enable silent mode
    modUtil.HideOpMode True
    
    ' Ensure validation results sheet
    Set validationSheet = EnsureValidationSheet(targetBook)
    issueRow = 2 ' Start after header row
    
    ' === 1. Check for duplicate Room IDs ===
    issueRow = CheckDuplicateRoomIDs(targetBook, validationSheet, issueRow)
    
    ' === 2. Check for duplicate Room Aliases ===
    issueRow = CheckDuplicateRoomAliases(targetBook, validationSheet, issueRow)
    
    ' === 3. Check for duplicate Room Numbers ===
    issueRow = CheckDuplicateRoomNumbers(targetBook, validationSheet, issueRow)
    
    ' === 4. Check for missing room references ===
    issueRow = CheckMissingRoomReferences(targetBook, validationSheet, issueRow)
    
    ' === 5. Check for circular dependencies ===
    issueRow = CheckCircularDependencies(targetBook, validationSheet, issueRow)
    
    ' Format results
    FormatValidationSheet validationSheet, issueRow
    
    ' Show results
    validationSheet.Activate
    
    ValidateRooms = (issueRow - 2)
    
    If issueRow = 2 Then
        ' No issues found
        validationSheet.Cells(issueRow, 1).value = "Success"
        validationSheet.Cells(issueRow, 2).value = "No validation issues found!"
        validationSheet.Cells(issueRow, 3).value = ""
        validationSheet.Cells(issueRow, 1).Font.Color = RGB(0, 128, 0)
        validationSheet.Cells(issueRow, 2).Font.Color = RGB(0, 128, 0)
        MsgBox "Validation completed successfully!" & vbCrLf & _
            "No issues found.", vbInformation, modMain.AppProjectName
    Else
        MsgBox "Validation completed." & vbCrLf & _
            (issueRow - 2) & " issue(s) found." & vbCrLf & vbCrLf & _
            "Please review the 'Validation Results' sheet.", _
            vbExclamation, modMain.AppProjectName
    End If
    
CleanExit:
    modUtil.HideOpMode False
    Exit Function
    
ErrHandler:
    modUtil.HideOpMode False
    modErr.ReportError "ValidateRooms", Err.Number, Erl, caption:=modMain.AppProjectName
    Resume CleanExit
End Function


' -----------------------------------------------------------------------------------
' Function  : GetFormattedRoomID
' Purpose   : Build a formatted Room ID from the numeric index using a prefix.
'
' Parameters:
'   roomIdx [Long] - Numeric index.
'
' Returns   : String - e.g., R001 (depends on ROOM_SHEET_ID_PREFIX).
'
' Notes     :
' -----------------------------------------------------------------------------------
Public Function GetFormattedRoomID(ByVal roomIdx As Long) As String
    GetFormattedRoomID = ROOM_SHEET_ID_PREFIX & Format(roomIdx, "000")
End Function

' -----------------------------------------------------------------------------------
' Function  : IsValidAGSRoomNo
' Purpose   : Validates whether a given number is a valid AGS room number
'
' Parameters: roomNo - The room number to validate
'
' Returns   : Boolean - True if valid AGS room number, False otherwise
'
' Notes     : AGS room numbers must be in the range 0-999
'             Adventure Game Studio supports room numbers from 0 to 999
' -----------------------------------------------------------------------------------
Public Function IsValidAGSRoomNo(ByVal roomNo As Long) As Boolean
    IsValidAGSRoomNo = (roomNo >= 0 And roomNo <= 999)
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
' Procedure : SetupRoom
' Purpose   : Initializes a newly created Room sheet by setting RoomID, RoomAlias,
'             RoomNo, removing stale template name references, and wiring control
'             actions (e.g., insert picture button).
'
' Parameters:
'   targetSheet       [Worksheet]  - The newly created Room worksheet to initialize
'   roomName          [String]     - Human-readable room name (e.g., "Temple Entrance")
'                                    used to generate the room alias and sheet name
'   roomIdx           [Long]       - Numeric index used to format the Room ID
'                                    (e.g., 1 => "R001")
'   roomNo            [Long]       - AGS room number (0-999)
'
' Returns   : (none)
'
' Notes     :
'   - RoomID format: ROOM_SHEET_ID_PREFIX & roomIdx (e.g., "R001")
'   - RoomAlias format: ROOM_SHEET_ALIAS_PREFIX & GetCleanRoomAlias(roomName)
'     (e.g., "Temple Entrance" => "r_TempleEntrance")
'   - Removes any lingering named range references to the add-in workbook
'   - Wires the "Insert Room Picture" button to the correct macro
' -----------------------------------------------------------------------------------
Private Sub SetupRoom(targetSheet As Worksheet, roomName As String, roomIdx As Long, roomNo As Long, sceneID As String)
    On Error GoTo ErrHandler
    Dim insertBtnShape As Shape
    Dim dispatcherSheet As Worksheet
    Dim dataRange As Range
    Dim nm As name
    
    ' Set Scene ID in Named Range:
    targetSheet.Range(modConst.NAME_CELL_SCENE_ID).value = sceneID
    
    ' Set RoomID named cell (e.g., "R001")
    targetSheet.Range(modConst.NAME_CELL_ROOM_ID).value = GetFormattedRoomID(roomIdx)
    
    ' Set RoomAlias named cell (e.g., "r_TempleEntrance")
    targetSheet.Range(modConst.NAME_CELL_ROOM_ALIAS).value = _
        modConst.ROOM_SHEET_ALIAS_PREFIX & GetCleanRoomAlias(roomName)
    
    ' Set RoomNo named cell (e.g., 42)
    targetSheet.Range(modConst.NAME_CELL_ROOM_NO).value = roomNo
    
    If modOptions.Opt_DefaultGameHeight > 0 Then
        targetSheet.Range(NAME_CELL_GAME_HEIGHT).value = modOptions.Opt_DefaultGameHeight
    End If

    If modOptions.Opt_DefaultGameWidth > 0 Then
        targetSheet.Range(NAME_CELL_GAME_WIDTH).value = modOptions.Opt_DefaultGameWidth
    End If
    
    If modOptions.Opt_DefaultBGHeight > 0 Then
        targetSheet.Range(NAME_CELL_BG_HEIGHT).value = modOptions.Opt_DefaultBGHeight
    End If

    If modOptions.Opt_DefaultBGWidth > 0 Then
        targetSheet.Range(NAME_CELL_BG_WIDTH).value = modOptions.Opt_DefaultBGWidth
    End If
    
    If modOptions.Opt_DefaultUIHeight > 0 Then
        targetSheet.Range(NAME_CELL_UI_HEIGHT).value = modOptions.Opt_DefaultUIHeight
    End If
    
    If LenB(modOptions.Opt_DefaultPerspective) > 0 Then
        targetSheet.Range(NAME_CELL_PERSPECTIVE).value = modOptions.Opt_DefaultPerspective
    End If

    If LenB(modOptions.Opt_DefaultParallax) > 0 Then
        targetSheet.Range(NAME_CELL_PARALLAX).value = modOptions.Opt_DefaultParallax
    End If

    If LenB(modOptions.Opt_DefaultSceneMode) > 0 Then
        targetSheet.Range(NAME_CELL_SCENE_MODE).value = modOptions.Opt_DefaultSceneMode
    End If
    
    ' Remove stale named range references to add-in template workbook
    For Each nm In targetSheet.Parent.Names
        If InStr(nm.RefersTo, "[" & RDDAddInWkBk.name & "]") > 0 Then
            nm.Delete
        End If
    Next nm
    
    ' Wire the "Insert Room Picture" button to the correct macro
    Set dispatcherSheet = modSheets.GetSheetByCodeName(modConst.SHEET_DISPATCHER)
    Set insertBtnShape = targetSheet.Shapes(modConst.BTN_INSERT_ROOM_PICTURE)
    insertBtnShape.OnAction = modConst.MACRO_BTN_INSERT_PICTURE
         
CleanExit:
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
        If targetSheet.name = sheetToExclude.name Then GoTo NextSheet
        
        ' Only check Room sheets
        If Not IsRoomSheet(targetSheet, currentRoomID) Then GoTo NextSheet
        
        ' Check for references in "DOORS TO..." area
        On Error Resume Next
        Set cell = targetSheet.Range(NAME_RANGE_DOORS_TO_ROOM_ID)
        On Error GoTo 0
        
        If Not cell Is Nothing Then
            If modRanges.RangeHasValue(cell, roomID, True, False) Then
                col.Add targetSheet.name
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

    ' List of characters to be removed
    removeCharArray = Array(" ", "-", ".", "(", ")", ":", "/", "'")

    ' Replace all characters with an empty string
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
' Procedure : CollectRoomMetadata
' Purpose   : Collects Room ID, Room No, and Room Alias from room sheets.
'
' Parameters:
'   targetBook     [Workbook]         - Workbook to scan
'   roomsDict      [Dictionary]       - Receives: Key=RoomID, Value="RoomNo|RoomAlias"
'   onlyFromSheet  [Worksheet]        - (Optional) If provided, only collect from this sheet
'
' Notes     :
'   - If onlyFromSheet is Nothing: collect from ALL room sheets
'   - If onlyFromSheet is provided: collect only from that sheet
' -----------------------------------------------------------------------------------
Private Sub CollectRoomMetadata(ByVal targetBook As Workbook, _
    ByRef roomsDict As Scripting.Dictionary, _
    Optional ByVal onlyFromSheet As Worksheet = Nothing)
    
    On Error GoTo ErrHandler
        
    Dim targetSheet As Worksheet: Set targetSheet = onlyFromSheet
    Dim roomID As String
    Dim roomNo As String
    Dim roomAlias As String
    
    If Not targetSheet Is Nothing Then
        
        roomID = modLists.GetNamedOrHeaderValue(targetSheet, NAME_CELL_ROOM_ID)
        If Len(roomID) > 0 Then
            roomNo = modLists.GetNamedOrHeaderValue(targetSheet, NAME_CELL_ROOM_NO)
            roomAlias = modLists.GetNamedOrHeaderValue(targetSheet, NAME_CELL_ROOM_ALIAS)
            roomsDict(roomID) = roomNo & DICT_VALUE_SEPERATOR & roomAlias
        End If
        
    Else
        For Each targetSheet In targetBook.Worksheets
        
            If modRooms.IsRoomSheet(targetSheet, roomID) Then
            
                If Len(roomID) > 0 Then
                    roomNo = modLists.GetNamedOrHeaderValue(targetSheet, NAME_CELL_ROOM_NO)
                    roomAlias = modLists.GetNamedOrHeaderValue(targetSheet, NAME_CELL_ROOM_ALIAS)
                    roomsDict(roomID) = roomNo & DICT_VALUE_SEPERATOR & roomAlias
                End If

            End If
        
        Next targetSheet
    End If
    
    Exit Sub
    
ErrHandler:
    ' Forward error
    Err.Raise Err.Number, "modRooms.CollectRoomMetadata", Err.Description
End Sub

' -----------------------------------------------------------------------------------
' Procedure : CollectSceneIDs
' Purpose   : Collects Scene IDs from room sheets.
'
' Parameters:
'   targetBook     [Workbook]         - Workbook to scan
'   scenesDict     [Dictionary]       - Receives: Key=SceneID, Value=True
'   onlyFromSheet  [Worksheet]        - (Optional) If provided, only collect from this sheet
' -----------------------------------------------------------------------------------
Private Sub CollectSceneIDs(ByVal targetBook As Workbook, _
    ByRef scenesDict As Scripting.Dictionary, _
    Optional ByVal onlyFromSheet As Worksheet = Nothing)
    
    On Error GoTo ErrHandler
    
    Dim targetSheet As Worksheet: Set targetSheet = onlyFromSheet
    Dim roomID As String
    Dim sceneID As String
    
    If Not targetSheet Is Nothing Then

        sceneID = modLists.GetNamedOrHeaderValue(targetSheet, NAME_CELL_SCENE_ID)
        If Len(sceneID) > 0 Then scenesDict(sceneID) = True

    Else
    
        For Each targetSheet In targetBook.Worksheets
        
            If modRooms.IsRoomSheet(targetSheet, roomID) Then

                sceneID = modLists.GetNamedOrHeaderValue(targetSheet, NAME_CELL_SCENE_ID)
                If Len(sceneID) > 0 Then scenesDict(sceneID) = True

            End If
        
        Next targetSheet
    End If
    
    Exit Sub
    
ErrHandler:
    ' Forward error
    Err.Raise Err.Number, "modRooms.CollectSceneIDs", Err.Description
End Sub

' -----------------------------------------------------------------------------------
' Procedure : CollectGeneralSettings
' Purpose   : Collects unique dimension values (Width, Height, UI Height) from rooms.
'
' Parameters:
'   targetBook     [Workbook]         - Workbook to scan
'   widthDict      [Dictionary]       - Receives unique width values
'   heightDict     [Dictionary]       - Receives unique height values
'   uiHeightDict   [Dictionary]       - Receives unique UI height values
'
' Notes     :
'   - Collects UNIQUE values only (no duplicates)
'   - Collects from ALL rooms (dimensions are not unique per room)
'   - Values are stored as keys with True as value
' -----------------------------------------------------------------------------------
Private Sub CollectGeneralSettings(ByVal targetBook As Workbook, _
    ByRef widthDict As Scripting.Dictionary, _
    ByRef heightDict As Scripting.Dictionary, _
    ByRef uiHeightDict As Scripting.Dictionary, _
    Optional ByVal onlyFromSheet As Worksheet = Nothing)
    
    On Error GoTo ErrHandler
        
    Dim targetSheet As Worksheet: Set targetSheet = onlyFromSheet
    Dim roomID As String
    Dim gameWidth As String
    Dim gameHeight As String
    Dim bgWidth As String
    Dim bgHeight As String
    Dim uiHeight As String
    
    If Not targetSheet Is Nothing Then

        ' Collect Game dimensions
        gameWidth = modLists.GetNamedOrHeaderValue(targetSheet, NAME_CELL_GAME_WIDTH)
        If Len(gameWidth) > 0 Then widthDict(gameWidth) = True
            
        gameHeight = modLists.GetNamedOrHeaderValue(targetSheet, NAME_CELL_GAME_HEIGHT)
        If Len(gameHeight) > 0 Then heightDict(gameHeight) = True
            
        ' Collect BG dimensions
        bgWidth = modLists.GetNamedOrHeaderValue(targetSheet, NAME_CELL_BG_WIDTH)
        If Len(bgWidth) > 0 Then widthDict(bgWidth) = True
            
        bgHeight = modLists.GetNamedOrHeaderValue(targetSheet, NAME_CELL_BG_HEIGHT)
        If Len(bgHeight) > 0 Then heightDict(bgHeight) = True
            
        ' Collect UI Height
        uiHeight = modLists.GetNamedOrHeaderValue(targetSheet, NAME_CELL_UI_HEIGHT)
        If Len(uiHeight) > 0 Then uiHeightDict(uiHeight) = True
    
    Else
        For Each targetSheet In targetBook.Worksheets
            If modRooms.IsRoomSheet(targetSheet, roomID) Then
            
                ' Collect Game dimensions
                gameWidth = modLists.GetNamedOrHeaderValue(targetSheet, NAME_CELL_GAME_WIDTH)
                If Len(gameWidth) > 0 Then widthDict(gameWidth) = True
            
                gameHeight = modLists.GetNamedOrHeaderValue(targetSheet, NAME_CELL_GAME_HEIGHT)
                If Len(gameHeight) > 0 Then heightDict(gameHeight) = True
            
                ' Collect BG dimensions
                bgWidth = modLists.GetNamedOrHeaderValue(targetSheet, NAME_CELL_BG_WIDTH)
                If Len(bgWidth) > 0 Then widthDict(bgWidth) = True
            
                bgHeight = modLists.GetNamedOrHeaderValue(targetSheet, NAME_CELL_BG_HEIGHT)
                If Len(bgHeight) > 0 Then heightDict(bgHeight) = True
            
                ' Collect UI Height
                uiHeight = modLists.GetNamedOrHeaderValue(targetSheet, NAME_CELL_UI_HEIGHT)
                If Len(uiHeight) > 0 Then uiHeightDict(uiHeight) = True

            End If
        Next targetSheet
    End If
    
    Exit Sub
    
ErrHandler:
    ' Forward error
    Err.Raise Err.Number, "modRooms.CollectGeneralSettings", Err.Description
End Sub

' -----------------------------------------------------------------------------------
' Procedure : CollectActors
' Purpose   : Collects Actor ID + Name from room sheets.
'
' Parameters:
'   targetBook     [Workbook]         - Workbook to scan
'   actorsDict     [Dictionary]       - Receives: Key=ActorID, Value=ActorName
'   onlyFromSheet  [Worksheet]        - (Optional) If provided, only collect from this sheet
' -----------------------------------------------------------------------------------
Private Sub CollectActors(ByVal targetBook As Workbook, _
    ByRef actorsDict As Scripting.Dictionary, _
    Optional ByVal onlyFromSheet As Worksheet = Nothing)
    
    On Error GoTo ErrHandler
    
    Dim targetSheet As Worksheet: Set targetSheet = onlyFromSheet
    Dim roomID As String
    
    If Not targetSheet Is Nothing Then
        
        modLists.CollectNamedRangePairsToDict targetSheet, NAME_RANGE_ACTORS_ACTOR_ID, _
            NAME_RANGE_ACTORS_ACTOR_NAME, actorsDict

    Else
        
        For Each targetSheet In targetBook.Worksheets
            
            If modRooms.IsRoomSheet(targetSheet, roomID) Then

                modLists.CollectNamedRangePairsToDict targetSheet, NAME_RANGE_ACTORS_ACTOR_ID, _
                    NAME_RANGE_ACTORS_ACTOR_NAME, actorsDict
            End If
        Next targetSheet
    End If
    
    Exit Sub
    
ErrHandler:
    ' Forward error
    Err.Raise Err.Number, "modRooms.CollectActors", Err.Description
End Sub

' -----------------------------------------------------------------------------------
' Procedure : CollectSounds
' Purpose   : Collects Sound ID + Description + Type from room sheets.
'
' Parameters:
'   targetBook     [Workbook]         - Workbook to scan
'   soundsDict     [Dictionary]       - Receives: Key=SoundID, Value="Description"
'   onlyFromSheet  [Worksheet]        - (Optional) If provided, only collect from this sheet
'
' Notes     :
' -----------------------------------------------------------------------------------
Private Sub CollectSounds(ByVal targetBook As Workbook, _
    ByRef soundsDict As Scripting.Dictionary, _
    Optional ByVal onlyFromSheet As Worksheet = Nothing)
    
    On Error GoTo ErrHandler
    
    Dim targetSheet As Worksheet: Set targetSheet = onlyFromSheet
    Dim roomID As String
  
    If Not targetSheet Is Nothing Then

        modLists.CollectNamedRangePairsToDict targetSheet, NAME_RANGE_SOUNDS_SOUND_ID, _
            NAME_RANGE_SOUNDS_DESCRIPTION, soundsDict

    Else
    
        For Each targetSheet In targetBook.Worksheets
        
            If modRooms.IsRoomSheet(targetSheet, roomID) Then
                
                ' Collect Sound ID + Description
                modLists.CollectNamedRangePairsToDict targetSheet, NAME_RANGE_SOUNDS_SOUND_ID, _
                    NAME_RANGE_SOUNDS_DESCRIPTION, soundsDict

            End If

        Next targetSheet
    End If
    
    Exit Sub
    
ErrHandler:
    ' Forward error
    Err.Raise Err.Number, "modRooms.CollectSounds", Err.Description
End Sub

' -----------------------------------------------------------------------------------
' Procedure : CollectSpecialFX
' Purpose   : Collects Special FX (Animation ID + Description + Type) from room sheets.
'
' Parameters:
'   targetBook        [Workbook]         - Workbook to scan
'   specialFXDict     [Dictionary]       - Receives: Key=AnimationID, Value="Description"
'   onlyFromSheet     [Worksheet]        - (Optional) If provided, only collect from this sheet
'
' Notes     :
' -----------------------------------------------------------------------------------
Private Sub CollectSpecialFX(ByVal targetBook As Workbook, _
    ByRef specialFXDict As Scripting.Dictionary, _
    Optional ByVal onlyFromSheet As Worksheet = Nothing)
    
    On Error GoTo ErrHandler
    
    Dim targetSheet As Worksheet: Set targetSheet = onlyFromSheet
    Dim roomID As String

    If Not targetSheet Is Nothing Then
    
        modLists.CollectNamedRangePairsToDict targetSheet, NAME_RANGE_SPECIAL_FX_ANIMATION_ID, _
            NAME_RANGE_SPECIAL_FX_DESCRIPTION, specialFXDict

    Else
    
        For Each targetSheet In targetBook.Worksheets
        
            If modRooms.IsRoomSheet(targetSheet, roomID) Then
            
                ' Collect Animation ID + Description
                modLists.CollectNamedRangePairsToDict targetSheet, NAME_RANGE_SPECIAL_FX_ANIMATION_ID, _
                    NAME_RANGE_SPECIAL_FX_DESCRIPTION, specialFXDict

            End If

        Next targetSheet
    End If
    
    Exit Sub
    
ErrHandler:
    ' Forward error
    Err.Raise Err.Number, "modRooms.CollectSpecialFX", Err.Description
End Sub

' -----------------------------------------------------------------------------------
' Procedure : CollectFlags
' Purpose   : Collects Flag ID + Description + Bool Type from room sheets.
'
' Parameters:
'   targetBook     [Workbook]         - Workbook to scan
'   flagsDict      [Dictionary]       - Receives: Key=FlagID, Value="Description"
'   flagsTypeDict  [Dictionary]       - Receives: Key=Flag Bool Type
'   onlyFromSheet  [Worksheet]        - (Optional) If provided, only collect from this sheet
'
' Notes     :
' -----------------------------------------------------------------------------------
Private Sub CollectFlags(ByVal targetBook As Workbook, _
    ByRef flagsDict As Scripting.Dictionary, _
    ByRef flagsTypeDict As Scripting.Dictionary, _
    Optional ByVal onlyFromSheet As Worksheet = Nothing)
    
    On Error GoTo ErrHandler
    
    Dim targetSheet As Worksheet: Set targetSheet = onlyFromSheet
    Dim roomID As String
    
    If Not targetSheet Is Nothing Then

    
        modLists.CollectNamedRangePairsToDict targetSheet, NAME_RANGE_FLAGS_FLAG_ID, _
            NAME_RANGE_FLAGS_DESCRIPTION, flagsDict
                    
        modLists.CollectNamedRangeValuesToDict targetSheet, NAME_RANGE_FLAGS_BOOL_TYPE, flagsTypeDict

    
    Else
        For Each targetSheet In targetBook.Worksheets
        
            If modRooms.IsRoomSheet(targetSheet, roomID) Then
            
                ' Collect Flag ID + Description
                modLists.CollectNamedRangePairsToDict targetSheet, NAME_RANGE_FLAGS_FLAG_ID, _
                    NAME_RANGE_FLAGS_DESCRIPTION, flagsDict
                    
                modLists.CollectNamedRangeValuesToDict targetSheet, NAME_RANGE_FLAGS_BOOL_TYPE, flagsTypeDict

                    
            End If

        Next targetSheet
    
    End If
    
    Exit Sub
    
ErrHandler:
    ' Forward error
    Err.Raise Err.Number, "modRooms.CollectFlags", Err.Description
End Sub

' -----------------------------------------------------------------------------------
' Procedure : CollectItems
' Purpose   : Collects Pickupable Objects (Item ID + Name) from room sheets.
'
' Parameters:
'   targetBook     [Workbook]         - Workbook to scan
'   itemsDict      [Dictionary]       - Receives: Key=ItemID, Value=ItemName
'   onlyFromSheet  [Worksheet]        - (Optional) If provided, only collect from this sheet
' -----------------------------------------------------------------------------------
Private Sub CollectItems(ByVal targetBook As Workbook, _
    ByRef itemsDict As Scripting.Dictionary, _
    Optional ByVal onlyFromSheet As Worksheet = Nothing)
    
    On Error GoTo ErrHandler
    
    Dim targetSheet As Worksheet: Set targetSheet = onlyFromSheet
    Dim roomID As String
    
    If Not targetSheet Is Nothing Then

        modLists.CollectNamedRangePairsToDict targetSheet, NAME_RANGE_PICKUPABLE_OBJECTS_ITEM_ID, _
            NAME_RANGE_PICKUPABLE_OBJECTS_NAME, itemsDict

    Else
        For Each targetSheet In targetBook.Worksheets
        
            If modRooms.IsRoomSheet(targetSheet, roomID) Then

                modLists.CollectNamedRangePairsToDict targetSheet, NAME_RANGE_PICKUPABLE_OBJECTS_ITEM_ID, _
                    NAME_RANGE_PICKUPABLE_OBJECTS_NAME, itemsDict

            End If
        
        Next targetSheet
    
    End If
    
    Exit Sub
    
ErrHandler:
    ' Forward error
    Err.Raise Err.Number, "modRooms.CollectItems", Err.Description
End Sub

' -----------------------------------------------------------------------------------
' Procedure : CollectStateObjects
' Purpose   : Collects Multi-State Objects (State ID + Name + State) from room sheets.
'
' Parameters:
'   targetBook        [Workbook]         - Workbook to scan
'   objectsDict       [Dictionary]       - Receives: Key=StateID, Value=ObjectName
'   objectsStateDict  [Dictionary]       - Receives: Key=objects State
'   onlyFromSheet     [Worksheet]        - (Optional) If provided, only collect from this sheet
' -----------------------------------------------------------------------------------
Private Sub CollectStateObjects(ByVal targetBook As Workbook, _
    ByRef objectsDict As Scripting.Dictionary, _
    ByRef objectsStateDict As Scripting.Dictionary, _
    Optional ByVal onlyFromSheet As Worksheet = Nothing)
    
    On Error GoTo ErrHandler
    
    Dim targetSheet As Worksheet: Set targetSheet = onlyFromSheet
    Dim roomID As String
    
    If Not targetSheet Is Nothing Then
    
        modLists.CollectNamedRangePairsToDict targetSheet, NAME_RANGE_MULTI_STATE_OBJECTS_STATE_ID, _
            NAME_RANGE_MULTI_STATE_OBJECTS_OBJECT_NAME, objectsDict
            
        modLists.CollectNamedRangeValuesToDict targetSheet, NAME_RANGE_MULTI_STATE_OBJECTS_STATE, objectsStateDict

    Else
        For Each targetSheet In targetBook.Worksheets
        
            If modRooms.IsRoomSheet(targetSheet, roomID) Then
            
                modLists.CollectNamedRangePairsToDict targetSheet, NAME_RANGE_MULTI_STATE_OBJECTS_STATE_ID, _
                    NAME_RANGE_MULTI_STATE_OBJECTS_OBJECT_NAME, objectsDict
                    
                modLists.CollectNamedRangeValuesToDict targetSheet, NAME_RANGE_MULTI_STATE_OBJECTS_STATE, objectsStateDict

            End If
        
        Next targetSheet
    
    End If

    Exit Sub
    
ErrHandler:
    ' Forward error
    Err.Raise Err.Number, "modRooms.CollectStateObjects", Err.Description
End Sub

' -----------------------------------------------------------------------------------
' Procedure : CollectHotspots
' Purpose   : Collects Touchable Objects (Hotspot ID + Name) from room sheets.
'
' Parameters:
'   targetBook     [Workbook]         - Workbook to scan
'   hotspotsDict   [Dictionary]       - Receives: Key=HotspotID, Value=HotspotName
'   onlyFromSheet  [Worksheet]        - (Optional) If provided, only collect from this sheet
' -----------------------------------------------------------------------------------
Private Sub CollectHotspots(ByVal targetBook As Workbook, _
    ByRef hotspotsDict As Scripting.Dictionary, _
    Optional ByVal onlyFromSheet As Worksheet = Nothing)
    
    On Error GoTo ErrHandler
    
    Dim targetSheet As Worksheet: Set targetSheet = onlyFromSheet
    Dim roomID As String
    
    If Not targetSheet Is Nothing Then
        modLists.CollectNamedRangePairsToDict targetSheet, NAME_RANGE_TOUCHABLE_OBJECTS_HOTSPOT_ID, _
            NAME_RANGE_TOUCHABLE_OBJECTS_HOTSPOT_NAME, hotspotsDict

    Else
        For Each targetSheet In targetBook.Worksheets
        
            If modRooms.IsRoomSheet(targetSheet, roomID) Then
            
                modLists.CollectNamedRangePairsToDict targetSheet, NAME_RANGE_TOUCHABLE_OBJECTS_HOTSPOT_ID, _
                    NAME_RANGE_TOUCHABLE_OBJECTS_HOTSPOT_NAME, hotspotsDict

            End If
        
        Next targetSheet
    
    End If
    
    Exit Sub
    
ErrHandler:
    ' Forward error
    Err.Raise Err.Number, "modRooms.CollectHotspots", Err.Description
End Sub


' -----------------------------------------------------------------------------------
' Procedure : UpdateRoomMetadataInDispatcherTable
' Purpose   : Updates Room ID and Room Alias in the DropDownLists table.
' -----------------------------------------------------------------------------------
Private Sub UpdateRoomMetadataInDispatcherTable(ByVal targetBook As Workbook, _
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
    modErr.ReportError "modMain.UpdateRoomMetadataInDispatcherTable", Err.Number, Erl, caption:=AppProjectName
End Sub

' -----------------------------------------------------------------------------------
' Function  : EnsureValidationSheet
' Purpose   : Creates or clears the "Validation Results" sheet
'
' Parameters:
'   targetBook  [Workbook]  - Workbook to create sheet in
'
' Returns   : Worksheet - The validation results sheet
' -----------------------------------------------------------------------------------
Private Function EnsureValidationSheet(ByVal targetBook As Workbook) As Worksheet
    Dim ValidSheet As Worksheet
    
    ' Check if sheet exists
    Set ValidSheet = modSheets.EnsureSheet("Validation Results", targetBook)
    
    If Not ValidSheet Is Nothing Then
        ' Clear existing content
        ValidSheet.Cells.Clear
    
        ' Write headers
        ValidSheet.Cells(1, 1).value = "Severity"
        ValidSheet.Cells(1, 2).value = "Issue Type"
        ValidSheet.Cells(1, 3).value = "Description"
    
        ' Format headers
        With ValidSheet.Range("A1:C1")
            .Font.Bold = True
            .Font.Size = 11
            .Interior.Color = RGB(68, 114, 196)
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
        End With
    
        Set EnsureValidationSheet = ValidSheet
    End If
End Function

' -----------------------------------------------------------------------------------
' Procedure : FormatValidationSheet
' Purpose   : Formats the validation results sheet
'
' Parameters:
'   sheet     [Worksheet]  - Sheet to format
'   lastRow   [Long]       - Last row with data
' -----------------------------------------------------------------------------------
Private Sub FormatValidationSheet(ByVal sheet As Worksheet, ByVal lastRow As Long)
    If lastRow > 1 Then
        ' Auto-fit columns
        sheet.Columns("A:C").AutoFit
        
        ' Adjust column widths with limits
        If sheet.Columns(1).ColumnWidth < 10 Then sheet.Columns(1).ColumnWidth = 10
        If sheet.Columns(2).ColumnWidth < 20 Then sheet.Columns(2).ColumnWidth = 20
        If sheet.Columns(3).ColumnWidth > 80 Then sheet.Columns(3).ColumnWidth = 80
        
        ' Add borders
        With sheet.Range(sheet.Cells(1, 1), sheet.Cells(lastRow - 1, 3))
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Borders.Color = RGB(200, 200, 200)
        End With
        
        ' Wrap text in description column
        sheet.Columns(3).WrapText = True
    End If
    
    ' Freeze header row
    sheet.Activate
    sheet.Range("A2").Select
    ActiveWindow.FreezePanes = True
End Sub

' -----------------------------------------------------------------------------------
' Function  : CheckDuplicateRoomIDs
' Purpose   : Checks for duplicate Room IDs across all room sheets
'
' Parameters:
'   targetBook        [Workbook]   - Workbook to check
'   validationSheet   [Worksheet]  - Sheet to write results to
'   startRow          [Long]       - Starting row for writing issues
'
' Returns   : Long - Next available row for writing
' -----------------------------------------------------------------------------------
Private Function CheckDuplicateRoomIDs(ByVal targetBook As Workbook, _
    ByVal validationSheet As Worksheet, _
    ByVal startRow As Long) As Long
    Dim roomIDDict As Scripting.Dictionary
    Dim roomSheet As Worksheet
    Dim roomID As String
    Dim sheetName As String
    Dim currentRow As Long
    
    Set roomIDDict = New Scripting.Dictionary
    currentRow = startRow
    
    ' Collect all Room IDs
    For Each roomSheet In targetBook.Worksheets
        If IsRoomSheet(roomSheet, roomID) Then
            sheetName = roomSheet.name
            
            If Len(roomID) > 0 Then
                If roomIDDict.Exists(roomID) Then
                    ' Duplicate found
                    validationSheet.Cells(currentRow, 1).value = "Error"
                    validationSheet.Cells(currentRow, 2).value = "Duplicate Room ID"
                    validationSheet.Cells(currentRow, 3).value = _
                        "Room ID '" & roomID & "' is used in sheets: '" & _
                        roomIDDict(roomID) & "' and '" & sheetName & "'"
                    validationSheet.Cells(currentRow, 1).Font.Color = RGB(255, 0, 0)
                    currentRow = currentRow + 1
                    
                    ' Update dictionary to track multiple occurrences
                    roomIDDict(roomID) = roomIDDict(roomID) & ", " & sheetName
                Else
                    roomIDDict.Add roomID, sheetName
                End If
            End If
        End If
    Next roomSheet
    
    CheckDuplicateRoomIDs = currentRow
End Function

' -----------------------------------------------------------------------------------
' Function  : CheckDuplicateRoomAliases
' Purpose   : Checks for duplicate Room Aliases across all room sheets
'
' Parameters:
'   targetBook        [Workbook]   - Workbook to check
'   validationSheet   [Worksheet]  - Sheet to write results to
'   startRow          [Long]       - Starting row for writing issues
'
' Returns   : Long - Next available row for writing
' -----------------------------------------------------------------------------------
Private Function CheckDuplicateRoomAliases(ByVal targetBook As Workbook, _
    ByVal validationSheet As Worksheet, _
    ByVal startRow As Long) As Long
    Dim aliasDict As Scripting.Dictionary
    Dim roomSheet As Worksheet
    Dim roomAlias As String
    Dim cell As Range
    Dim currentRow As Long
    
    Set aliasDict = New Scripting.Dictionary
    currentRow = startRow
    
    ' Collect all Room Aliases
    For Each roomSheet In targetBook.Worksheets
        If IsRoomSheet(roomSheet) Then
            On Error Resume Next
            Set cell = roomSheet.Range(modConst.NAME_CELL_ROOM_ALIAS)
            On Error GoTo 0
            
            If Not cell Is Nothing Then
                roomAlias = Trim$(CStr(cell.value))
                
                If Len(roomAlias) > 0 Then
                    If aliasDict.Exists(roomAlias) Then
                        ' Duplicate found
                        validationSheet.Cells(currentRow, 1).value = "Error"
                        validationSheet.Cells(currentRow, 2).value = "Duplicate Room Alias"
                        validationSheet.Cells(currentRow, 3).value = _
                            "Room Alias '" & roomAlias & "' is used in sheets: '" & _
                            aliasDict(roomAlias) & "' and '" & roomSheet.name & "'"
                        validationSheet.Cells(currentRow, 1).Font.Color = RGB(255, 0, 0)
                        currentRow = currentRow + 1
                        
                        aliasDict(roomAlias) = aliasDict(roomAlias) & ", " & roomSheet.name
                    Else
                        aliasDict.Add roomAlias, roomSheet.name
                    End If
                End If
            End If
        End If
    Next roomSheet
    
    CheckDuplicateRoomAliases = currentRow
End Function

' -----------------------------------------------------------------------------------
' Function  : CheckDuplicateRoomNumbers
' Purpose   : Checks for duplicate Room Numbers across all room sheets
'
' Parameters:
'   targetBook        [Workbook]   - Workbook to check
'   validationSheet   [Worksheet]  - Sheet to write results to
'   startRow          [Long]       - Starting row for writing issues
'
' Returns   : Long - Next available row for writing
' -----------------------------------------------------------------------------------
Private Function CheckDuplicateRoomNumbers(ByVal targetBook As Workbook, _
    ByVal validationSheet As Worksheet, _
    ByVal startRow As Long) As Long
    Dim roomNoDict As Scripting.Dictionary
    Dim roomSheet As Worksheet
    Dim roomNo As Long
    Dim cell As Range
    Dim cellValue As Variant
    Dim currentRow As Long
    
    Set roomNoDict = New Scripting.Dictionary
    currentRow = startRow
    
    ' Collect all Room Numbers
    For Each roomSheet In targetBook.Worksheets
        If IsRoomSheet(roomSheet) Then
            On Error Resume Next
            Set cell = roomSheet.Range(modConst.NAME_CELL_ROOM_NO)
            On Error GoTo 0
            
            If Not cell Is Nothing Then
                cellValue = cell.value
                
                If IsNumeric(cellValue) Then
                    roomNo = CLng(cellValue)
                    
                    If roomNo >= 0 Then ' Only check valid room numbers
                        Dim roomNoKey As String
                        roomNoKey = CStr(roomNo)
                        
                        If roomNoDict.Exists(roomNoKey) Then
                            ' Duplicate found
                            validationSheet.Cells(currentRow, 1).value = "Error"
                            validationSheet.Cells(currentRow, 2).value = "Duplicate Room Number"
                            validationSheet.Cells(currentRow, 3).value = _
                                "Room Number '" & roomNo & "' is used in sheets: '" & _
                                roomNoDict(roomNoKey) & "' and '" & roomSheet.name & "'"
                            validationSheet.Cells(currentRow, 1).Font.Color = RGB(255, 0, 0)
                            currentRow = currentRow + 1
                            
                            roomNoDict(roomNoKey) = roomNoDict(roomNoKey) & ", " & roomSheet.name
                        Else
                            roomNoDict.Add roomNoKey, roomSheet.name
                        End If
                    End If
                End If
            End If
        End If
    Next roomSheet
    
    CheckDuplicateRoomNumbers = currentRow
End Function

' -----------------------------------------------------------------------------------
' Function  : CheckMissingRoomReferences
' Purpose   : Checks if all room references in "Doors To" ranges exist
'
' Parameters:
'   targetBook        [Workbook]   - Workbook to check
'   validationSheet   [Worksheet]  - Sheet to write results to
'   startRow          [Long]       - Starting row for writing issues
'
' Returns   : Long - Next available row for writing
' -----------------------------------------------------------------------------------
Private Function CheckMissingRoomReferences(ByVal targetBook As Workbook, _
    ByVal validationSheet As Worksheet, _
    ByVal startRow As Long) As Long
    Dim roomSheet As Worksheet
    Dim refRange As Range
    Dim cell As Range
    Dim refRoomID As String
    Dim refRoomAlias As String
    Dim currentRow As Long
    Dim foundSheet As Worksheet
    
    currentRow = startRow
    
    ' Check each room sheet
    For Each roomSheet In targetBook.Worksheets
        If IsRoomSheet(roomSheet) Then
            ' Check "Doors To Room ID" references
            On Error Resume Next
            Set refRange = roomSheet.Range(modConst.NAME_RANGE_DOORS_TO_ROOM_ID)
            On Error GoTo 0
            
            If Not refRange Is Nothing Then
                For Each cell In refRange.Cells
                    refRoomID = Trim$(CStr(cell.value))
                    
                    If Len(refRoomID) > 0 Then
                        ' Check if referenced room exists
                        If Not HasRoomID(targetBook, refRoomID, foundSheet) Then
                            validationSheet.Cells(currentRow, 1).value = "Warning"
                            validationSheet.Cells(currentRow, 2).value = "Missing Room Reference"
                            validationSheet.Cells(currentRow, 3).value = _
                                "Sheet '" & roomSheet.name & "' references non-existent Room ID: '" & refRoomID & "'"
                            validationSheet.Cells(currentRow, 1).Font.Color = RGB(255, 165, 0)
                            currentRow = currentRow + 1
                        End If
                    End If
                Next cell
            End If
            
            ' Check "Doors To Room Alias" references
            On Error Resume Next
            Set refRange = roomSheet.Range(modConst.NAME_RANGE_DOORS_TO_ROOM_ALIAS)
            On Error GoTo 0
            
            If Not refRange Is Nothing Then
                For Each cell In refRange.Cells
                    refRoomAlias = Trim$(CStr(cell.value))
                    
                    If Len(refRoomAlias) > 0 Then
                        ' Check if referenced room exists
                        If Not HasRoomAlias(targetBook, refRoomAlias, foundSheet) Then
                            validationSheet.Cells(currentRow, 1).value = "Warning"
                            validationSheet.Cells(currentRow, 2).value = "Missing Room Reference"
                            validationSheet.Cells(currentRow, 3).value = _
                                "Sheet '" & roomSheet.name & "' references non-existent Room Alias: '" & refRoomAlias & "'"
                            validationSheet.Cells(currentRow, 1).Font.Color = RGB(255, 165, 0)
                            currentRow = currentRow + 1
                        End If
                    End If
                Next cell
            End If
        End If
    Next roomSheet
    
    CheckMissingRoomReferences = currentRow
End Function

' -----------------------------------------------------------------------------------
' Function  : CheckCircularDependencies
' Purpose   : Detects circular dependencies between rooms (e.g., Room A -> Room B -> Room A)
'
' Parameters:
'   targetBook        [Workbook]   - Workbook to check
'   validationSheet   [Worksheet]  - Sheet to write results to
'   startRow          [Long]       - Starting row for writing issues
'
' Returns   : Long - Next available row for writing
'
' Notes     :
'   - Uses depth-first search to detect cycles
'   - Reports the cycle path when found
' -----------------------------------------------------------------------------------
Private Function CheckCircularDependencies(ByVal targetBook As Workbook, _
    ByVal validationSheet As Worksheet, _
    ByVal startRow As Long) As Long
    Dim roomGraph As Scripting.Dictionary
    Dim roomSheet As Worksheet
    Dim currentRow As Long
    Dim roomID As String
    Dim visitedRooms As Scripting.Dictionary
    Dim pathStack As Collection
    Dim cycleFound As Boolean
    Dim cycleMessage As String
    
    currentRow = startRow
    Set roomGraph = New Scripting.Dictionary
    
    ' Build room dependency graph
    For Each roomSheet In targetBook.Worksheets
        If IsRoomSheet(roomSheet, roomID) Then
            If Not roomGraph.Exists(roomID) Then
                roomGraph.Add roomID, GetRoomDependencies(roomSheet)
            End If
        End If
    Next roomSheet
    
    ' Check each room for cycles using DFS
    Set visitedRooms = New Scripting.Dictionary
    Dim startRoomID As Variant
    
    For Each startRoomID In roomGraph.Keys
        Set pathStack = New Collection
        cycleFound = False
        cycleMessage = ""
        
        If DetectCycleDFS(CStr(startRoomID), roomGraph, visitedRooms, pathStack, cycleMessage) Then
            ' Cycle detected
            validationSheet.Cells(currentRow, 1).value = "Warning"
            validationSheet.Cells(currentRow, 2).value = "Circular Dependency"
            validationSheet.Cells(currentRow, 3).value = cycleMessage
            validationSheet.Cells(currentRow, 1).Font.Color = RGB(255, 165, 0)
            currentRow = currentRow + 1
        End If
    Next startRoomID
    
    CheckCircularDependencies = currentRow
End Function

' -----------------------------------------------------------------------------------
' Function  : GetRoomDependencies
' Purpose   : Retrieves all room references (dependencies) from a room sheet
'
' Parameters:
'   roomSheet  [Worksheet]  - Room sheet to analyze
'
' Returns   : Collection - Collection of referenced room IDs
' -----------------------------------------------------------------------------------
Private Function GetRoomDependencies(ByVal roomSheet As Worksheet) As Collection
    Dim dependencies As Collection
    Dim refRange As Range
    Dim cell As Range
    Dim refRoomID As String
    Dim foundSheet As Worksheet
    
    Set dependencies = New Collection
    
    ' Get references from "Doors To Room ID" range
    On Error Resume Next
    Set refRange = roomSheet.Range(modConst.NAME_RANGE_DOORS_TO_ROOM_ID)
    On Error GoTo 0
    
    If Not refRange Is Nothing Then
        For Each cell In refRange.Cells
            refRoomID = Trim$(CStr(cell.value))
            
            If Len(refRoomID) > 0 Then
                ' Verify the reference exists
                If HasRoomID(roomSheet.Parent, refRoomID, foundSheet) Then
                    On Error Resume Next
                    dependencies.Add refRoomID, refRoomID
                    On Error GoTo 0
                End If
            End If
        Next cell
    End If
    
    Set GetRoomDependencies = dependencies
End Function

' -----------------------------------------------------------------------------------
' Function  : DetectCycleDFS
' Purpose   : Performs depth-first search to detect cycles in room dependencies
'
' Parameters:
'   currentRoomID  [String]                  - Current room being checked
'   roomGraph      [Dictionary]              - Graph of room dependencies
'   visitedRooms   [Dictionary]              - Tracks visited rooms (global)
'   pathStack      [Collection]              - Current path being explored
'   cycleMessage   [String, ByRef]           - Message describing the cycle
'
' Returns   : Boolean - True if cycle detected, False otherwise
' -----------------------------------------------------------------------------------
Private Function DetectCycleDFS(ByVal currentRoomID As String, _
    ByVal roomGraph As Scripting.Dictionary, _
    ByVal visitedRooms As Scripting.Dictionary, _
    ByVal pathStack As Collection, _
    ByRef cycleMessage As String) As Boolean
    Dim i As Long
    Dim pathItem As Variant
    Dim dependencies As Collection
    Dim depRoomID As Variant
    
    ' Check if current room is already in the current path (cycle detected)
    For i = 1 To pathStack.count
        If pathStack(i) = currentRoomID Then
            ' Build cycle message
            cycleMessage = "Cycle detected: "
            For Each pathItem In pathStack
                cycleMessage = cycleMessage & pathItem & " -> "
            Next pathItem
            cycleMessage = cycleMessage & currentRoomID
            DetectCycleDFS = True
            Exit Function
        End If
    Next i
    
    ' Mark as visited for this session
    If visitedRooms.Exists(currentRoomID) Then
        DetectCycleDFS = False
        Exit Function
    End If
    
    ' Add to current path
    pathStack.Add currentRoomID
    
    ' Get dependencies and explore
    If roomGraph.Exists(currentRoomID) Then
        Set dependencies = roomGraph(currentRoomID)
        
        For Each depRoomID In dependencies
            If DetectCycleDFS(CStr(depRoomID), roomGraph, visitedRooms, pathStack, cycleMessage) Then
                DetectCycleDFS = True
                Exit Function
            End If
        Next depRoomID
    End If
    
    ' Remove from path (backtrack)
    pathStack.Remove pathStack.count
    
    ' Mark as fully processed
    If Not visitedRooms.Exists(currentRoomID) Then
        visitedRooms.Add currentRoomID, True
    End If
    
    DetectCycleDFS = False
End Function



