Attribute VB_Name = "modRooms"
' ===================================================================================
' Module    : modRooms
' Purpose   : Create, initialize and manage "Room" worksheets; aggregate list data.
'
' Dependencies: modProps, modUtil, modTags, modRanges, modSheets
' Notes     :
'   - Public API section exposes operations used by UI/other modules.
'   - Private Helpers contain internal utilities for this module only.
' ===================================================================================
Option Explicit
Option Private Module

' ===== Public API ==================================================================

' -----------------------------------------------------------------------------------
' Function  : AddRoom
' Purpose   : Clones required templates, ensures helper sheets exist, creates a new
'             Room sheet with the next sequential ID, initializes it, and activates it.
'
' Parameters:
'   objActWb              [Workbook] - Target workbook.
'   strNewName            [String]   - Name for the new Room sheet (already computed).
'   lngIdx                [Long]     - Numeric index used for the Room ID formatting
'   blnUpdateAggregations [Boolean]  - If True, UpdateLists is called after creation.
'
' Returns   : new Room Worksheet
'
' Notes     :
'   - Ensures dispatcher and lists sheets by copying them from RDDAddInWkBk if missing.
'   - Creates a new sheet from SHEET_ROOM_TEMPLATE and tags it with ROOM_SHEET_ID_TAG_NAME.
'   - Calls SetupRoom to wire controls/values; toggles HideOpMode during operations.
' -----------------------------------------------------------------------------------
Public Function AddRoom(ByVal objActWb As Workbook, ByRef strNewName As String, ByRef lngIdx As Long, Optional ByVal blnUpdateAggregations As Boolean = True) As Worksheet
    On Error GoTo errHandler

    Dim wksTmpl As Worksheet
    Dim wksTarget As Worksheet
    
    modUtil.HideOpMode True
            
    If Not modSheets.SheetCodeNameExists(modConst.SHEET_DISPATCHER, objActWb) And Not modTags.SheetWithTagExists(objActWb, SHEET_DISPATCHER) Then
    
        Set wksTmpl = RDDAddInWkBk.Worksheets(modConst.SHEET_DISPATCHER)
        wksTmpl.Visible = xlSheetVisible
        
        wksTmpl.Copy After:=objActWb.Sheets(objActWb.Sheets.Count)
        Set wksTarget = ActiveSheet 'objActWb.Sheets(objActWb.Sheets.Count)
        
        wksTarget.Visible = xlSheetHidden
        wksTarget.Name = "DO_NOT_DELETE"
        modProps.ClearAllCustomProperties wksTarget
        modTags.TagSheet wksTarget, SHEET_DISPATCHER
        
        Set wksTmpl = Nothing
        Set wksTarget = Nothing
    End If
            
    Set wksTmpl = RDDAddInWkBk.Worksheets(modConst.SHEET_ROOM_TEMPLATE)
    wksTmpl.Visible = xlSheetVisible
    
    wksTmpl.Copy After:=objActWb.Sheets(objActWb.Sheets.Count)
    Set wksTarget = ActiveSheet ' objActWb.Sheets(objActWb.Sheets.Count)
    
    wksTarget.Name = strNewName
    
    modProps.ClearAllCustomProperties wksTarget
    modTags.TagSheet wksTarget, ROOM_SHEET_ID_TAG_NAME, GetFormattedRoomID(lngIdx)
    
    SetupRoom wksTarget, lngIdx
    
    Set AddRoom = wksTarget
        
    If blnUpdateAggregations Then
        UpdateLists
    End If
    
    modUtil.HideOpMode False
    
    Exit Function
    
errHandler:
    LogError "AddRoom", Err.Number, Erl
    modUtil.HideOpMode False
End Function

' -----------------------------------------------------------------------------------
' Procedure : IsRoomSheet
' Purpose   : checks if active sheet is as room sheet
'
' Parameters:
'
'   wks  [Worksheet] - sheet to test
' strID  [String]    - (Optional, ByRef) returns the room ID if it is a room sheet.
'
'
' Returns   : [Boolean] - True if the sheet is a room sheet
'
' -----------------------------------------------------------------------------------
Public Function IsRoomSheet(ByRef wks As Worksheet, Optional ByRef strID As String = vbNullString) As Boolean
    Dim strValue As String
    If modTags.HasSheetTag(wks, ROOM_SHEET_ID_TAG_NAME, strID) Then
        IsRoomSheet = True
    End If
End Function

' -----------------------------------------------------------------------------------
' Function  : RemoveRoom
' Purpose   : Deletes the active Room sheet after verifying that no other Room sheet
'             references it; confirms with the user; updates aggregated lists.
'
' Parameters:
'   wks                    [Worksheet]   - Room sheet to delete (must not be Nothing).
'   blnUpdateAggregations  [Boolean]     - If True, UpdateLists is called after deletion.
'   colReferencingSheets   [Collection]  - (Optional, ByRef) receives referencing sheet names.
'
' Returns   : Boolean - True on success; False is not used (errors are raised/logged).
'
' Notes     :
'   - Collects Room sheets via modSheets.BuildDictFromSheetsByTag (excludes target sheet).
'   - Checks references via GetAllSheetsUsingRoomID.
'   - Uses HideOpMode during deletion; hands back referencing sheets if found.
' -----------------------------------------------------------------------------------
Public Function RemoveRoom(ByVal wks As Worksheet, _
    Optional ByVal blnUpdateAggregations As Boolean = True, _
    Optional ByRef colReferencingSheets As Collection) As Boolean
    On Error GoTo errHandler

    Dim wb As Workbook: Set wb = wks.Parent
    Dim strRoomID As String

    If wks Is Nothing Then
        Err.Raise vbObjectError + 513, "modRooms.RemoveRoom", "Argument 'wks' must not be Nothing."
    End If
    
    If Not IsRoomSheet(wks, strRoomID) Then
        Err.Raise vbObjectError + 514, "modRooms.RemoveRoom", "The provided sheet is not a Room sheet."
    End If
    
    ' Collect all Room-sheets except the active one
    Dim dicRooms As Object  ' Scripting.Dictionary
    Set dicRooms = modSheets.BuildDictFromSheetsByTag(wb, ROOM_SHEET_ID_TAG_NAME, wksExclude:=wks)
    
    ' Check references to the active room sheet in all other room sheets
    Dim colUsedIn As Collection
    Set colUsedIn = GetAllSheetsUsingRoomID(strRoomID, dicRooms)
    If Not colUsedIn Is Nothing Then
        If colUsedIn.Count > 0 Then
            ' Hand result back to caller for UI, then raise a error
            Set colReferencingSheets = colUsedIn
            Err.Raise vbObjectError + 515, _
                "modRooms.RemoveRoom", _
                "Room sheet cannot be deleted because it is referenced by other Room sheets."
        End If
    End If
        
    modUtil.HideOpMode True
    wks.Delete
    Set wks = Nothing
    
    If blnUpdateAggregations Then
        UpdateLists
    End If
    
    modUtil.HideOpMode False
    Exit Function
    
errHandler:
    LogError "RemoveRoom", Err.Number, Erl
End Function

' -----------------------------------------------------------------------------------
' Function  : GetNextRoomIndex
' Purpose   : Computes the next free numeric index by scanning existing Room sheets
'             and returning (max index + 1).
'
' Parameters:
'   wb      [Workbook] - Workbook to scan
'
' Returns   : Long - Next available Room index.
'
' Notes     :
'   - Detects Room sheets via ROOM_SHEET_ID_TAG_NAME and parses the ID.
'
' -----------------------------------------------------------------------------------
Public Function GetNextRoomIndex(ByVal wb As Workbook) As Long
    ' Returns next free numeric index based on existing taged Room* sheets
    Dim wks As Worksheet, lngNum As Long, lngMax As Long
    Dim strValue As String
    For Each wks In wb.Worksheets
        If modTags.HasSheetTag(wks, ROOM_SHEET_ID_TAG_NAME, strValue) Then
            lngNum = val(Mid$(CStr(strValue), Len(ROOM_SHEET_ID_TAG_VAL_PRE) + 1))
            If lngNum > lngMax Then lngMax = lngNum
        End If
    Next wks
    GetNextRoomIndex = lngMax + 1
End Function

' -----------------------------------------------------------------------------------
' Function  : HasRoomSheet
' Purpose   : Determine whether the workbook contains a Room sheet with the given ID.
'
' Parameters:
'   wb       [Workbook]              - Workbook to scan.
'   strRoomID[String]                - Room ID to search for (e.g., "R001").
'   r_wks    [Worksheet]             - (Optional ByRef) Receives the first matching sheet if found.
'
' Returns   : Boolean - True if found; otherwise False.
'
' Notes     :
' -----------------------------------------------------------------------------------
Public Function HasRoomSheet(ByVal wb As Workbook, ByVal strRoomID As String, Optional ByRef r_wks As Worksheet = Nothing) As Boolean
    Dim wks As Worksheet
    Dim strValue As String
    For Each wks In wb.Worksheets
        If modTags.HasSheetTag(wks, ROOM_SHEET_ID_TAG_NAME, strValue) Then
            If StrComp(strRoomID, strValue, vbBinaryCompare) = 0 Then
                Set r_wks = wks
                HasRoomSheet = True
                Exit Function
            End If
        End If
    Next wks
End Function

' -----------------------------------------------------------------------------------
' Function  : UpdateLists
' Purpose   : Aggregates Room IDs, Objects, and Scene IDs from all Room sheets into
'             the Lists sheet and updates the corresponding named ranges.
'
' Parameters: (none)
'
' Returns   : (none)
'
' Notes     :
'   - Room ID via IsRoomSheet (tag), Room Alias via NAME_CELL_ROOM_ALIAS.
'   - Collects object names via named ranges (pickupable/multistate/touchable).
' -----------------------------------------------------------------------------------
Public Sub UpdateLists()
    On Error GoTo errHandler
    Dim dicRooms As Object: Set dicRooms = CreateObject("Scripting.Dictionary")
    Dim dicObjects As Object:  Set dicObjects = CreateObject("Scripting.Dictionary")
    Dim dicScenes As Object: Set dicScenes = CreateObject("Scripting.Dictionary")
    
    Dim dicExisting As Object: Set dicExisting = CreateObject("Scripting.Dictionary")
    Dim wks As Worksheet
    Dim wbActive As Workbook
    Dim strRoomID As String
    
    Set wbActive = ActiveWorkbook
    
    For Each wks In wbActive.Worksheets
        If modRooms.IsRoomSheet(wks, strRoomID) Then
            ' Room ID
            If Len(strRoomID) = 0 Then GoTo SkipWksIteration
            If Len(strRoomID) > 0 Then dicRooms(strRoomID) = True
            
            ' Room Alias
            Dim strRoomAlias As String: strRoomAlias = modLists.GetNamedOrHeaderValue(wks, NAME_CELL_ROOM_ALIAS, Array("Room Alias", NAME_CELL_ROOM_ALIAS))
            If Len(strRoomID) > 0 Then dicRooms(strRoomID) = strRoomAlias
            
            ' Scene ID
            Dim strSceneId As String: strSceneId = modLists.GetNamedOrHeaderValue(wks, NAME_CELL_SCENE_ID, Array("Scene ID", NAME_CELL_SCENE_ID))
            If Len(strSceneId) > 0 Then dicScenes(strSceneId) = True
            
            modLists.CollectNamedRangeValues wks, NAME_RANGE_PICKUPABLE_OBJ, dicObjects
            modLists.CollectNamedRangeValues wks, NAME_RANGE_MULTISTATE_OBJ, dicObjects
            modLists.CollectNamedRangeValues wks, NAME_RANGE_TOUCHABLE_OBJ, dicObjects
            
        End If
SkipWksIteration:
    Next wks
    
    Dim wksLists As Worksheet: Set wksLists = modSheets.GetSheetByCodeName(SHEET_DISPATCHER)
    
    If wksLists Is Nothing Then: Set wksLists = modTags.GetSheetByTag(wbActive, SHEET_DISPATCHER)
    
    If Not wksLists Is Nothing Then

        wksLists.Columns(LISTS_COL_ROOM_ID).Clear
        wksLists.Columns(LISTS_COL_ROOM_ALIAS).Clear
        
        wksLists.Cells(1, LISTS_COL_ROOM_ID).Value = LISTS_HEADER_ROOM_ID
        wksLists.Cells(1, LISTS_COL_ROOM_ALIAS).Value = LISTS_HEADER_ROOM_ALIAS
        wksLists.Cells(1, LISTS_COL_OBJECTS).Value = LISTS_HEADER_OBJECTS
        wksLists.Cells(1, LISTS_COL_SCENE_ID).Value = LISTS_HEADER_SCENE_ID
        wksLists.Range("A1:ZZ1").Font.Bold = True
    
        ' Write Room IDs sorted, must always be rewritten, as it is related to the room pages
        modLists.WriteDictSetToColumn wksLists, dicRooms, 2, LISTS_COL_ROOM_ID, True
        
        ' Append only missing Object names
        modLists.CollectColumnValues wksLists, Array(LISTS_HEADER_OBJECTS), dicExisting
        modLists.AppendMissingDictKeysToColumn wksLists, LISTS_COL_OBJECTS, dicExisting, dicObjects
        
        ' Append only missing Scene IDs
        modLists.CollectColumnValues wksLists, Array(LISTS_HEADER_SCENE_ID), dicExisting
        modLists.AppendMissingDictKeysToColumn wksLists, LISTS_COL_SCENE_ID, dicExisting, dicScenes
    
        modLists.UpdateNamedListRange NAME_LIST_ROOM_IDS, wksLists, LISTS_COL_ROOM_ID
        modLists.UpdateNamedListRange NAME_LIST_OBJECTS, wksLists, LISTS_COL_OBJECTS
        modLists.UpdateNamedListRange NAME_LIST_SCENE_IDS, wksLists, LISTS_COL_SCENE_ID
    End If
    Exit Sub
    
errHandler:
    LogError "UpdateLists", Err.Number, Erl
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
    On Error GoTo errHandler
    Dim dicRooms As Object:   Set dicRooms = CreateObject("Scripting.Dictionary")
    Dim dicObjects As Object: Set dicObjects = CreateObject("Scripting.Dictionary")
    Dim dicScenes As Object:  Set dicScenes = CreateObject("Scripting.Dictionary")
    
    Dim wks As Worksheet
    Dim wbActive As Workbook
    Dim strRoomID As String
    
    Set wbActive = ActiveWorkbook
    ' collect datas
    For Each wks In wbActive.Worksheets
        If modRooms.IsRoomSheet(wks, strRoomID) Then
            If Len(strRoomID) = 0 Then GoTo SkipWksIteration
            If Len(strRoomID) > 0 Then dicRooms(strRoomID) = True
            
            ' Room Alias
            Dim strRoomAlias As String: strRoomAlias = modLists.GetNamedOrHeaderValue(wks, NAME_CELL_ROOM_ALIAS, Array("Room Alias", NAME_CELL_ROOM_ALIAS))
            If Len(strRoomID) > 0 Then dicRooms(strRoomID) = strRoomAlias
            
            Dim strSceneId As String
            strSceneId = modLists.GetNamedOrHeaderValue(wks, NAME_CELL_SCENE_ID, Array("Scene ID", NAME_CELL_SCENE_ID))
            If Len(strSceneId) > 0 Then dicScenes(strSceneId) = True
            
            modLists.CollectNamedRangeValues wks, NAME_RANGE_PICKUPABLE_OBJ, dicObjects
            modLists.CollectNamedRangeValues wks, NAME_RANGE_MULTISTATE_OBJ, dicObjects
            modLists.CollectNamedRangeValues wks, NAME_RANGE_TOUCHABLE_OBJ, dicObjects
        End If
SkipWksIteration:
    Next wks
    
    Dim wksLists As Worksheet: Set wksLists = modSheets.GetSheetByCodeName(SHEET_DISPATCHER)
    
    If wksLists Is Nothing Then: Set wksLists = modTags.GetSheetByTag(wbActive, SHEET_DISPATCHER)
    
    If Not wksLists Is Nothing Then
        ' Clear target columns
        wksLists.Columns(LISTS_COL_ROOM_ID).Clear     ' Room IDs
        wksLists.Columns(LISTS_COL_ROOM_ALIAS).Clear
        wksLists.Columns(LISTS_COL_SCENE_ID).Clear    ' Scene IDs
        wksLists.Columns(LISTS_COL_OBJECTS).Clear     ' Objects
        
        ' Headers
        wksLists.Cells(1, LISTS_COL_ROOM_ID).Value = LISTS_HEADER_ROOM_ID
        wksLists.Cells(1, LISTS_COL_ROOM_ALIAS).Value = LISTS_HEADER_ROOM_ALIAS
        wksLists.Cells(1, LISTS_COL_OBJECTS).Value = LISTS_HEADER_OBJECTS
        wksLists.Cells(1, LISTS_COL_SCENE_ID).Value = LISTS_HEADER_SCENE_ID
        wksLists.Range("A1:ZZ1").Font.Bold = True
    
        ' Write sorted values
        modLists.WriteDictSetToColumn wksLists, dicRooms, 2, LISTS_COL_ROOM_ID, True
        modLists.WriteDictSetToColumn wksLists, dicObjects, 2, LISTS_COL_OBJECTS
        modLists.WriteDictSetToColumn wksLists, dicScenes, 2, LISTS_COL_SCENE_ID
        
        ' Update named ranges
        modLists.UpdateNamedListRange NAME_LIST_ROOM_IDS, wksLists, LISTS_COL_ROOM_ID
        modLists.UpdateNamedListRange NAME_LIST_OBJECTS, wksLists, LISTS_COL_OBJECTS
        modLists.UpdateNamedListRange NAME_LIST_SCENE_IDS, wksLists, LISTS_COL_SCENE_ID
    End If
    
    Exit Sub
    
errHandler:
    LogError "UpdateLists", Err.Number, Erl
End Sub

' -----------------------------------------------------------------------------------
' Function  : GetFormattedRoomID
' Purpose   : Build a formatted Room ID from the numeric index using a prefix.
'
' Parameters:
'   lngIdx [Long] - Numeric index.
'
' Returns   : String - e.g., R001 (depends on ROOM_SHEET_ID_TAG_VAL_PRE).
'
' Notes     :
' -----------------------------------------------------------------------------------
Public Function GetFormattedRoomID(ByVal lngIdx As Long) As String
    GetFormattedRoomID = ROOM_SHEET_ID_TAG_VAL_PRE & Format(lngIdx, "000")
End Function

' -----------------------------------------------------------------------------------
' Procedure : ApplyParallaxRangeCover
' Purpose   : Ensures and toggles a named range "cover" according to a dropdown state.
'
' Parameters:
'   wks [Worksheet] - Target Room worksheet.
'
' Returns   : (none)
' -----------------------------------------------------------------------------------
Public Sub ApplyParallaxRangeCover(wks As Worksheet)
    modRangeCover.EnsureRangeCover wks, wks.Range(NAME_RANGE_ADD_PARALLAX), ROOM_ADD_PARALLAX_SET_COVER_NAME
    modRangeCover.ApplyCoverVisibilityByDropdown wks, NAME_CELL_PARALLAX, ROOM_ADD_PARALLAX_SET_HIDE_TOKEN, ROOM_ADD_PARALLAX_SET_COVER_NAME
End Sub

' ===== Private Helpers =============================================================

' -----------------------------------------------------------------------------------
' Function  : SetupRoom
' Purpose   : Initializes a newly created Room sheet: sets RoomID, removes stale
'             name links, and wires the "insert room picture" button action.
'
' Parameters:
'   wks       [Worksheet]  - Target Room worksheet to initialize.
'   lngIdx    [Long]       - Numeric index used to format the Room ID cell value.
'
' Returns   : none
'
' Notes     :
' -----------------------------------------------------------------------------------
Private Sub SetupRoom(wks As Worksheet, ByVal lngIdx As Long)
    Dim shpBtn As Shape
    Dim wksDisp As Worksheet
    Dim rngData As Range
    
    'Set 'RoomID' named cell on the template
    wks.Range(modConst.NAME_CELL_ROOM_ID).Value = GetFormattedRoomID(lngIdx)
    'Set 'RoomAlias" named Cell on the Template
    wks.Range(modConst.NAME_CELL_ROOM_ALIAS).Value = "r_" & GetCleanRoomAlias(wks.Name)
    
    ' remove wrong links
    Dim nm As Name
    For Each nm In wks.Parent.Names
        If InStr(nm.RefersTo, "[" & RDDAddInWkBk.Name & "]") > 0 Then
            nm.Delete
        End If
    Next
    
    ' update button link
    Set wksDisp = modSheets.GetSheetByCodeName(modConst.SHEET_DISPATCHER)
    Set shpBtn = wks.Shapes(modConst.BTN_INSERT_ROOM_PICTURE)
    shpBtn.OnAction = modConst.MACRO_BTN_INSERT_PICTURE
        
    ' Add data validations
    
    ' Type
    Set rngData = wks.Range(NAME_RANGE_PUZZLES_TYPE)
    ApplyListValidation rngData, NAME_LIST_PUZZLE_TYPES, "Type", "Choose a type from the list."

    
    
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
Private Sub ApplyListValidation(ByVal target As Range, ByVal nameRef As String, _
                                ByVal title As String, ByVal msg As String)
    With target.Validation
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
'   strRoomID  [String]  - Room sheet id to search for.
'   dicRooms   [Object]  - Dictionary of Room worksheets to inspect (values = Worksheet).
'
' Returns   : Collection - Sheet names that reference strRoomID.
'
' Notes     :
' -----------------------------------------------------------------------------------
Private Function GetAllSheetsUsingRoomID(ByVal strRoomID As String, ByVal dicRooms As Object) As Collection
    ' Returns a collection of sheet names that reference strRoomID
    Dim col As New Collection
    Dim vKey As Variant, wks As Worksheet
    Dim rng As Range
    
    For Each vKey In dicRooms.Keys
        Set wks = dicRooms(vKey)
        
        ' Framed area "DOORS TO..."
        Set rng = wks.Range(NAME_RANGE_DOORS_TO) 'FindFramedRangeByHeading(wks, "DOORS TO", False)
        If Not rng Is Nothing Then
            If modRanges.RangeHasValue(rng, strRoomID, True, False) Then
                col.Add wks.Name
                GoTo NextSheet
            End If
        End If
                
NextSheet:
        Set rng = Nothing
    Next
    
    Set GetAllSheetsUsingRoomID = col
End Function

' -----------------------------------------------------------------------------------
' Function  : GetCleanRoomAlias
' Purpose   : Produces a simplified alias by removing specific punctuation from a name.
'
' Parameters:
'   strInput [String] - Source name to normalize.
'
' Returns   : String - Cleaned alias string.
' -----------------------------------------------------------------------------------
Private Function GetCleanRoomAlias(ByVal strInput As String) As String
    Dim arrRemove() As Variant
    Dim i As Long

    ' Liste der zu entfernenden Zeichen
    arrRemove = Array(" ", "-", ".", "(", ")", ":", "/", "'")

    ' Alle Zeichen durch leeren String ersetzen
    For i = LBound(arrRemove) To UBound(arrRemove)
        strInput = Replace(strInput, arrRemove(i), "")
    Next i

    GetCleanRoomAlias = strInput
End Function
