Attribute VB_Name = "modRooms"
' modRooms
Option Explicit
Option Private Module

' ================================================
' --- Ribbon Callback Targets ---
' ================================================

' -----------------------------------------------------------------------------------
' Function  : AddRoom
' Purpose   : Clones required templates, ensures helper sheets exist, creates a new
'             Room sheet with the next sequential ID, initializes it, and activates it.
'
' Parameters:
'   (none)
'
' Returns   :
'
' Notes     :
'   - Ensures dispatcher and lists sheets by copying them from RDDAddInWkBk if missing.
'   - Creates a new sheet from SHEET_ROOM_TEMPLATE; name = ROOM_SHEET_PREFIX & index.
'   - Calls SetupRoom to wire controls/values; toggles HideOpMode during operations.
' -----------------------------------------------------------------------------------
Public Sub AddRoom(ByRef strNewName As String, ByRef lngIdx As Long)
    Dim wksTmpl As Worksheet, wksTarget As Worksheet
    
    Dim objActWks As Worksheet: Set objActWks = ActiveSheet
    Dim objActWb As Workbook: Set objActWb = ActiveWorkbook
    
    Application.StatusBar = False
    
    modUtil.HideOpMode True
        
    modMain.EnsureWorkbookIsTagged objActWb
    
    If Not modSheets.SheetCodeNameExists(modConst.SHEET_DISPATCHER, objActWb) And Not modTags.SheetWithTagExists(objActWb, SHEET_DISPATCHER) Then
        Set wksTmpl = RDDAddInWkBk.Worksheets(modConst.SHEET_DISPATCHER)
        wksTmpl.Visible = xlSheetVisible
        wksTmpl.Copy After:=objActWb.Sheets(objActWb.Sheets.Count)
        Set wksTarget = ActiveSheet
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
    Set wksTarget = ActiveSheet
    wksTarget.Name = strNewName
    modProps.ClearAllCustomProperties wksTarget
    modTags.TagSheet wksTarget, ROOM_SHEET_ID_TAG_NAME, GetFormattedRoomID(lngIdx)
    
    SetupRoom wksTarget, lngIdx
    
    modUtil.HideOpMode False
    
    Application.Goto wksTarget.Range("A1"), True
End Sub

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
'
' Behavior  : -
' Notes     : -
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
'   (none)
'
' Returns   :
'
' Notes     :
'   - Gathers other Room sheets via BuildDictFromSheetsByName (excludes active).
'   - Checks references using GetSheetsUsingRoomName (doors / RoomID column).
'   - Shows blocking messages for references or user cancellation.
'   - Toggles HideOpMode; calls UpdateLists after deletion.
' -----------------------------------------------------------------------------------
Public Sub RemoveRoom()
    Dim wksActive As Worksheet: Set wksActive = ActiveSheet
    
    Application.StatusBar = False
    
    If Not IsRoomSheet(wksActive) Then
        MsgBox "Active sheet is not a 'Room' sheet.", vbInformation
        Exit Sub
    End If
    
    ' Collect all Room-sheets except the active one
    Dim dicRooms As Object  ' Scripting.Dictionary
    Set dicRooms = BuildDictFromSheetsByName(ThisWorkbook, ROOM_SHEET_PREFIX, SNMM_Prefix, wksExclude:=wksActive)
    
    ' Check references to the active room sheet in all other room sheets
    Dim colUsedIn As Collection
    Set colUsedIn = GetSheetsUsingRoomName(wksActive.Name, dicRooms)
    
    If colUsedIn.Count > 0 Then
        Dim strList As String
        Dim strMsgPart As String
        strList = JoinCollection(colUsedIn, ", ")
        If colUsedIn.Count > 1 Then
            strMsgPart = "worksheets"
        Else
            strMsgPart = "worksheet"
        End If
        
        MsgBox "The worksheet is used in the following " & strMsgPart & " and cannot be deleted.: " & vbNewLine & strList, vbCritical
        Exit Sub
    End If
    
    ' No references ? ask for confirmation and delete
    Dim response As VbMsgBoxResult
    response = MsgBox("Are you sure you want to delete the sheet '" & wksActive.Name & "'?" & vbCrLf & _
                      "This action cannot be undone.", vbYesNo + vbExclamation, "Confirm Sheet Deletion")
    If response <> vbYes Then
        Application.StatusBar = "Deletion cancelled."
        Exit Sub
    End If
    
    modUtil.HideOpMode True
    wksActive.Delete
    Set wksActive = Nothing
    UpdateLists
    modUtil.HideOpMode True
End Sub

' -----------------------------------------------------------------------------------
' Function  : SetupRoom
' Purpose   : Initializes a newly created Room sheet: sets RoomID, removes stale
'             name links, and wires the "insert room picture" button action.
'
' Parameters:
'   wks       [Worksheet]  - Target Room worksheet to initialize.
'   lngIdx    [Long]
'
' Returns   :
'
' Notes     :
'   - Sets named cell NAME_CELL_ROOM_ID to the sheet ID.
'   - Deletes workbook Names that still refer to the add-in workbook.
'   - Sets button OnAction to MACRO_BTN_INSERT_PICTURE via BTN_INSERT_ROOM_PICTURE.
' -----------------------------------------------------------------------------------
Private Sub SetupRoom(wks As Worksheet, ByVal lngIdx As Long)
    Dim shpBtn As Shape
    Dim wksDisp As Worksheet
    
    'Set 'RoomID' named cell on the template
    wks.Range(modConst.NAME_CELL_ROOM_ID).Value = GetFormattedRoomID(lngIdx)
    
    ' remove wrong links
    Dim nm As Name
    For Each nm In wks.Parent.names
        If InStr(nm.RefersTo, "[" & RDDAddInWkBk.Name & "]") > 0 Then
            nm.Delete
        End If
    Next
    
    ' update button link
    Set wksDisp = modSheets.GetSheetByCodeName(modConst.SHEET_DISPATCHER)
    Set shpBtn = wks.Shapes(modConst.BTN_INSERT_ROOM_PICTURE)
    shpBtn.OnAction = modConst.MACRO_BTN_INSERT_PICTURE
        
    ' Add data validations later if desired
End Sub

' -----------------------------------------------------------------------------------
' Function  : GetNextRoomIndex
' Purpose   : Computes the next free numeric index by scanning existing Room sheets
'             and returning (max index + 1).
'
' Parameters:
'   wb      [Workbook]
'
' Returns   : Long - Next available Room index.
'
' Notes     :
'   - Detects Room sheets via ROOM_SHEET_ID_TAG.
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
' Parameters:
'   (none)
'
' Returns   :
'
' Notes     :
'   - Uses GetNamedOrHeaderValue and CollectColumnBlockGroupValues per Room sheet.
'   - Writes headers and sorted Room IDs; appends only missing Objects/Scenes.
'   - Updates NAME_LIST_ROOM_IDS / NAME_LIST_OBJECTS / NAME_LIST_SCENE_IDS.
' -----------------------------------------------------------------------------------
Public Sub UpdateLists()
    Dim dicRooms As Object: Set dicRooms = CreateObject("Scripting.Dictionary")
    Dim dicObjects As Object:  Set dicObjects = CreateObject("Scripting.Dictionary")
    Dim dicScenes As Object: Set dicScenes = CreateObject("Scripting.Dictionary")
    
    Dim dicExisting As Object: Set dicExisting = CreateObject("Scripting.Dictionary")
    Dim wks As Worksheet
    Dim wbActive As Workbook
    
    Set wbActive = ActiveWorkbook
    For Each wks In wbActive.Worksheets
        If Left$(wks.Name, Len(ROOM_SHEET_PREFIX)) = ROOM_SHEET_PREFIX Then
            On Error Resume Next
            Dim strRoomID As String: strRoomID = Trim$(CStr(wks.Range(modConst.NAME_CELL_ROOM_ID).Value))
            On Error GoTo 0
            If Len(strRoomID) = 0 Then strRoomID = wks.Name
            If Len(strRoomID) > 0 Then dicRooms(strRoomID) = True
            
            Dim strSceneId As String: strSceneId = modLists.GetNamedOrHeaderValue(wks, NAME_CELL_SCENE_ID, Array("Scene ID", NAME_CELL_SCENE_ID, "Szene ID"))
            If Len(strSceneId) > 0 Then dicScenes(strSceneId) = True
            
            Call modLists.CollectColumnBlockGroupValues(wks, ROOM_OBJ_GROUP_HEADER_ROW, _
                ROOM_OBJ_GROUP_END_ROW, _
                Array(ROOM_HDR_NM_PICKUPABLE_OBJ, ROOM_HDR_NM_MULTISTATE_OBJ, ROOM_HDR_NM_TOUCHABLE_OBJ), _
                ROOM_OBJ_GROUP_CATEGORY_COLUMN_WIDTH, dicObjects)
        End If
    Next wks
    
    Dim wksLists As Worksheet: Set wksLists = modSheets.GetSheetByCodeName(SHEET_DISPATCHER)
    
    If Not wksLists Is Nothing Then

        wksLists.Columns(LISTS_COL_ROOM_ID).Clear
        
        wksLists.Cells(1, LISTS_COL_ROOM_ID).Value = LISTS_HEADER_ROOM_ID
        wksLists.Cells(1, LISTS_COL_OBJECTS).Value = LISTS_HEADER_OBJECTS
        wksLists.Cells(1, LISTS_COL_SCENE_ID).Value = LISTS_HEADER_SCENE_ID
        wksLists.Range("A1:ZZ1").Font.Bold = True
    
        ' --- Write Room IDs sorted, must always be rewritten, as it is related to the room pages
        modLists.WriteDictSetToColumn wksLists, dicRooms, 2, LISTS_COL_ROOM_ID
        
        ' --- Append only missing Object names
        modLists.CollectColumnValues wksLists, Array(LISTS_HEADER_OBJECTS), dicExisting
        modLists.AppendMissingDictKeysToColumn wksLists, LISTS_COL_OBJECTS, dicExisting, dicObjects
        
        ' --- Append only missing Scene IDs
        modLists.CollectColumnValues wksLists, Array(LISTS_HEADER_SCENE_ID), dicExisting
        modLists.AppendMissingDictKeysToColumn wksLists, LISTS_COL_SCENE_ID, dicExisting, dicScenes
    
        modLists.UpdateNamedListRange NAME_LIST_ROOM_IDS, wksLists, LISTS_COL_ROOM_ID
        modLists.UpdateNamedListRange NAME_LIST_OBJECTS, wksLists, LISTS_COL_OBJECTS
        modLists.UpdateNamedListRange NAME_LIST_SCENE_IDS, wksLists, LISTS_COL_SCENE_ID
    End If
End Sub

' -----------------------------------------------------------------------------------
' Function  : SyncLists
' Purpose   : Aggregates Room IDs, Objects, and Scene IDs from all Room sheets and
'             writes them into the Lists sheet. Clears the three target columns,
'             writes headers, outputs sorted values, and updates named ranges.
'
' Parameters:
'   (none)
'
' Returns   :
'
' Notes     :
'   - Scans worksheets whose names start with ROOM_SHEET_PREFIX.
'   - Reads RoomID from NAME_ROOM_ID (fallback: sheet name).
'   - Reads Scene ID via GetNamedOrHeaderValue.
'   - Collects object names via CollectColumnBlockGroupValues.
'   - Writes into sheet SHEET_LISTS if present; updates NAME_LIST_* named ranges.
' -----------------------------------------------------------------------------------
Public Sub SyncLists()
    Dim dicRooms As Object:   Set dicRooms = CreateObject("Scripting.Dictionary")
    Dim dicObjects As Object: Set dicObjects = CreateObject("Scripting.Dictionary")
    Dim dicScenes As Object:  Set dicScenes = CreateObject("Scripting.Dictionary")
    
    Dim wks As Worksheet
    Dim wbActive As Workbook
    
    Set wbActive = ActiveWorkbook
    ' collect datas
    For Each wks In wbActive.Worksheets
        If Left$(wks.Name, Len(ROOM_SHEET_PREFIX)) = ROOM_SHEET_PREFIX Then
            On Error Resume Next
            Dim strRoomID As String: strRoomID = Trim$(CStr(wks.Range(modConst.NAME_CELL_ROOM_ID).Value))
            On Error GoTo 0
            If Len(strRoomID) = 0 Then strRoomID = wks.Name
            If Len(strRoomID) > 0 Then dicRooms(strRoomID) = True
            
            Dim strSceneId As String
            strSceneId = modLists.GetNamedOrHeaderValue(wks, NAME_CELL_SCENE_ID, Array("Scene ID", NAME_CELL_SCENE_ID, "Szene ID"))
            If Len(strSceneId) > 0 Then dicScenes(strSceneId) = True
            
            Call modLists.CollectColumnBlockGroupValues( _
                wks, ROOM_OBJ_GROUP_HEADER_ROW, _
                ROOM_OBJ_GROUP_END_ROW, _
                Array(ROOM_HDR_NM_PICKUPABLE_OBJ, ROOM_HDR_NM_MULTISTATE_OBJ, ROOM_HDR_NM_TOUCHABLE_OBJ), _
                ROOM_OBJ_GROUP_CATEGORY_COLUMN_WIDTH, dicObjects)
        End If
    Next wks
    
    Dim wksLists As Worksheet: Set wksLists = modSheets.GetSheetByCodeName(SHEET_DISPATCHER)
    
    If Not wksLists Is Nothing Then
        ' Clear target columns
        wksLists.Columns(LISTS_COL_ROOM_ID).Clear     ' Room IDs
        wksLists.Columns(LISTS_COL_SCENE_ID).Clear    ' Scene IDs
        wksLists.Columns(LISTS_COL_OBJECTS).Clear     ' Objects
        
        ' Headers
        wksLists.Cells(1, LISTS_COL_ROOM_ID).Value = LISTS_HEADER_ROOM_ID
        wksLists.Cells(1, LISTS_COL_OBJECTS).Value = LISTS_HEADER_OBJECTS
        wksLists.Cells(1, LISTS_COL_SCENE_ID).Value = LISTS_HEADER_SCENE_ID
        wksLists.Range("A1:ZZ1").Font.Bold = True
    
        ' Write sorted values
        WriteDictSetToColumn wksLists, dicRooms, 2, LISTS_COL_ROOM_ID
        WriteDictSetToColumn wksLists, dicObjects, 2, LISTS_COL_OBJECTS
        WriteDictSetToColumn wksLists, dicScenes, 2, LISTS_COL_SCENE_ID
        
        ' Update named ranges
        UpdateNamedListRange NAME_LIST_ROOM_IDS, wksLists, LISTS_COL_ROOM_ID
        UpdateNamedListRange NAME_LIST_OBJECTS, wksLists, LISTS_COL_OBJECTS
        UpdateNamedListRange NAME_LIST_SCENE_IDS, wksLists, LISTS_COL_SCENE_ID
    End If
End Sub

' -----------------------------------------------------------------------------------
' Function  : GetSheetsUsingRoomName
' Purpose   : Finds all Room sheets (from a provided dictionary) that reference the
'             given room name either inside the "DOORS TO..." framed range or in the
'             "RoomID" column.
'
' Parameters:
'   sRoomName  [String]  - Room sheet name to search for.
'   dicRooms   [Object]  - Dictionary of Room worksheets to inspect (values = Worksheet).
'
' Returns   : Collection - Sheet names that reference sRoomName.
'
' Notes     :
' -----------------------------------------------------------------------------------
Private Function GetSheetsUsingRoomName(ByVal sRoomName As String, ByVal dicRooms As Object) As Collection
    ' Returns a collection of sheet names that reference sRoomName
    Dim col As New Collection
    Dim vKey As Variant, wks As Worksheet
    Dim rng As Range
    
    For Each vKey In dicRooms.Keys
        Set wks = dicRooms(vKey)
        
        ' Framed area "DOORS TO..."
        Set rng = wks.Range(NAME_RANGE_DOORS_TO) 'FindFramedRangeByHeading(wks, "DOORS TO", False)
        If Not rng Is Nothing Then
            If RangeHasValue(rng, sRoomName, True, False) Then
                col.Add wks.Name
                GoTo NextSheet
            End If
        End If
        
        ' Column "RoomID" in Framed area "PUZZLES"
        Set rng = wks.Range(NAME_RANGE_PUZZLES_ROOM_ID)  'GetColumnRangeByHeader(wks, "RoomID", True, True)
        If Not rng Is Nothing Then
            If RangeHasValue(rng, sRoomName, True, False) Then
                col.Add wks.Name
                GoTo NextSheet
            End If
        End If
        
NextSheet:
        Set rng = Nothing
    Next
    
    Set GetSheetsUsingRoomName = col
End Function

Public Function GetFormattedRoomID(ByVal lngIdx As Long) As String
    GetFormattedRoomID = ROOM_SHEET_ID_TAG_VAL_PRE & Format(lngIdx, "000")
End Function


