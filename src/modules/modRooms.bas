Attribute VB_Name = "modRooms"
' modRooms
Option Explicit
Option Private Module

Public gblnRoomSheetChanged As Boolean

' Clones a hidden "Room_Template" sheet, assigns next RoomID, wires validations
Public Sub AddRoom()
    Dim wksTmpl As Worksheet, wksTarget As Worksheet, lngIdx As Long, strNewName As String
    
    HideOpMode True
        
    If Not modUtil.SheetCodeNameExists(modConst.SHEET_DISPATCHER) Then
        Set wksTmpl = RDDAddInWkBk.Worksheets(modConst.SHEET_DISPATCHER)
        wksTmpl.Visible = xlSheetVisible
        wksTmpl.Copy After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count)
        Set wksTarget = ActiveSheet
        wksTarget.Visible = xlSheetHidden
        wksTarget.name = "DO_NOT_DELETE"
        Set wksTmpl = Nothing
        Set wksTarget = Nothing
    End If
    
    If Not modUtil.SheetCodeNameExists(modConst.SHEET_LISTS) Then
        Set wksTmpl = RDDAddInWkBk.Worksheets(modConst.SHEET_LISTS)
        wksTmpl.Visible = xlSheetVisible
        wksTmpl.Copy After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count)
        wksTmpl.Visible = xlSheetHidden
        Set wksTmpl = Nothing
    End If
        
    Set wksTmpl = RDDAddInWkBk.Worksheets(modConst.SHEET_ROOM_TEMPLATE)
    lngIdx = NextRoomIndex()
    strNewName = modConst.ROOM_SHEET_PREFIX & Format(lngIdx, "000")
    wksTmpl.Visible = xlSheetVisible
    wksTmpl.Copy After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count)
    Set wksTarget = ActiveSheet
    wksTmpl.Visible = xlSheetHidden
    wksTarget.name = strNewName
    
    SetupRoom wksTarget
    
    HideOpMode False
    
    Application.GoTo wksTarget.Range("A1"), True
End Sub


Private Sub SetupRoom(wks As Worksheet)
    Dim shpBtn As Shape
    Dim wksDisp As Worksheet
    
    'Set 'RoomID' named cell on the template
    wks.Range(modConst.NAME_ROOM_ID).Value = wks.name
    
    ' remove wrong links
    Dim nm As name
    For Each nm In wks.Parent.Names
        If InStr(nm.RefersTo, "[" & RDDAddInWkBk.name & "]") > 0 Then
            nm.Delete
        End If
    Next
    
    ' update button link
    Set wksDisp = modUtil.GetSheetByCodeName(modConst.SHEET_DISPATCHER)
    Set shpBtn = wks.Shapes(modConst.BTN_INSERT_ROOM_PICTURE)
    shpBtn.OnAction = modConst.MACRO_BTN_INSERT_PICTURE
        
    ' Add data validations later if desired
End Sub

Private Function NextRoomIndex() As Long
    ' Returns next free numeric index based on existing Room* sheets
    Dim wks As Worksheet, lngNum As Long, lngMax As Long
    For Each wks In ActiveWorkbook.Worksheets
        If Left$(wks.name, Len(ROOM_SHEET_PREFIX)) = ROOM_SHEET_PREFIX Then
            lngNum = val(Mid$(wks.name, 5))
            If lngNum > lngMax Then lngMax = lngNum
        End If
    Next wks
    NextRoomIndex = lngMax + 1
End Function

Public Sub RemoveRoom()
    Dim wksActive As Worksheet
    
    Set wksActive = ActiveSheet
    If Left$(wksActive.name, Len(ROOM_SHEET_PREFIX)) = ROOM_SHEET_PREFIX Then
        Dim response As VbMsgBoxResult
        response = MsgBox("Are you sure you want to delete the sheet '" & wksActive.name & "'?" & vbCrLf & _
                          "This action cannot be undone.", vbYesNo + vbExclamation, "Confirm Sheet Deletion")

        If response = vbYes Then
            HideOpMode True
            wksActive.Delete
            Set wksActive = Nothing
            Call UpdateLists
            HideOpMode True
        Else
            MsgBox "Deletion cancelled.", vbInformation
        End If
    Else
        MsgBox "Active sheet is not a 'Room' sheet.", vbInformation
    End If
End Sub

' ==== Lists aggregation ====
Public Sub UpdateLists()
    Dim dicRooms As Object: Set dicRooms = CreateObject("Scripting.Dictionary")
    Dim dicObjects As Object:  Set dicObjects = CreateObject("Scripting.Dictionary")
    Dim dicScenes As Object: Set dicScenes = CreateObject("Scripting.Dictionary")
    
    Dim dicExisting As Object: Set dicExisting = CreateObject("Scripting.Dictionary")
    Dim wks As Worksheet
    Dim wbActive As Workbook
    
    Set wbActive = ActiveWorkbook
    For Each wks In wbActive.Worksheets
        If Left$(wks.name, Len(ROOM_SHEET_PREFIX)) = ROOM_SHEET_PREFIX Then
            On Error Resume Next
            Dim strRoomId As String: strRoomId = Trim$(CStr(wks.Range(modConst.NAME_ROOM_ID).Value))
            On Error GoTo 0
            If Len(strRoomId) = 0 Then strRoomId = wks.name
            If Len(strRoomId) > 0 Then dicRooms(strRoomId) = True
            
            Dim strSceneId As String: strSceneId = modUtil.GetNamedOrHeaderValue(wks, NAME_SCENE_ID, Array("Scene ID", NAME_SCENE_ID, "Szene ID"))
            If Len(strSceneId) > 0 Then dicScenes(strSceneId) = True
            
            Call modUtil.CollectColumnBlockGroupValues(wks, ROOM_OBJ_GROUP_HEADER_ROW, _
                ROOM_OBJ_GROUP_END_ROW, _
                Array(ROOM_HDR_NM_PICKUPABLE_OBJ, ROOM_HDR_NM_MULTISTATE_OBJ, ROOM_HDR_NM_TOUCHABLE_OBJ), _
                ROOM_OBJ_GROUP_CATEGORY_COLUMN_WIDTH, dicObjects)
        End If
    Next wks
    
    Dim wksLists As Worksheet: Set wksLists = modUtil.GetSheetByCodeName(SHEET_LISTS)
    
    If Not wksLists Is Nothing Then

        wksLists.Columns(LISTS_COL_ROOM_ID).Clear
        
        wksLists.Cells(1, LISTS_COL_ROOM_ID).Value = LISTS_HDR_NM_ROOM_ID
        wksLists.Cells(1, LISTS_COL_OBJECTS).Value = LISTS_HDR_NM_OBJECTS
        wksLists.Cells(1, LISTS_COL_SCENE_ID).Value = LISTS_HDR_NM_SCENE_ID
        wksLists.Range("A1:ZZ1").Font.Bold = True
    
        ' --- Write Room IDs sorted, must always be rewritten, as it is related to the room pages
        WriteDictSetToColumn wksLists, dicRooms, 2, LISTS_COL_ROOM_ID
        
        ' --- Append only missing Object names
        modUtil.CollectColumnValues wksLists, Array(LISTS_HDR_NM_OBJECTS), dicExisting
        modUtil.AppendMissingDictKeysToColumn wksLists, LISTS_COL_OBJECTS, dicExisting, dicObjects
        
        ' --- Append only missing Scene IDs
        modUtil.CollectColumnValues wksLists, Array(LISTS_HDR_NM_SCENE_ID), dicExisting
        modUtil.AppendMissingDictKeysToColumn wksLists, LISTS_COL_SCENE_ID, dicExisting, dicScenes
    
        UpdateNamedListRange NAME_LIST_ROOM_IDS, wksLists, LISTS_COL_ROOM_ID
        UpdateNamedListRange NAME_LIST_OBJECTS, wksLists, LISTS_COL_OBJECTS
        UpdateNamedListRange NAME_LIST_SCENE_IDS, wksLists, LISTS_COL_SCENE_ID
    End If
End Sub


