Attribute VB_Name = "modSearch"
' ====================================================================================
' Module    : modSearch
' Purpose   : Search functionality for finding usages of Items, Actors, and Flags
'             across all Room sheets in an RDD workbook.
'
' Public API:
'   - FindItemUsages         : Searches for all usages of an Item ID or Name across Room sheets
'   - FindActorUsages        : Searches for all usages of an Actor ID or Name across Room sheets.
'   - FindHotspotUsages      : Searches for all usages of a Hotspot ID or Name across Room sheets.
'   - FindFlagUsages         : Searches for all usages of a Flag ID across Room sheets.
'   - NavigateToSearchResult : Navigates to the cell referenced in a search result.
'   - FindItemDefinition     : Find Item by ID or Name across rooms
'   - FindFlagDefinition     : Find Flag by ID across rooms
'   - FindHotspotDefinition  : Find Hotspot by ID or Name across rooms
'   - FindPuzzleLocation     : Finds a puzzle ID across Room sheets
'
' Dependencies:
'   - modRooms      : IsRoomSheet
'   - modConst      : Named range constants
'   - modErr        : Error reporting
'   - modMain       : AppProjectName
'
' Notes:
'   - All Find* functions return a Collection of clsSearchResult objects
'   - Search is case-insensitive
'   - Supports comma-separated values in cells (e.g., "i_key, i_map")
' ====================================================================================
Option Explicit
Option Private Module

' Search categories for context display
Public Enum SearchCategory
    SC_Item = 1
    SC_Actor = 2
    SC_Flag = 3
    SC_Hotspot = 4
End Enum

' ====================================================================================
' Public API
' ====================================================================================

' -----------------------------------------------------------------------------------
' Function  : FindItemUsages
' Purpose   : Searches for all usages of an Item ID or Name across Room sheets.
'             Searches in: Puzzles_Requires, Puzzles_Grants
'
' Parameters:
'   searchTerm    [String]   - Item ID or Name to search for (e.g., "i_key" or "Golden Key")
'   targetBook    [Workbook] - Workbook to search in
'
' Returns   : Collection of clsSearchResult objects
' -----------------------------------------------------------------------------------
Public Function FindItemUsages(ByVal searchTerm As String, _
    ByVal targetBook As Workbook) As Collection
    
    On Error GoTo ErrHandler
    
    Dim results As Collection
    Set results = New Collection
    
    Dim ws As Worksheet
    Dim roomID As String
    
    ' Trim and validate search term
    searchTerm = Trim$(searchTerm)
    If Len(searchTerm) = 0 Then
        Set FindItemUsages = results
        Exit Function
    End If
    
    ' Search all Room sheets
    For Each ws In targetBook.Worksheets
        If modRooms.IsRoomSheet(ws, roomID) Then
            ' Search in Puzzles - Requires column
            SearchInNamedRange ws, searchTerm, _
                modConst.NAME_RANGE_PUZZLES_REQUIRES, _
                "Puzzle Requires", results
            
            ' Search in Puzzles - Grants column
            SearchInNamedRange ws, searchTerm, _
                modConst.NAME_RANGE_PUZZLES_GRANTS, _
                "Puzzle Grants", results
            
            ' Search in Pickupable Objects - Item ID (other definitions)
            SearchInNamedRange ws, searchTerm, _
                modConst.NAME_RANGE_PICKUPABLE_OBJECTS_ITEM_ID, _
                "Item Definition", results
            
            ' Search in Pickupable Objects - Item Name (other definitions)
            SearchInNamedRange ws, searchTerm, _
                modConst.NAME_RANGE_PICKUPABLE_OBJECTS_NAME, _
                "Item Definition (Name)", results
        End If
    Next ws
    
    Set FindItemUsages = results
    Exit Function
    
ErrHandler:
    modErr.ReportError "modSearch.FindItemUsages", Err.Number, Erl, caption:=modMain.AppProjectName
    Set FindItemUsages = New Collection
End Function

' -----------------------------------------------------------------------------------
' Function  : FindActorUsages
' Purpose   : Searches for all usages of an Actor ID or Name across Room sheets.
'             Searches in: Actors_Condition, Puzzles_Owner, Puzzles_Target
'
' Parameters:
'   searchTerm    [String]   - Actor ID or Name to search for (e.g., "cEgo" or "Player")
'   targetBook    [Workbook] - Workbook to search in
'
' Returns   : Collection of clsSearchResult objects
' -----------------------------------------------------------------------------------
Public Function FindActorUsages(ByVal searchTerm As String, _
    ByVal targetBook As Workbook) As Collection
    
    On Error GoTo ErrHandler
    
    Dim results As Collection
    Set results = New Collection
    
    Dim ws As Worksheet
    Dim roomID As String
    
    ' Trim and validate search term
    searchTerm = Trim$(searchTerm)
    If Len(searchTerm) = 0 Then
        Set FindActorUsages = results
        Exit Function
    End If
    
    ' Search all Room sheets
    For Each ws In targetBook.Worksheets
        If modRooms.IsRoomSheet(ws, roomID) Then
            ' Search in Actors - Condition column
            SearchInNamedRange ws, searchTerm, _
                modConst.NAME_RANGE_ACTORS_CONDITION, _
                "Actor Condition", results
            
            ' Search in Puzzles - Owner column
            SearchInNamedRange ws, searchTerm, _
                modConst.NAME_RANGE_PUZZLES_OWNER, _
                "Puzzle Owner", results
            
            ' Search in Puzzles - Target column
            SearchInNamedRange ws, searchTerm, _
                modConst.NAME_RANGE_PUZZLES_TARGET, _
                "Puzzle Target", results
            
            ' Search in Actors - Actor ID (other definitions)
            SearchInNamedRange ws, searchTerm, _
                modConst.NAME_RANGE_ACTORS_ACTOR_ID, _
                "Actor Definition", results
            
            ' Search in Actors - Actor Name (other definitions)
            SearchInNamedRange ws, searchTerm, _
                modConst.NAME_RANGE_ACTORS_ACTOR_NAME, _
                "Actor Definition (Name)", results
        End If
    Next ws
    
    Set FindActorUsages = results
    Exit Function
    
ErrHandler:
    modErr.ReportError "modSearch.FindActorUsages", Err.Number, Erl, caption:=modMain.AppProjectName
    Set FindActorUsages = New Collection
End Function

' -----------------------------------------------------------------------------------
' Function  : FindFlagUsages
' Purpose   : Searches for all usages of a Flag ID across Room sheets.
'             Searches in: Puzzles_Requires, Puzzles_DependsOn, Actors_Condition
'
' Parameters:
'   searchTerm    [String]   - Flag ID to search for (e.g., "g_hasMap")
'   targetBook    [Workbook] - Workbook to search in
'
' Returns   : Collection of clsSearchResult objects
' -----------------------------------------------------------------------------------
Public Function FindFlagUsages(ByVal searchTerm As String, _
    ByVal targetBook As Workbook) As Collection
    
    On Error GoTo ErrHandler
    
    Dim results As Collection
    Set results = New Collection
    
    Dim ws As Worksheet
    Dim roomID As String
    
    ' Trim and validate search term
    searchTerm = Trim$(searchTerm)
    If Len(searchTerm) = 0 Then
        Set FindFlagUsages = results
        Exit Function
    End If
    
    ' Search all Room sheets
    For Each ws In targetBook.Worksheets
        If modRooms.IsRoomSheet(ws, roomID) Then
            ' Search in Puzzles - Requires column
            SearchInNamedRange ws, searchTerm, _
                modConst.NAME_RANGE_PUZZLES_REQUIRES, _
                "Puzzle Requires", results
            
            ' Search in Puzzles - DependsOn column
            SearchInNamedRange ws, searchTerm, _
                modConst.NAME_RANGE_PUZZLES_DEPENDS_ON, _
                "Puzzle DependsOn", results
            
            ' Search in Actors - Condition column
            SearchInNamedRange ws, searchTerm, _
                modConst.NAME_RANGE_ACTORS_CONDITION, _
                "Actor Condition", results
        End If
    Next ws
    
    Set FindFlagUsages = results
    Exit Function
    
ErrHandler:
    modErr.ReportError "modSearch.FindFlagUsages", Err.Number, Erl, caption:=modMain.AppProjectName
    Set FindFlagUsages = New Collection
End Function

' -----------------------------------------------------------------------------------
' Function  : FindHotspotUsages
' Purpose   : Searches for all usages of a Hotspot ID or Name across Room sheets.
'             Searches in: Puzzles_Target, Actors_Condition
'
' Parameters:
'   searchTerm    [String]   - Hotspot ID or Name to search for (e.g., "hDoor" or "Old Door")
'   targetBook    [Workbook] - Workbook to search in
'
' Returns   : Collection of clsSearchResult objects
' -----------------------------------------------------------------------------------
Public Function FindHotspotUsages(ByVal searchTerm As String, _
    ByVal targetBook As Workbook) As Collection
    
    On Error GoTo ErrHandler
    
    Dim results As Collection
    Set results = New Collection
    
    Dim ws As Worksheet
    Dim roomID As String
    
    ' Trim and validate search term
    searchTerm = Trim$(searchTerm)
    If Len(searchTerm) = 0 Then
        Set FindHotspotUsages = results
        Exit Function
    End If
    
    ' Search all Room sheets
    For Each ws In targetBook.Worksheets
        If modRooms.IsRoomSheet(ws, roomID) Then
            ' Search in Puzzles - Target column (hotspots can be puzzle targets)
            SearchInNamedRange ws, searchTerm, _
                modConst.NAME_RANGE_PUZZLES_TARGET, _
                "Puzzle Target", results
            
            ' Search in Actors - Condition column (hotspots in conditions)
            SearchInNamedRange ws, searchTerm, _
                modConst.NAME_RANGE_ACTORS_CONDITION, _
                "Actor Condition", results
            
            ' Search in Touchable Objects - Hotspot ID (other definitions)
            SearchInNamedRange ws, searchTerm, _
                modConst.NAME_RANGE_TOUCHABLE_OBJECTS_HOTSPOT_ID, _
                "Hotspot Definition", results
            
            ' Search in Touchable Objects - Hotspot Name (other definitions)
            SearchInNamedRange ws, searchTerm, _
                modConst.NAME_RANGE_TOUCHABLE_OBJECTS_HOTSPOT_NAME, _
                "Hotspot Definition (Name)", results
        End If
    Next ws
    
    Set FindHotspotUsages = results
    Exit Function
    
ErrHandler:
    modErr.ReportError "modSearch.FindHotspotUsages", Err.Number, Erl, caption:=modMain.AppProjectName
    Set FindHotspotUsages = New Collection
End Function

' -----------------------------------------------------------------------------------
' Procedure : NavigateToSearchResult
' Purpose   : Navigates to the cell referenced in a search result.
'
' Parameters:
'   sheetName  [String] - Name of the worksheet
'   cellAddr   [String] - Cell address (e.g., "D15")
'   targetBook [Workbook] - Workbook containing the sheet
' -----------------------------------------------------------------------------------
Public Sub NavigateToSearchResult(ByVal sheetName As String, _
    ByVal cellAddr As String, _
    ByVal targetBook As Workbook)
    
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    
    Set ws = targetBook.Worksheets(sheetName)
    ws.Activate
    ws.Range(cellAddr).Select
    Application.GoTo ws.Range(cellAddr), Scroll:=True
    
    Exit Sub
    
ErrHandler:
    modErr.ReportError "modSearch.NavigateToSearchResult", Err.Number, Erl, caption:=modMain.AppProjectName
End Sub

' -----------------------------------------------------------------------------------
' Function: FindPuzzleLocation
' Purpose: Searches for a puzzle by its ID in all room sheets.
'
' Parameters:
'   puzzleID    [String]    - The puzzle ID to search for (e.g., “R001_P01”)
'   outSheet    [Worksheet] - (ByRef) The worksheet found
'   outRow      [Long]      - (ByRef) The row of the puzzle
'
' Returns   : Boolean - True if found
'
' Notes:
'   - Searches all room sheets via modRooms.IsRoomSheet
'   - Uses named range NAME_RANGE_PUZZLES_PUZZLE_ID
' -----------------------------------------------------------------------------------
Public Function FindPuzzleLocation( _
    ByVal puzzleID As String, _
    ByRef outSheet As Worksheet, _
    ByRef outRow As Long) As Boolean
    
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim roomID As String
    Dim puzzleIDRange As Range
    Dim cell As Range
    Dim rowIdx As Long
    
    ' Durchsuche alle Room-Sheets
    For Each ws In ActiveWorkbook.Worksheets
        If modRooms.IsRoomSheet(ws, roomID) Then
            ' Versuche Named Range zu holen
            On Error Resume Next
            Set puzzleIDRange = ws.Range(modConst.NAME_RANGE_PUZZLES_PUZZLE_ID)
            On Error GoTo ErrHandler
            
            If Not puzzleIDRange Is Nothing Then
                ' Durchsuche die Puzzle-IDs
                For rowIdx = 1 To puzzleIDRange.Rows.count
                    If Trim$(CStr(puzzleIDRange.Cells(rowIdx, 1).value)) = puzzleID Then
                        ' Gefunden!
                        Set outSheet = ws
                        outRow = puzzleIDRange.Cells(rowIdx, 1).Row
                        FindPuzzleLocation = True
                        Exit Function
                    End If
                Next rowIdx
            End If
            
            Set puzzleIDRange = Nothing
        End If
    Next ws
    
    ' Nicht gefunden
    FindPuzzleLocation = False
    Exit Function
    
ErrHandler:
    FindPuzzleLocation = False
End Function

' -----------------------------------------------------------------------------------
' Function  : FindItemDefinition
' Purpose   : Searches for an Item definition by ID across all Room sheets.
'
' Parameters:
'   itemID       [String]    - Item ID or NAme to find (e.g., "i_key1", or "key1")
'   targetBook   [Workbook]  - Workbook to search
'   outSheet     [Worksheet] - (ByRef) Returns the sheet containing the item
'   outRow       [Long]      - (ByRef) Returns the row of the item
'
' Returns   : Boolean - True if found
' -----------------------------------------------------------------------------------
Public Function FindItemDefinition(ByVal itemID As String, _
    ByVal targetBook As Workbook, _
    ByRef outSheet As Worksheet, _
    ByRef outRow As Long) As Boolean
    
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim roomID As String
    Dim itemIDRange As Range
    Dim itemNameRange As Range
    Dim i As Long
    
    FindItemDefinition = False
    
    For Each ws In targetBook.Worksheets
        If modRooms.IsRoomSheet(ws, roomID) Then
            On Error Resume Next
            Set itemIDRange = ws.Range(modConst.NAME_RANGE_PICKUPABLE_OBJECTS_ITEM_ID)
            On Error GoTo ErrHandler
            
            If Not itemIDRange Is Nothing Then
                For i = 1 To itemIDRange.Rows.count
                    If StrComp(Trim$(CStr(itemIDRange.Cells(i, 1).value)), itemID, vbTextCompare) = 0 Then
                        Set outSheet = ws
                        outRow = itemIDRange.Cells(i, 1).Row
                        FindItemDefinition = True
                        Exit Function
                    End If
                Next i
            End If
            
            Set itemIDRange = Nothing
            
            ' Search by item Name if not found by ID
            On Error Resume Next
            Set itemNameRange = ws.Range(modConst.NAME_RANGE_PICKUPABLE_OBJECTS_NAME)
            On Error GoTo ErrHandler
            
            If Not itemNameRange Is Nothing Then
                For i = 1 To itemNameRange.Rows.count
                    If StrComp(Trim$(CStr(itemNameRange.Cells(i, 1).value)), itemID, vbTextCompare) = 0 Then
                        Set outSheet = ws
                        outRow = itemNameRange.Cells(i, 1).Row
                        FindItemDefinition = True
                        Exit Function
                    End If
                Next i
            End If
            
            Set itemNameRange = Nothing
        End If
    Next ws
    
    Exit Function
    
ErrHandler:
    FindItemDefinition = False
End Function

' -----------------------------------------------------------------------------------
' Function  : FindFlagDefinition
' Purpose   : Searches for a Flag definition by ID across all Room sheets.
'
' Parameters:
'   flagID       [String]    - Flag ID to find (e.g., "g_hasMap")
'   targetBook   [Workbook]  - Workbook to search
'   outSheet     [Worksheet] - (ByRef) Returns the sheet containing the flag
'   outRow       [Long]      - (ByRef) Returns the row of the flag
'
' Returns   : Boolean - True if found
' -----------------------------------------------------------------------------------
Public Function FindFlagDefinition(ByVal flagID As String, _
    ByVal targetBook As Workbook, _
    ByRef outSheet As Worksheet, _
    ByRef outRow As Long) As Boolean
    
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim roomID As String
    Dim flagIDRange As Range
    Dim i As Long
    
    FindFlagDefinition = False
    
    For Each ws In targetBook.Worksheets
        If modRooms.IsRoomSheet(ws, roomID) Then
            On Error Resume Next
            Set flagIDRange = ws.Range(modConst.NAME_RANGE_FLAGS_FLAG_ID)
            On Error GoTo ErrHandler
            
            If Not flagIDRange Is Nothing Then
                For i = 1 To flagIDRange.Rows.count
                    If StrComp(Trim$(CStr(flagIDRange.Cells(i, 1).value)), flagID, vbTextCompare) = 0 Then
                        Set outSheet = ws
                        outRow = flagIDRange.Cells(i, 1).Row
                        FindFlagDefinition = True
                        Exit Function
                    End If
                Next i
            End If
            
            Set flagIDRange = Nothing
        End If
    Next ws
    
    Exit Function
    
ErrHandler:
    FindFlagDefinition = False
End Function

' -----------------------------------------------------------------------------------
' Function  : FindHotspotDefinition
' Purpose   : Searches for a Hotspot definition by ID or Name across all Room sheets.
'
' Parameters:
'   hotspotID    [String]    - Hotspot ID or Name to find (e.g., "hOldDoor" or "Old Door")
'   targetBook   [Workbook]  - Workbook to search
'   outSheet     [Worksheet] - (ByRef) Returns the sheet containing the hotspot
'   outRow       [Long]      - (ByRef) Returns the row of the hotspot
'
' Returns   : Boolean - True if found
' -----------------------------------------------------------------------------------
Public Function FindHotspotDefinition(ByVal hotspotID As String, _
    ByVal targetBook As Workbook, _
    ByRef outSheet As Worksheet, _
    ByRef outRow As Long) As Boolean
    
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim roomID As String
    Dim hotspotIDRange As Range
    Dim hotspotNameRange As Range
    Dim i As Long
    
    FindHotspotDefinition = False
    
    For Each ws In targetBook.Worksheets
        If modRooms.IsRoomSheet(ws, roomID) Then
            ' Search by Hotspot ID first
            On Error Resume Next
            Set hotspotIDRange = ws.Range(modConst.NAME_RANGE_TOUCHABLE_OBJECTS_HOTSPOT_ID)
            On Error GoTo ErrHandler
            
            If Not hotspotIDRange Is Nothing Then
                For i = 1 To hotspotIDRange.Rows.count
                    If StrComp(Trim$(CStr(hotspotIDRange.Cells(i, 1).value)), hotspotID, vbTextCompare) = 0 Then
                        Set outSheet = ws
                        outRow = hotspotIDRange.Cells(i, 1).Row
                        FindHotspotDefinition = True
                        Exit Function
                    End If
                Next i
            End If
            
            Set hotspotIDRange = Nothing
            
            ' Search by Hotspot Name if not found by ID
            On Error Resume Next
            Set hotspotNameRange = ws.Range(modConst.NAME_RANGE_TOUCHABLE_OBJECTS_HOTSPOT_NAME)
            On Error GoTo ErrHandler
            
            If Not hotspotNameRange Is Nothing Then
                For i = 1 To hotspotNameRange.Rows.count
                    If StrComp(Trim$(CStr(hotspotNameRange.Cells(i, 1).value)), hotspotID, vbTextCompare) = 0 Then
                        Set outSheet = ws
                        outRow = hotspotNameRange.Cells(i, 1).Row
                        FindHotspotDefinition = True
                        Exit Function
                    End If
                Next i
            End If
            
            Set hotspotNameRange = Nothing
        End If
    Next ws
    
    Exit Function
    
ErrHandler:
    FindHotspotDefinition = False
End Function

' ====================================================================================
' Private Helpers
' ====================================================================================

' -----------------------------------------------------------------------------------
' Procedure : SearchInNamedRange
' Purpose   : Searches for a term within a named range and adds matches to results.
'             Supports comma-separated values in cells.
'
' Parameters:
'   ws           [Worksheet]  - Worksheet containing the named range
'   searchTerm   [String]     - Term to search for (case-insensitive)
'   rangeName    [String]     - Name of the range to search
'   contextDesc  [String]     - Description for result context (e.g., "Puzzle Requires")
'   results      [Collection] - Collection to add results to (ByRef)
' -----------------------------------------------------------------------------------
Private Sub SearchInNamedRange(ByVal ws As Worksheet, _
    ByVal searchTerm As String, _
    ByVal rangeName As String, _
    ByVal contextDesc As String, _
    ByRef results As Collection)
    
    On Error GoTo ErrHandler
    
    Dim targetRange As Range
    Dim cell As Range
    Dim cellValue As String
    Dim resultData(0 To 3) As String  ' SheetName, CellAddress, Context, FullValue
    
    ' Try to get the named range
    On Error Resume Next
    Set targetRange = ws.Range(rangeName)
    On Error GoTo ErrHandler
    
    If targetRange Is Nothing Then Exit Sub
    
    ' Search each cell
    For Each cell In targetRange.Cells
        cellValue = Trim$(CStr(cell.value))
        
        If Len(cellValue) > 0 Then
            ' Check if search term is contained (supports comma-separated lists)
            If ContainsToken(cellValue, searchTerm) Then
                ' Build result data array
                resultData(0) = ws.name                          ' Sheet name
                resultData(1) = cell.Address(False, False)       ' Cell address (e.g., "D15")
                resultData(2) = contextDesc                      ' Context description
                resultData(3) = TruncateText(cellValue, 50)      ' Truncated cell value
                
                ' Add to results
                results.Add resultData
            End If
        End If
    Next cell
    
    Exit Sub
    
ErrHandler:
    ' Silently continue on errors (e.g., named range doesn't exist)
    Resume Next
End Sub

' -----------------------------------------------------------------------------------
' Function  : ContainsToken
' Purpose   : Checks if a cell value contains the search term as a token.
'             Handles comma-separated values and performs case-insensitive matching.
'
' Parameters:
'   cellValue   [String] - The cell value to search in
'   searchTerm  [String] - The term to search for
'
' Returns   : Boolean - True if found
' -----------------------------------------------------------------------------------
Private Function ContainsToken(ByVal cellValue As String, ByVal searchTerm As String) As Boolean
    Dim tokens() As String
    Dim token As Variant
    Dim i As Long
    
    ContainsToken = False
    
    ' First check simple contains (faster for single values)
    If InStr(1, cellValue, searchTerm, vbTextCompare) = 0 Then
        Exit Function
    End If
    
    ' If cell contains comma, split and check each token
    If InStr(cellValue, ",") > 0 Then
        tokens = Split(cellValue, ",")
        For i = LBound(tokens) To UBound(tokens)
            If StrComp(Trim$(tokens(i)), searchTerm, vbTextCompare) = 0 Then
                ContainsToken = True
                Exit Function
            End If
        Next i
    Else
        ' Single value - exact match check
        If StrComp(Trim$(cellValue), searchTerm, vbTextCompare) = 0 Then
            ContainsToken = True
        ' Also allow partial match for cases like "i_key" in "uses i_key"
        ElseIf InStr(1, cellValue, searchTerm, vbTextCompare) > 0 Then
            ContainsToken = True
        End If
    End If
End Function

' -----------------------------------------------------------------------------------
' Function  : TruncateText
' Purpose   : Truncates text to a maximum length, adding "..." if truncated.
'
' Parameters:
'   text      [String] - Text to truncate
'   maxLength [Long]   - Maximum length
'
' Returns   : String - Truncated text
' -----------------------------------------------------------------------------------
Private Function TruncateText(ByVal text As String, ByVal maxLength As Long) As String
    If Len(text) <= maxLength Then
        TruncateText = text
    Else
        TruncateText = Left$(text, maxLength - 3) & "..."
    End If
End Function


