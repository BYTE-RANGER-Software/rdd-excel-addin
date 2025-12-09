Attribute VB_Name = "modPDC"
' -----------------------------------------------------------------------------------
' Module    : modPDC
' Purpose   : Build, update, and synchronize a simple "Puzzle Dependency Chart (PDC)"
'             based on Ron Gilbert's methodology. Creates nodes (puzzles, items,
'             flags) and edges (dependencies, requirements) from Room sheet data.
'
' Public API:
'   - BuildPdcData        : Scans all Room sheets, extracts puzzle data, builds
'                           nodes and edges for the PDC
'   - GeneratePuzzleChart : Creates visual chart from PDCData
'   - UpdatePuzzleChart   : Updates existing chart
'   - SyncPuzzleChart     : Synchronizes chart with data
'   - ValidateModel       : Validates puzzle IDs and edge references
'   - NavigateToPuzzle    : Navigates to puzzle source (called by shape click)
'   - FindPuzzleLocation  : Finds a puzzle in Room sheets by ID
'
' Data Structure (Sheet "PDCData"):
'   Nodes (A:F): NodeID, NodeName, NodeType, Room, Difficulty, Status
'   Edges (H:L): EdgeID, FromID, ToID, EdgeType, Notes
'
' Node Types:
'   - puzzle : A puzzle step (from Puzzle ID)
'   - item   : An inventory item (prefix "i_" or "i")
'   - flag   : A knowledge flag (prefix "g_" for global, "r_" for room)
'   - state  : A state or condition (other tokens from Requires)
'
' Edge Types:
'   - depends  : Puzzle depends on another puzzle (from DependsOn column)
'   - requires : Puzzle requires item/flag/state (from Requires column)
'
' Dependencies:
'   - modRooms   : IsRoomSheet
'   - modSheets  :
'   - modRanges  :
'   - modConst   : Named range constants
'   - modErr     : Error reporting
'   - modMain    : AppProjectName
'
' Notes     :
'   - Uses Named Ranges for robust column detection
'   - Automatically detects node types from ID prefixes
'   - Supports comma-separated lists in DependsOn and Requires
'   - Puzzle nodes are clickable and navigate to source row
' -----------------------------------------------------------------------------------
Option Explicit
Option Private Module

' ===== Private Constants ============================================================

' Node type identifiers
Private Const NODE_TYPE_PUZZLE  As String = "puzzle"
Private Const NODE_TYPE_ITEM    As String = "item"
Private Const NODE_TYPE_FLAG    As String = "flag"
Private Const NODE_TYPE_STATE   As String = "state"

' Edge type identifiers
Private Const EDGE_TYPE_DEPENDS  As String = "depends"
Private Const EDGE_TYPE_REQUIRES As String = "requires"

' ID prefixes for type detection
Private Const PREFIX_ITEM_LOWER  As String = "i_"
Private Const PREFIX_ITEM_SHORT  As String = "i"
Private Const PREFIX_FLAG_GLOBAL As String = "g_"
Private Const PREFIX_FLAG_ROOM   As String = "r_"

' PDCData column offsets (1-based)
Private Const COL_NODE_ID       As Long = 1   ' A
Private Const COL_NODE_NAME     As Long = 2   ' B
Private Const COL_NODE_TYPE     As Long = 3   ' C
Private Const COL_NODE_ROOM     As Long = 4   ' D
Private Const COL_NODE_DIFF     As Long = 5   ' E
Private Const COL_NODE_STATUS   As Long = 6   ' F

Private Const COL_EDGE_ID       As Long = 8   ' H
Private Const COL_EDGE_FROM     As Long = 9   ' I
Private Const COL_EDGE_TO       As Long = 10  ' J
Private Const COL_EDGE_TYPE     As Long = 11  ' K
Private Const COL_EDGE_NOTES    As Long = 12  ' L

' ===== Public API ====================================================================

' -----------------------------------------------------------------------------------
' Procedure : BuildPdcData
' Purpose   : Scans all Room sheets, extracts puzzle data from the PUZZLES section,
'             and builds a comprehensive nodes and edges dataset for the PDC.
'
' Parameters:
'   srcBook       [Workbook] - The workbook to scan for room sheets
'   outNodesCount [Long]     - (ByRef) Returns the number of nodes created
'   outEdgesCount [Long]     - (ByRef) Returns the number of edges created
'
' Returns   : (none)
'
' Output Structure (Sheet "PDCData"):
'   Nodes (A:F): NodeID | NodeName | NodeType | Room | Difficulty | Status
'   Edges (H:L): EdgeID | FromID | ToID | EdgeType | Notes
'
' Notes:
'   - Uses modRooms.IsRoomSheet() for proper Room sheet detection
'   - Reads data via Named Ranges (NAME_RANGE_PUZZLES_*)
'   - Automatically creates implicit nodes for items/flags found in Requires
'   - Deduplicates nodes across all rooms
' -----------------------------------------------------------------------------------
Public Sub BuildPdcData( _
    ByVal srcBook As Workbook, _
    ByRef outNodesCount As Long, _
    ByRef outEdgesCount As Long)
    
    On Error GoTo ErrHandler
    
    Dim roomSheet As Worksheet
    Dim targetSheet As Worksheet
    Dim nodesDict As Scripting.Dictionary    ' Key = NodeID, Value = node data array
    Dim edgesCol As Collection               ' Collection of edge data arrays
    Dim roomID As String
    
    ' Initialize outputs
    outNodesCount = 0
    outEdgesCount = 0
    
    ' Initialize collections
    Set nodesDict = New Scripting.Dictionary
    Set edgesCol = New Collection
    
    ' Scan all Room sheets
    For Each roomSheet In srcBook.Worksheets
        If modRooms.IsRoomSheet(roomSheet, roomID) Then
            CollectPuzzleData roomSheet, roomID, nodesDict, edgesCol
        End If
    Next roomSheet
    
    ' Create/clear PDCData sheet and write results
    Set targetSheet = modSheets.EnsureSheet("PDCData")
    targetSheet.Cells.Clear
    
    ' Write headers
    WriteNodeHeaders targetSheet
    WriteEdgeHeaders targetSheet
    
    ' Write data
    WriteNodes targetSheet, nodesDict
    WriteEdges targetSheet, edgesCol
    
    ' Format
    targetSheet.Columns.AutoFit
    
    ' Return counts
    outNodesCount = nodesDict.count
    outEdgesCount = edgesCol.count
    
    Exit Sub
    
ErrHandler:
    modErr.ReportError "BuildPdcData", Err.Number, Erl, caption:=modMain.AppProjectName
End Sub

' -----------------------------------------------------------------------------------
' Procedure : GeneratePuzzleChart
' Purpose   : Builds a fresh "Chart" sheet, creates node shapes from PDCData
'             and draws connectors for edges.
'
' Parameters:
'   srcBook [Workbook] - The workbook containing PDCData sheet
'
' Notes:
'   - Deletes any existing sheet named "Chart" without prompt.
'   - Node shape name = Node ID, fill color based on NodeType.
'   - Uses grid placement starting at (100,100).
' -----------------------------------------------------------------------------------
Public Sub GeneratePuzzleChart(ByVal srcBook As Workbook)
    On Error GoTo ErrHandler

    Dim dataSheet As Worksheet, chartSheet As Worksheet
    Dim nodesDict As Scripting.Dictionary
    Dim rowIndex As Long, lastRow As Long
    Dim nodeShape As Shape
    Dim posX As Double, posY As Double
    Dim nodeID As String, nodeName As String, nodeType As String
    Dim fromId As String, toId As String
    Dim fromShape As Shape, toShape As Shape
    Dim shapeCount As Long

    Set dataSheet = srcBook.Sheets("PDCData")

    ' Recreate "Chart" sheet
    On Error Resume Next
    Application.DisplayAlerts = False
    srcBook.Sheets("Chart").Delete
    Application.DisplayAlerts = True
    On Error GoTo ErrHandler

    Set chartSheet = srcBook.Sheets.Add
    chartSheet.name = "Chart"

    Set nodesDict = New Scripting.Dictionary

    ' Create node shapes from Nodes section (A:F)
    lastRow = dataSheet.Cells(dataSheet.Rows.count, COL_NODE_ID).End(xlUp).Row
    posX = 100: posY = 100
    shapeCount = 0

    For rowIndex = 2 To lastRow
        nodeID = Trim$(CStr(dataSheet.Cells(rowIndex, COL_NODE_ID).value))
        If LenB(nodeID) = 0 Then GoTo NextNode
        
        nodeName = dataSheet.Cells(rowIndex, COL_NODE_NAME).value
        nodeType = dataSheet.Cells(rowIndex, COL_NODE_TYPE).value

        Set nodeShape = chartSheet.Shapes.AddShape(msoShapeRoundedRectangle, posX, posY, 120, 40)
        With nodeShape
            .TextFrame.Characters.text = nodeName
            .name = nodeID
            .Fill.ForeColor.RGB = GetColorByType(nodeType)
            .TextFrame.HorizontalAlignment = xlHAlignCenter
            .TextFrame.VerticalAlignment = xlVAlignCenter
            .TextFrame2.TextRange.Font.Size = 9
        End With

        nodesDict.Add nodeID, nodeShape
        
        ' Grid placement (5 per row)
        shapeCount = shapeCount + 1
        posX = 100 + ((shapeCount) Mod 5) * 150
        If shapeCount Mod 5 = 0 Then posY = posY + 80
NextNode:
    Next rowIndex

    ' Draw connectors from Edges section (H:L)
    lastRow = dataSheet.Cells(dataSheet.Rows.count, COL_EDGE_ID).End(xlUp).Row
    For rowIndex = 2 To lastRow
        fromId = Trim$(CStr(dataSheet.Cells(rowIndex, COL_EDGE_FROM).value))
        toId = Trim$(CStr(dataSheet.Cells(rowIndex, COL_EDGE_TO).value))

        If nodesDict.Exists(fromId) And nodesDict.Exists(toId) Then
            Set fromShape = nodesDict(fromId)
            Set toShape = nodesDict(toId)

            Dim conn As Shape
            Set conn = chartSheet.Shapes.AddConnector(msoConnectorElbow, 0, 0, 100, 100)
            With conn
                .ConnectorFormat.BeginConnect fromShape, 3   ' Right side
                .ConnectorFormat.EndConnect toShape, 1       ' Left side
                .Line.ForeColor.RGB = GetEdgeColor(CStr(dataSheet.Cells(rowIndex, COL_EDGE_TYPE).value))
                .Line.EndArrowheadStyle = msoArrowheadTriangle
            End With
        End If
    Next rowIndex

    chartSheet.Activate
    Exit Sub
    
ErrHandler:
    modErr.ReportError "GeneratePuzzleChart", Err.Number, Erl, caption:=modMain.AppProjectName
End Sub

' -----------------------------------------------------------------------------------
' Procedure : UpdatePuzzleChart
' Purpose   : Updates existing node shapes' text and color on "Chart", removes all
'             old connectors, and redraws connectors based on PDCData.
'
' Parameters:
'   srcBook [Workbook] - The workbook containing PDCData and Chart sheets
'
' Notes:
'   - Called by modMain (user interaction handled there)
' -----------------------------------------------------------------------------------
Public Sub UpdatePuzzleChart(ByVal srcBook As Workbook)
    On Error GoTo ErrHandler

    Dim dataSheet As Worksheet, chartSheet As Worksheet
    Dim lastRow As Long, rowIndex As Long
    Dim nodeID As String, nodeName As String, nodeType As String
    Dim targetShape As Shape
    Dim fromId As String, toId As String
    Dim fromShape As Shape, toShape As Shape
    Dim nodesDict As Scripting.Dictionary

    Set dataSheet = srcBook.Sheets("PDCData")
    Set chartSheet = srcBook.Sheets("Chart")
    Set nodesDict = New Scripting.Dictionary
    
    ' Build dictionary of existing shapes
    For Each targetShape In chartSheet.Shapes
        If Not targetShape.Connector Then
            If Not nodesDict.Exists(targetShape.name) Then
                nodesDict.Add targetShape.name, targetShape
            End If
        End If
    Next targetShape

    ' 1) Update node texts & colors
    lastRow = dataSheet.Cells(dataSheet.Rows.count, COL_NODE_ID).End(xlUp).Row
    For rowIndex = 2 To lastRow
        nodeID = Trim$(CStr(dataSheet.Cells(rowIndex, COL_NODE_ID).value))
        nodeName = dataSheet.Cells(rowIndex, COL_NODE_NAME).value
        nodeType = dataSheet.Cells(rowIndex, COL_NODE_TYPE).value

        If nodesDict.Exists(nodeID) Then
            Set targetShape = nodesDict(nodeID)
            targetShape.TextFrame.Characters.text = nodeName
            targetShape.Fill.ForeColor.RGB = GetColorByType(nodeType)
        End If
    Next rowIndex

    ' 2) Delete old connectors
    Dim shpIndex As Long
    For shpIndex = chartSheet.Shapes.count To 1 Step -1
        If chartSheet.Shapes(shpIndex).Connector Then
            chartSheet.Shapes(shpIndex).Delete
        End If
    Next shpIndex

    ' 3) Draw new connectors
    lastRow = dataSheet.Cells(dataSheet.Rows.count, COL_EDGE_ID).End(xlUp).Row
    For rowIndex = 2 To lastRow
        fromId = Trim$(CStr(dataSheet.Cells(rowIndex, COL_EDGE_FROM).value))
        toId = Trim$(CStr(dataSheet.Cells(rowIndex, COL_EDGE_TO).value))

        If nodesDict.Exists(fromId) And nodesDict.Exists(toId) Then
            Set fromShape = nodesDict(fromId)
            Set toShape = nodesDict(toId)

            Dim connectorShape As Shape
            Set connectorShape = chartSheet.Shapes.AddConnector(msoConnectorElbow, 0, 0, 100, 100)
            With connectorShape
                .ConnectorFormat.BeginConnect fromShape, 3
                .ConnectorFormat.EndConnect toShape, 1
                .Line.ForeColor.RGB = GetEdgeColor(CStr(dataSheet.Cells(rowIndex, COL_EDGE_TYPE).value))
                .Line.EndArrowheadStyle = msoArrowheadTriangle
            End With
        End If
    Next rowIndex

    MsgBox "Chart aktualisiert!", vbInformation, modMain.AppProjectName
    Exit Sub
    
ErrHandler:
    modErr.ReportError "UpdatePuzzleChart", Err.Number, Erl, caption:=modMain.AppProjectName
End Sub

' -----------------------------------------------------------------------------------
' Procedure : SyncPuzzleChart
' Purpose   : Synchronizes the chart with PDCData: updates existing nodes, creates
'             missing nodes, removes old connectors, redraws edges.
'
' Parameters:
'   srcBook [Workbook] - The workbook containing PDCData and Chart sheets
'
' Notes:
'   - Called by modMain.SynchonizePuzzleDependencyChart (user interaction handled there)
' -----------------------------------------------------------------------------------
Public Sub SyncPuzzleChart(ByVal srcBook As Workbook)
    On Error GoTo ErrHandler

    Dim dataSheet As Worksheet, chartSheet As Worksheet
    Dim lastRow As Long, rowIndex As Long
    Dim nodeID As String, nodeName As String, nodeType As String
    Dim targetShape As Shape
    Dim fromId As String, toId As String
    Dim fromShape As Shape, toShape As Shape
    Dim nodesDict As Scripting.Dictionary
    Dim posX As Double, posY As Double
    Dim shapeCount As Long

    Set dataSheet = srcBook.Sheets("PDCData")
    Set chartSheet = srcBook.Sheets("Chart")
    Set nodesDict = New Scripting.Dictionary

    ' Initial placement for new nodes
    posX = 100: posY = 100
    shapeCount = 0

    ' Capture existing non-connector shapes
    For Each targetShape In chartSheet.Shapes
        If Not targetShape.Connector Then
            If Not nodesDict.Exists(targetShape.name) Then
                nodesDict.Add targetShape.name, targetShape
                shapeCount = shapeCount + 1
            End If
        End If
    Next targetShape

    ' Calculate next position
    If shapeCount > 0 Then
        posX = 100 + (shapeCount Mod 5) * 150
        posY = 100 + Int(shapeCount / 5) * 80
    End If

    ' Process all nodes from data
    lastRow = dataSheet.Cells(dataSheet.Rows.count, COL_NODE_ID).End(xlUp).Row
    For rowIndex = 2 To lastRow
        nodeID = Trim$(CStr(dataSheet.Cells(rowIndex, COL_NODE_ID).value))
        If LenB(nodeID) = 0 Then GoTo NextSyncNode
        
        nodeName = dataSheet.Cells(rowIndex, COL_NODE_NAME).value
        nodeType = dataSheet.Cells(rowIndex, COL_NODE_TYPE).value

        If nodesDict.Exists(nodeID) Then
            ' Update existing shape
            Set targetShape = nodesDict(nodeID)
            targetShape.TextFrame.Characters.text = nodeName
            targetShape.Fill.ForeColor.RGB = GetColorByType(nodeType)
        Else
            ' Create new shape
            Set targetShape = chartSheet.Shapes.AddShape(msoShapeRoundedRectangle, posX, posY, 120, 40)
            With targetShape
                .name = nodeID
                .TextFrame.Characters.text = nodeName
                .Fill.ForeColor.RGB = GetColorByType(nodeType)
                .TextFrame.HorizontalAlignment = xlHAlignCenter
            End With
            nodesDict.Add nodeID, targetShape

            ' Advance position
            shapeCount = shapeCount + 1
            posX = 100 + (shapeCount Mod 5) * 150
            If shapeCount Mod 5 = 0 Then posY = posY + 80
        End If
NextSyncNode:
    Next rowIndex

    ' Remove old connectors
    Dim shpIndex As Long
    For shpIndex = chartSheet.Shapes.count To 1 Step -1
        If chartSheet.Shapes(shpIndex).Connector Then
            chartSheet.Shapes(shpIndex).Delete
        End If
    Next shpIndex

    ' Redraw connectors
    lastRow = dataSheet.Cells(dataSheet.Rows.count, COL_EDGE_ID).End(xlUp).Row
    For rowIndex = 2 To lastRow
        fromId = Trim$(CStr(dataSheet.Cells(rowIndex, COL_EDGE_FROM).value))
        toId = Trim$(CStr(dataSheet.Cells(rowIndex, COL_EDGE_TO).value))

        If nodesDict.Exists(fromId) And nodesDict.Exists(toId) Then
            Set fromShape = nodesDict(fromId)
            Set toShape = nodesDict(toId)

            Dim conn As Shape
            Set conn = chartSheet.Shapes.AddConnector(msoConnectorElbow, 0, 0, 100, 100)
            With conn
                .ConnectorFormat.BeginConnect fromShape, 3
                .ConnectorFormat.EndConnect toShape, 1
                .Line.ForeColor.RGB = GetEdgeColor(CStr(dataSheet.Cells(rowIndex, COL_EDGE_TYPE).value))
                .Line.EndArrowheadStyle = msoArrowheadTriangle
            End With
        End If
    Next rowIndex

    MsgBox "Chart synchronisiert!", vbInformation, modMain.AppProjectName
    Exit Sub
    
ErrHandler:
    modErr.ReportError "SyncPuzzleChart", Err.Number, Erl, caption:=modMain.AppProjectName
End Sub

' -----------------------------------------------------------------------------------
' Procedure : ValidateModel
' Purpose   : Validates the PDC model: checks for duplicate IDs, missing references,
'             and orphan nodes.
' -----------------------------------------------------------------------------------
Public Sub ValidateModel()
    On Error GoTo ErrHandler

    Dim dataSheet As Worksheet, issuesSheet As Worksheet
    Dim nodesDict As Scripting.Dictionary
    Dim rowOut As Long, rowIndex As Long, lastRow As Long
    Dim nodeID As String, fromId As String, toId As String
    Dim issueCount As Long
    
    Set dataSheet = EnsureSheet("PDCData")
    Set issuesSheet = EnsureSheet("PDC_Validation")
    Set nodesDict = New Scripting.Dictionary
    
    issuesSheet.Cells.Clear
    WriteValidationHeaders issuesSheet
    rowOut = 2
    issueCount = 0

    ' Collect all node IDs
    lastRow = dataSheet.Cells(dataSheet.Rows.count, COL_NODE_ID).End(xlUp).Row
    For rowIndex = 2 To lastRow
        nodeID = Trim$(CStr(dataSheet.Cells(rowIndex, COL_NODE_ID).value))
        If LenB(nodeID) > 0 Then
            If nodesDict.Exists(nodeID) Then
                issuesSheet.Cells(rowOut, 1).value = "Duplicate"
                issuesSheet.Cells(rowOut, 2).value = "Node ID existiert mehrfach: " & nodeID
                issuesSheet.Cells(rowOut, 3).value = nodeID
                rowOut = rowOut + 1
                issueCount = issueCount + 1
            Else
                nodesDict(nodeID) = True
            End If
        End If
    Next rowIndex

    ' Check edges for missing references
    lastRow = dataSheet.Cells(dataSheet.Rows.count, COL_EDGE_ID).End(xlUp).Row
    For rowIndex = 2 To lastRow
        fromId = Trim$(CStr(dataSheet.Cells(rowIndex, COL_EDGE_FROM).value))
        toId = Trim$(CStr(dataSheet.Cells(rowIndex, COL_EDGE_TO).value))
        
        If LenB(fromId) > 0 And Not nodesDict.Exists(fromId) Then
            issuesSheet.Cells(rowOut, 1).value = "Missing Reference"
            issuesSheet.Cells(rowOut, 2).value = "Edge From-ID nicht gefunden: " & fromId
            issuesSheet.Cells(rowOut, 3).value = fromId
            rowOut = rowOut + 1
            issueCount = issueCount + 1
        End If
        
        If LenB(toId) > 0 And Not nodesDict.Exists(toId) Then
            issuesSheet.Cells(rowOut, 1).value = "Missing Reference"
            issuesSheet.Cells(rowOut, 2).value = "Edge To-ID nicht gefunden: " & toId
            issuesSheet.Cells(rowOut, 3).value = toId
            rowOut = rowOut + 1
            issueCount = issueCount + 1
        End If
    Next rowIndex

    issuesSheet.Columns.AutoFit
    
    If issueCount = 0 Then
        MsgBox "Validierung erfolgreich - keine Probleme gefunden!", _
               vbInformation, modMain.AppProjectName
    Else
        MsgBox "Validierung abgeschlossen:" & vbCrLf & _
               issueCount & " Problem(e) gefunden." & vbCrLf & _
               "Details im Sheet 'PDC_Validation'.", _
               vbExclamation, modMain.AppProjectName
        issuesSheet.Activate
    End If
    
    Exit Sub
    
ErrHandler:
    modErr.ReportError "ValidateModel", Err.Number, Erl, caption:=modMain.AppProjectName
End Sub

' ===== Private Helpers: Data Collection =============================================

' -----------------------------------------------------------------------------------
' Procedure : CollectPuzzleData
' Purpose   : Extracts puzzle data from a single Room sheet and adds nodes/edges
'             to the provided collections.
'
' Parameters:
'   ws         [Worksheet]  - Room sheet to process
'   roomID     [String]     - Room ID (e.g., "R001")
'   nodesDict  [Dictionary] - Dictionary to add nodes (Key=NodeID)
'   edgesCol   [Collection] - Collection to add edges
' -----------------------------------------------------------------------------------
Private Sub CollectPuzzleData( _
    ByVal ws As Worksheet, _
    ByVal roomID As String, _
    ByRef nodesDict As Scripting.Dictionary, _
    ByRef edgesCol As Collection)
    
    On Error GoTo ErrHandler
    
    Dim puzzleIDRange As Range
    Dim titleRange As Range
    Dim dependsOnRange As Range
    Dim requiresRange As Range
    Dim difficultyRange As Range
    Dim statusRange As Range
    Dim notesRange As Range
    
    ' Try to get Named Ranges - exit if PUZZLES section doesn't exist
    On Error Resume Next
    Set puzzleIDRange = ws.Range(modConst.NAME_RANGE_PUZZLES_PUZZLE_ID)
    Set titleRange = ws.Range(modConst.NAME_RANGE_PUZZLES_TITLE)
    Set dependsOnRange = ws.Range(modConst.NAME_RANGE_PUZZLES_DEPENDS_ON)
    Set requiresRange = ws.Range(modConst.NAME_RANGE_PUZZLES_REQUIRES)
    Set difficultyRange = ws.Range(modConst.NAME_RANGE_PUZZLES_DIFFICULTY)
    Set statusRange = ws.Range(modConst.NAME_RANGE_PUZZLES_STATUS)
    Set notesRange = ws.Range(modConst.NAME_RANGE_PUZZLES_NOTES)
    On Error GoTo ErrHandler
    
    ' Exit if essential ranges don't exist
    If puzzleIDRange Is Nothing Then Exit Sub
    
    Dim rowIndex As Long
    Dim puzzleID As String, title As String, difficulty As String, status As String
    Dim dependsOn As String, requires As String, notes As String
    
    ' Process each row in the puzzle range
    For rowIndex = 1 To puzzleIDRange.Rows.count
        puzzleID = Trim$(CStr(puzzleIDRange.Cells(rowIndex, 1).value))
        
        ' Skip empty rows
        If LenB(puzzleID) = 0 Then GoTo NextPuzzleRow
        
        ' Get puzzle data
        title = modRanges.GetRangeValue(titleRange, rowIndex)
        difficulty = modRanges.GetRangeValue(difficultyRange, rowIndex)
        status = modRanges.GetRangeValue(statusRange, rowIndex)
        dependsOn = modRanges.GetRangeValue(dependsOnRange, rowIndex)
        requires = modRanges.GetRangeValue(requiresRange, rowIndex)
        notes = modRanges.GetRangeValue(notesRange, rowIndex)
        
        ' Add puzzle node (if not exists)
        If Not nodesDict.Exists(puzzleID) Then
            nodesDict.Add puzzleID, CreateNodeData(puzzleID, title, NODE_TYPE_PUZZLE, _
                ws.name, difficulty, status)
        End If
        
        ' Process DependsOn -> creates "depends" edges
        ProcessDependencies puzzleID, dependsOn, EDGE_TYPE_DEPENDS, notes, _
            nodesDict, edgesCol
        
        ' Process Requires -> creates "requires" edges and implicit nodes
        ProcessRequirements puzzleID, requires, notes, nodesDict, edgesCol
        
NextPuzzleRow:
    Next rowIndex
    
    Exit Sub
    
ErrHandler:
    modErr.ReportError "CollectPuzzleData [" & ws.name & "]", Err.Number, Erl, _
        caption:=modMain.AppProjectName
End Sub

' -----------------------------------------------------------------------------------
' Procedure : ProcessDependencies
' Purpose   : Parses the DependsOn string and creates edges for each dependency.
'
' Parameters:
'   puzzleID   [String]     - Current puzzle ID (edge target)
'   dependsOn  [String]     - Comma-separated list of dependency IDs
'   edgeType   [String]     - Type of edge (e.g., "depends")
'   notes      [String]     - Notes to attach to edges
'   nodesDict  [Dictionary] - Dictionary of nodes (for reference checking)
'   edgesCol   [Collection] - Collection to add edges
' -----------------------------------------------------------------------------------
Private Sub ProcessDependencies( _
    ByVal puzzleID As String, _
    ByVal dependsOn As String, _
    ByVal edgeType As String, _
    ByVal notes As String, _
    ByRef nodesDict As Scripting.Dictionary, _
    ByRef edgesCol As Collection)
    
    If LenB(Trim$(dependsOn)) = 0 Then Exit Sub
    
    Dim parts() As String
    Dim idx As Long
    Dim depID As String
    
    parts = Split(dependsOn, ",")
    
    For idx = LBound(parts) To UBound(parts)
        depID = Trim$(parts(idx))
        If LenB(depID) > 0 Then
            ' Edge: From dependency -> To current puzzle
            edgesCol.Add CreateEdgeData(depID, puzzleID, edgeType, notes)
        End If
    Next idx
End Sub

' -----------------------------------------------------------------------------------
' Procedure : ProcessRequirements
' Purpose   : Parses the Requires string, creates implicit nodes for items/flags,
'             and creates "requires" edges.
'
' Parameters:
'   puzzleID   [String]     - Current puzzle ID (edge target)
'   requires   [String]     - Comma-separated list of required items/flags/states
'   notes      [String]     - Notes to attach to edges
'   nodesDict  [Dictionary] - Dictionary of nodes
'   edgesCol   [Collection] - Collection to add edges
' -----------------------------------------------------------------------------------
Private Sub ProcessRequirements( _
    ByVal puzzleID As String, _
    ByVal requires As String, _
    ByVal notes As String, _
    ByRef nodesDict As Scripting.Dictionary, _
    ByRef edgesCol As Collection)
    
    If LenB(Trim$(requires)) = 0 Then Exit Sub
    
    Dim parts() As String
    Dim idx As Long
    Dim reqID As String
    Dim nodeType As String
    Dim nodeName As String
    
    parts = Split(requires, ",")
    
    For idx = LBound(parts) To UBound(parts)
        reqID = Trim$(parts(idx))
        If LenB(reqID) = 0 Then GoTo NextReq
        
        ' Determine node type from prefix
        nodeType = DetectNodeType(reqID)
        
        ' Create implicit node if it doesn't exist and it's not a puzzle reference
        If Not nodesDict.Exists(reqID) Then
            ' Generate readable name from ID
            nodeName = GenerateNodeName(reqID, nodeType)
            nodesDict.Add reqID, CreateNodeData(reqID, nodeName, nodeType, "", "", "")
        End If
        
        ' Create edge: From requirement -> To puzzle
        edgesCol.Add CreateEdgeData(reqID, puzzleID, EDGE_TYPE_REQUIRES, notes)
NextReq:
    Next idx
End Sub

' ===== Private Helpers: Data Structures =============================================

' -----------------------------------------------------------------------------------
' Function  : CreateNodeData
' Purpose   : Creates a node data array for storage in the nodes dictionary.
' -----------------------------------------------------------------------------------
Private Function CreateNodeData( _
    ByVal nodeID As String, _
    ByVal nodeName As String, _
    ByVal nodeType As String, _
    ByVal room As String, _
    ByVal difficulty As String, _
    ByVal status As String) As Variant
    
    CreateNodeData = Array(nodeID, nodeName, nodeType, room, difficulty, status)
End Function

' -----------------------------------------------------------------------------------
' Function  : CreateEdgeData
' Purpose   : Creates an edge data array for storage in the edges collection.
' -----------------------------------------------------------------------------------
Private Function CreateEdgeData( _
    ByVal fromId As String, _
    ByVal toId As String, _
    ByVal edgeType As String, _
    ByVal notes As String) As Variant
    
    CreateEdgeData = Array(fromId, toId, edgeType, notes)
End Function

' ===== Private Helpers: Type Detection ==============================================

' -----------------------------------------------------------------------------------
' Function  : DetectNodeType
' Purpose   : Determines the node type based on the ID prefix.
'
' Parameters:
'   nodeID [String] - Node identifier
'
' Returns:
'   String - One of: "puzzle", "item", "flag", "state"
' -----------------------------------------------------------------------------------
Private Function DetectNodeType(ByVal nodeID As String) As String
    Dim lowerID As String
    lowerID = LCase$(nodeID)
    
    ' Check for item prefix
    If Left$(lowerID, 2) = PREFIX_ITEM_LOWER Then
        DetectNodeType = NODE_TYPE_ITEM
        Exit Function
    End If
    
    ' Check for flag prefixes (global or room)
    If Left$(lowerID, 2) = PREFIX_FLAG_GLOBAL Or _
       Left$(lowerID, 2) = PREFIX_FLAG_ROOM Then
        DetectNodeType = NODE_TYPE_FLAG
        Exit Function
    End If
    
    ' Check if it looks like a puzzle ID (contains underscore with letter prefix)
    If InStr(nodeID, "_P") > 0 Or InStr(nodeID, "_p") > 0 Then
        DetectNodeType = NODE_TYPE_PUZZLE
        Exit Function
    End If
    
    ' Default to state for anything else
    DetectNodeType = NODE_TYPE_STATE
End Function

' -----------------------------------------------------------------------------------
' Function  : GenerateNodeName
' Purpose   : Generates a human-readable name from a node ID.
'
' Parameters:
'   nodeID   [String] - Node identifier (e.g., "i_Key", "g_KnowsCode")
'   nodeType [String] - Type of node
'
' Returns:
'   String - Readable name (e.g., "Key", "Knows Code")
' -----------------------------------------------------------------------------------
Private Function GenerateNodeName( _
    ByVal nodeID As String, _
    ByVal nodeType As String) As String
    
    Dim result As String
    result = nodeID
    
    ' Remove common prefixes
    If Left$(LCase$(result), 2) = PREFIX_ITEM_LOWER Then
        result = Mid$(result, 3)
    ElseIf Left$(LCase$(result), 2) = PREFIX_FLAG_GLOBAL Or _
           Left$(LCase$(result), 2) = PREFIX_FLAG_ROOM Then
        result = Mid$(result, 3)
    End If
    
    ' Replace underscores with spaces
    result = Replace(result, "_", " ")
    
    ' Add type prefix for clarity
    Select Case nodeType
        Case NODE_TYPE_ITEM
            result = "[Item] " & result
        Case NODE_TYPE_FLAG
            result = "[Flag] " & result
        Case NODE_TYPE_STATE
            result = "[State] " & result
    End Select
    
    GenerateNodeName = result
End Function

' ===== Private Helpers: Writing Output ==============================================

' -----------------------------------------------------------------------------------
' Procedure : WriteNodeHeaders
' Purpose   : Writes the node section headers to the target sheet.
' -----------------------------------------------------------------------------------
Private Sub WriteNodeHeaders(ByVal ws As Worksheet)
    With ws
        .Cells(1, COL_NODE_ID).value = "NodeID"
        .Cells(1, COL_NODE_NAME).value = "NodeName"
        .Cells(1, COL_NODE_TYPE).value = "NodeType"
        .Cells(1, COL_NODE_ROOM).value = "Room"
        .Cells(1, COL_NODE_DIFF).value = "Difficulty"
        .Cells(1, COL_NODE_STATUS).value = "Status"
        
        ' Format headers
        .Range(.Cells(1, COL_NODE_ID), .Cells(1, COL_NODE_STATUS)).Font.Bold = True
        .Range(.Cells(1, COL_NODE_ID), .Cells(1, COL_NODE_STATUS)).Interior.Color = RGB(200, 200, 200)
    End With
End Sub

' -----------------------------------------------------------------------------------
' Procedure : WriteEdgeHeaders
' Purpose   : Writes the edge section headers to the target sheet.
' -----------------------------------------------------------------------------------
Private Sub WriteEdgeHeaders(ByVal ws As Worksheet)
    With ws
        .Cells(1, COL_EDGE_ID).value = "EdgeID"
        .Cells(1, COL_EDGE_FROM).value = "FromID"
        .Cells(1, COL_EDGE_TO).value = "ToID"
        .Cells(1, COL_EDGE_TYPE).value = "EdgeType"
        .Cells(1, COL_EDGE_NOTES).value = "Notes"
        
        ' Format headers
        .Range(.Cells(1, COL_EDGE_ID), .Cells(1, COL_EDGE_NOTES)).Font.Bold = True
        .Range(.Cells(1, COL_EDGE_ID), .Cells(1, COL_EDGE_NOTES)).Interior.Color = RGB(200, 200, 200)
    End With
End Sub

' -----------------------------------------------------------------------------------
' Procedure : WriteValidationHeaders
' Purpose   : Writes validation result headers.
' -----------------------------------------------------------------------------------
Private Sub WriteValidationHeaders(ByVal ws As Worksheet)
    With ws
        .Cells(1, 1).value = "IssueType"
        .Cells(1, 2).value = "Message"
        .Cells(1, 3).value = "Reference"
        .Range("A1:C1").Font.Bold = True
        .Range("A1:C1").Interior.Color = RGB(255, 200, 200)
    End With
End Sub

' -----------------------------------------------------------------------------------
' Procedure : WriteNodes
' Purpose   : Writes all nodes from the dictionary to the sheet.
' -----------------------------------------------------------------------------------
Private Sub WriteNodes( _
    ByVal ws As Worksheet, _
    ByRef nodesDict As Scripting.Dictionary)
    
    Dim rowOut As Long
    Dim key As Variant
    Dim nodeData As Variant
    
    rowOut = 2
    
    For Each key In nodesDict.Keys
        nodeData = nodesDict(key)
        
        ws.Cells(rowOut, COL_NODE_ID).value = nodeData(0)     ' NodeID
        ws.Cells(rowOut, COL_NODE_NAME).value = nodeData(1)   ' NodeName
        ws.Cells(rowOut, COL_NODE_TYPE).value = nodeData(2)   ' NodeType
        ws.Cells(rowOut, COL_NODE_ROOM).value = nodeData(3)   ' Room
        ws.Cells(rowOut, COL_NODE_DIFF).value = nodeData(4)   ' Difficulty
        ws.Cells(rowOut, COL_NODE_STATUS).value = nodeData(5) ' Status
        
        rowOut = rowOut + 1
    Next key
End Sub

' -----------------------------------------------------------------------------------
' Procedure : WriteEdges
' Purpose   : Writes all edges from the collection to the sheet.
' -----------------------------------------------------------------------------------
Private Sub WriteEdges( _
    ByVal ws As Worksheet, _
    ByRef edgesCol As Collection)
    
    Dim rowOut As Long
    Dim idx As Long
    Dim edgeData As Variant
    
    rowOut = 2
    
    For idx = 1 To edgesCol.count
        edgeData = edgesCol(idx)
        
        ws.Cells(rowOut, COL_EDGE_ID).value = idx              ' EdgeID (sequential)
        ws.Cells(rowOut, COL_EDGE_FROM).value = edgeData(0)    ' FromID
        ws.Cells(rowOut, COL_EDGE_TO).value = edgeData(1)      ' ToID
        ws.Cells(rowOut, COL_EDGE_TYPE).value = edgeData(2)    ' EdgeType
        ws.Cells(rowOut, COL_EDGE_NOTES).value = edgeData(3)   ' Notes
        
        rowOut = rowOut + 1
    Next idx
End Sub

' -----------------------------------------------------------------------------------
' Procedure: NavigateToPuzzle
' Purpose: Navigates to the puzzle in the corresponding room sheet.
' Triggered by double-clicking on a node shape in the chart.
'
' Parameters:
' nodeId (string) - Shape name (determined via Application.Caller)
'
' Notes:
'   - Called by modCallbacks.OnPdcNodeClick (after Strg + click detection)
'   - Application.Caller contains the shape name (= NodeID)
'   - Searches for the puzzle in all room sheets via named range
'   - Activates the sheet and selects the puzzle row
' -----------------------------------------------------------------------------------
Public Sub NavigateToPuzzle(ByRef nodeID As String)
    On Error GoTo ErrHandler

    Dim targetSheet As Worksheet
    Dim targetRow As Long
        
    If LenB(nodeID) = 0 Then
        MsgBox "Konnte Node-ID nicht ermitteln.", vbExclamation, modMain.AppProjectName
        Exit Sub
    End If
    
    ' Find the puzzle
    If FindPuzzleLocation(nodeID, targetSheet, targetRow) Then
        ' Navigate to the puzzle
        targetSheet.Activate
        targetSheet.Cells(targetRow, 1).Select
        
        ' Optional: Highlight row
        Application.GoTo targetSheet.Cells(targetRow, 1), Scroll:=True
    Else
        ' No puzzle found - could be an item/flag
        MsgBox "'" & nodeID & "' is not a puzzle or was not found." & vbCrLf & _
               "(Items and flags do not have a source line)", _
               vbInformation, modMain.AppProjectName
    End If
    
    Exit Sub
    
ErrHandler:
    modErr.ReportError "NavigateToPuzzle", Err.Number, Erl, caption:=modMain.AppProjectName
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


' ===== Private Helpers: Utility =====================================================

' -----------------------------------------------------------------------------------
' Function  : GetColorByType
' Purpose   : Returns an RGB color for a given node type.
' -----------------------------------------------------------------------------------
Private Function GetColorByType(ByVal nodeType As String) As Long
    Select Case LCase$(nodeType)
        Case NODE_TYPE_PUZZLE: GetColorByType = RGB(255, 255, 153)  ' Yellow
        Case NODE_TYPE_ITEM:   GetColorByType = RGB(153, 255, 153)  ' Green
        Case NODE_TYPE_FLAG:   GetColorByType = RGB(153, 204, 255)  ' Blue
        Case NODE_TYPE_STATE:  GetColorByType = RGB(255, 204, 153)  ' Orange
        Case Else:             GetColorByType = RGB(200, 200, 200)  ' Gray
    End Select
End Function

' -----------------------------------------------------------------------------------
' Function  : GetEdgeColor
' Purpose   : Returns an RGB color for a given edge type.
' -----------------------------------------------------------------------------------
Private Function GetEdgeColor(ByVal edgeType As String) As Long
    Select Case LCase$(edgeType)
        Case EDGE_TYPE_DEPENDS:  GetEdgeColor = RGB(50, 50, 50)     ' Dark gray
        Case EDGE_TYPE_REQUIRES: GetEdgeColor = RGB(0, 100, 0)      ' Dark green
        Case Else:               GetEdgeColor = RGB(100, 100, 100)  ' Medium gray
    End Select
End Function
