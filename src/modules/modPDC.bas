Attribute VB_Name = "modPDC"
' -----------------------------------------------------------------------------------
' Module    : modPDC
' Purpose   : Build, update, and synchronize a simple "Puzzle Dependency Chart"
'             from worksheet data; create nodes, draw connectors, validate model.
'
' Notes     :
'   - This module focuses on chart orchestration and data extraction.
'   - Data conventions:
'       * Sheet "PDCData": A:C = Node (ID, Name, Type), E:F = Edges (From, To)
'       * Sheet "Chart": Shapes named by Node ID; connectors represent edges
'   - Business rules (e.g., type-to-color) are kept minimal here.
' -----------------------------------------------------------------------------------
Option Explicit
Option Private Module

' ===== Public API ====================================================================


' -----------------------------------------------------------------------------------
' Function  : GeneratePuzzleChart
' Purpose   : Builds a fresh "Chart" sheet, creates node shapes from table "PDCData"
'             (columns A:C), and draws connectors from columns E:F.
'
' Parameters:
'   (none)
'
' Returns   :
'
' Notes     :
'   - Deletes any existing sheet named "Chart" without prompt.
'   - Node shape name = Node ID, fill color via GetColorByType.
'   - Simple grid placement, starts at (100,100) with fixed spacing.
'   - Uses elbow connectors with triangle arrowheads.
' -----------------------------------------------------------------------------------
Public Sub GeneratePuzzleChart()
    On Error GoTo ErrHandler

    Dim dataSheet As Worksheet, chartSheet As Worksheet
    Dim nodesDict As Scripting.Dictionary
    Dim rowIndex As Long, lastRow As Long
    Dim nodeShape As Shape
    Dim posX As Double, posY As Double
    Dim nodeId As String, nodeName As String, nodeType As String
    Dim fromId As String, toId As String
    Dim fromShape As Shape, toShape As Shape

    Set dataSheet = ActiveWorkbook.Sheets("PDCData")

    ' Recreate "Chart" sheet (ignore if not present)
    On Error Resume Next
    Application.DisplayAlerts = False
    ActiveWorkbook.Sheets("Chart").Delete
    Application.DisplayAlerts = True
    On Error GoTo ErrHandler

    Set chartSheet = ActiveWorkbook.Sheets.Add
    chartSheet.Name = "Chart"

    Set nodesDict = New Scripting.Dictionary

    ' Create node shapes
    lastRow = dataSheet.Cells(dataSheet.Rows.Count, 1).End(xlUp).Row
    posX = 100: posY = 100

    For rowIndex = 2 To lastRow ' skip header row
        nodeId = dataSheet.Cells(rowIndex, 1).Value
        nodeName = dataSheet.Cells(rowIndex, 2).Value
        nodeType = dataSheet.Cells(rowIndex, 3).Value

        Set nodeShape = chartSheet.Shapes.AddShape(msoShapeRoundedRectangle, posX, posY, 120, 40)
        nodeShape.TextFrame.Characters.text = nodeName
        nodeShape.Name = nodeId
        nodeShape.Fill.ForeColor.RGB = GetColorByType(nodeType)

        nodesDict.Add nodeId, nodeShape

        posX = posX + 180
        If posX > 1000 Then
            posX = 100
            posY = posY + 100
        End If
    Next rowIndex

    ' Draw connectors
    lastRow = dataSheet.Cells(dataSheet.Rows.Count, 5).End(xlUp).Row
    For rowIndex = 2 To lastRow
        fromId = dataSheet.Cells(rowIndex, 5).Value
        toId = dataSheet.Cells(rowIndex, 6).Value

        Set fromShape = nodesDict(fromId)
        Set toShape = nodesDict(toId)

        chartSheet.Shapes.AddConnector(msoConnectorElbow, 0, 0, 100, 100).Select
        With Selection.ShapeRange
            .ConnectorFormat.BeginConnect fromShape, 3
            .ConnectorFormat.EndConnect toShape, 1
            .Line.ForeColor.RGB = RGB(50, 50, 50)
            .Line.EndArrowheadStyle = msoArrowheadTriangle
        End With
    Next rowIndex

    Exit Sub
ErrHandler:
    modErr.ReportError "GeneratePuzzleChart", Err.Number, Erl, caption:=modMain.AppProjectName
End Sub

' -----------------------------------------------------------------------------------
' Function  : UpdatePuzzleChart
' Purpose   : Updates existing node shapes' text and color on "Chart", removes all
'             old connectors, and redraws connectors based on table "PDCData".
'
' Parameters:
'   (none)
'
' Returns   :
'
' Notes     :
'   - Skips nodes not found by name, colors via GetColorByType.
'   - Removes every connector on the chart before redrawing edges.
' -----------------------------------------------------------------------------------
Public Sub UpdatePuzzleChart()
    On Error GoTo ErrHandler

    Dim dataSheet As Worksheet, chartSheet As Worksheet
    Dim lastRow As Long, rowIndex As Long
    Dim nodeId As String, nodeName As String, nodeType As String
    Dim targetShape As Shape
    Dim fromId As String, toId As String
    Dim fromShape As Shape, toShape As Shape

    Set dataSheet = ActiveWorkbook.Sheets("PDCData")
    Set chartSheet = ActiveWorkbook.Sheets("Chart")

    ' 1) Update node texts & colors
    lastRow = dataSheet.Cells(dataSheet.Rows.Count, 1).End(xlUp).Row
    For rowIndex = 2 To lastRow
        nodeId = dataSheet.Cells(rowIndex, 1).Value
        nodeName = dataSheet.Cells(rowIndex, 2).Value
        nodeType = dataSheet.Cells(rowIndex, 3).Value

        On Error Resume Next
        Set targetShape = chartSheet.Shapes(nodeId)
        On Error GoTo ErrHandler

        If Not targetShape Is Nothing Then
            targetShape.TextFrame.Characters.text = nodeName
            targetShape.Fill.ForeColor.RGB = GetColorByType(nodeType)
        End If
    Next rowIndex

    ' 2) Delete old connections (remove all connectors)
    Dim shpIndex As Long
    For shpIndex = chartSheet.Shapes.Count To 1 Step -1
        If chartSheet.Shapes(shpIndex).Connector Then
            chartSheet.Shapes(shpIndex).Delete
        End If
    Next shpIndex

    ' 3) Draw new connections (From/To)
    lastRow = dataSheet.Cells(dataSheet.Rows.Count, 5).End(xlUp).Row
    For rowIndex = 2 To lastRow
        fromId = dataSheet.Cells(rowIndex, 5).Value
        toId = dataSheet.Cells(rowIndex, 6).Value

        On Error Resume Next
        Set fromShape = chartSheet.Shapes(fromId)
        Set toShape = chartSheet.Shapes(toId)
        On Error GoTo ErrHandler

        If Not fromShape Is Nothing And Not toShape Is Nothing Then
            Dim connectorShape As Shape
            Set connectorShape = chartSheet.Shapes.AddConnector(msoConnectorElbow, 0, 0, 100, 100)
            With connectorShape
                .ConnectorFormat.BeginConnect fromShape, 3
                .ConnectorFormat.EndConnect toShape, 1
                .Line.ForeColor.RGB = RGB(50, 50, 50)
                .Line.EndArrowheadStyle = msoArrowheadTriangle
            End With
        End If
    Next rowIndex

    MsgBox "Diagram updated!", vbInformation, AppProjectName
    Exit Sub
ErrHandler:
    modErr.ReportError "UpdatePuzzleChart", Err.Number, Erl, caption:=modMain.AppProjectName
End Sub

' -----------------------------------------------------------------------------------
' Function  : SyncPuzzleChart
' Purpose   : Synchronizes the chart with the data: updates existing nodes, creates
'             missing nodes with grid placement, removes old connectors, redraws edges.
'
' Parameters:
'   (none)
'
' Returns   :
'
' Notes     :
'   - Builds a dictionary of existing non-connector shapes by Name.
'   - New nodes are placed in rows of five, starting at (100,100).
'   - Connectors are fully rebuilt from columns E (From) and F (To).
' -----------------------------------------------------------------------------------
Public Sub SyncPuzzleChart()
    On Error GoTo ErrHandler

    Dim dataSheet As Worksheet, chartSheet As Worksheet
    Dim lastRow As Long, rowIndex As Long
    Dim nodeId As String, nodeName As String, nodeType As String
    Dim targetShape As Shape
    Dim fromId As String, toId As String
    Dim fromShape As Shape, toShape As Shape
    Dim nodesDict As Scripting.Dictionary
    Dim posX As Double, posY As Double
    Dim shapeCount As Long

    Set dataSheet = ActiveWorkbook.Sheets("PDCData")
    Set chartSheet = ActiveWorkbook.Sheets("Chart")
    Set nodesDict = New Scripting.Dictionary

    ' 1) Initial placement (for potentially new nodes)
    posX = 100: posY = 100
    shapeCount = 0

    ' 2) Capture existing non-connector shapes
    For Each targetShape In chartSheet.Shapes
        If Not targetShape.Connector Then
            On Error Resume Next
            If Not nodesDict.Exists(targetShape.Name) Then
                nodesDict.Add targetShape.Name, targetShape
            End If
            On Error GoTo ErrHandler
        End If
    Next targetShape

    shapeCount = nodesDict.Count
    
    ' 3) Process all nodes from data table
    lastRow = dataSheet.Cells(dataSheet.Rows.Count, 1).End(xlUp).Row
    For rowIndex = 2 To lastRow
        nodeId = Trim$(dataSheet.Cells(rowIndex, 1).Value)
        nodeName = dataSheet.Cells(rowIndex, 2).Value
        nodeType = dataSheet.Cells(rowIndex, 3).Value

        If nodesDict.Exists(nodeId) Then
            ' Update existing shape
            Set targetShape = nodesDict(nodeId)
            targetShape.TextFrame.Characters.text = nodeName
            targetShape.Fill.ForeColor.RGB = GetColorByType(nodeType)
        Else
            ' Create new shape
            Set targetShape = chartSheet.Shapes.AddShape(msoShapeRoundedRectangle, posX, posY, 120, 40)
            targetShape.Name = nodeId
            targetShape.TextFrame.Characters.text = nodeName
            targetShape.Fill.ForeColor.RGB = GetColorByType(nodeType)
            targetShape.TextFrame.HorizontalAlignment = xlHAlignCenter
            
            ' Add to dictionary
            If nodesDict.Exists(nodeId) Then
                nodesDict.Remove nodeId
            End If
            nodesDict.Add nodeId, targetShape

            ' Advance simple grid position (5 shapes per row)
            shapeCount = shapeCount + 1
            posX = 100 + ((shapeCount - 1) Mod 5) * 180
            posY = 100 + Int((shapeCount - 1) / 5) * 100
        End If
    Next rowIndex

    ' 4) Remove old connectors
    Dim shpIndex As Long
    For shpIndex = chartSheet.Shapes.Count To 1 Step -1
        If chartSheet.Shapes(shpIndex).Connector Then
            chartSheet.Shapes(shpIndex).Delete
        End If
    Next shpIndex

    ' 5) Redraw new connectors based on E:F
    lastRow = dataSheet.Cells(dataSheet.Rows.Count, 5).End(xlUp).Row
    For rowIndex = 2 To lastRow
        fromId = Trim$(dataSheet.Cells(rowIndex, 5).Value)
        toId = Trim$(dataSheet.Cells(rowIndex, 6).Value)

        If nodesDict.Exists(fromId) And nodesDict.Exists(toId) Then
            Set fromShape = nodesDict(fromId)
            Set toShape = nodesDict(toId)

            Dim connectorShape As Shape
            Set connectorShape = chartSheet.Shapes.AddConnector(msoConnectorElbow, 0, 0, 100, 100)
            With connectorShape
                .ConnectorFormat.BeginConnect fromShape, 3
                .ConnectorFormat.EndConnect toShape, 1
                .Line.ForeColor.RGB = RGB(50, 50, 50)
                .Line.EndArrowheadStyle = msoArrowheadTriangle
            End With
        End If
    Next rowIndex

    MsgBox "Puzzle chart synchronized!", vbInformation, AppProjectName
    Exit Sub
ErrHandler:
    modErr.ReportError "SyncPuzzleChart", Err.Number, Erl, caption:=modMain.AppProjectName
End Sub

' -----------------------------------------------------------------------------------
' Function  : BuildPdcData
' Purpose   : Scans all "Room*" sheets, reads puzzle rows, and expands the DependsOn
'             list into edge rows on sheet "PDCData".
'
' Parameters:
'   (none)
'
' Returns   :
'
' Notes     :
'   - Ensures/clears sheet "PDCData", writes headers: ID, From, To, Type, Condition, Notes.
'   - Detects puzzle table via LocatePuzzleTable, maps columns via MapColumns.
'   - Emits one row per dependency entry; ID is sequential.
' -----------------------------------------------------------------------------------
Public Sub BuildPdcData()
    On Error GoTo ErrHandler

    Dim ws As Worksheet, targetSheet As Worksheet, rowOut As Long
    Dim headerDict As Scripting.Dictionary: Set headerDict = New Scripting.Dictionary
    Set targetSheet = EnsureSheet("PDCData")
    targetSheet.Cells.Clear
    WriteHeaders targetSheet, Array("ID", "From", "To", "Type", "Condition", "Notes")
    rowOut = 2

    For Each ws In ThisWorkbook.Worksheets
        If Left$(ws.Name, 4) = "Room" Then
            Dim headerCell As Range: Set headerCell = LocatePuzzleTable(ws) ' finds header row by signature
            If Not headerCell Is Nothing Then
                Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, headerCell.column).End(xlUp).Row
                Dim colMap As Object: Set colMap = MapColumns(ws.Rows(headerCell.Row))
                Dim rowIndex As Long
                For rowIndex = headerCell.Row + 1 To lastRow
                    Dim toId As String: toId = Trim$(CStr(ws.Cells(rowIndex, colMap("PuzzleID")).Value))
                    If Len(toId) = 0 Then GoTo NextRR
                    Dim dependsOn As String: dependsOn = CStr(ws.Cells(rowIndex, colMap("DependsOn")).Value)
                    Dim edgeType As String: edgeType = CStr(ws.Cells(rowIndex, colMap("Typ")).Value)
                    Dim conditionText As String: conditionText = CStr(ws.Cells(rowIndex, colMap("ErfordertItem")).Value)
                    Dim noteText As String: noteText = CStr(ws.Cells(rowIndex, colMap("Notes")).Value)
                    Dim depParts() As String, idx As Long
                    If Len(Trim$(dependsOn)) = 0 Then
                        ' orphan; still emit node info via empty From if desired (not implemented)
                    Else
                        depParts = Split(dependsOn, ",")
                        For idx = LBound(depParts) To UBound(depParts)
                            Dim fromId As String: fromId = Trim$(depParts(idx))
                            If Len(fromId) > 0 Then
                                targetSheet.Cells(rowOut, 1).Value = rowOut - 1
                                targetSheet.Cells(rowOut, 2).Value = fromId
                                targetSheet.Cells(rowOut, 3).Value = toId
                                targetSheet.Cells(rowOut, 4).Value = IIf(Len(edgeType) > 0, edgeType, "requires")
                                targetSheet.Cells(rowOut, 5).Value = conditionText
                                targetSheet.Cells(rowOut, 6).Value = noteText
                                rowOut = rowOut + 1
                            End If
                        Next idx
                    End If
NextRR:
                Next rowIndex
            End If
        End If
    Next ws

    Exit Sub
ErrHandler:
    modErr.ReportError "BuildPdcData", Err.Number, Erl, caption:=modMain.AppProjectName
End Sub

' -----------------------------------------------------------------------------------
' Function  : ValidateModel
' Purpose   : Validates the model: ensures unique PuzzleIDs across Room sheets and
'             checks that all edges in "PDCData" reference existing IDs.
'
' Parameters:
'   (none)
'
' Returns   :
'
' Notes     :
'   - Writes issues to sheet "Validation" with headers: Type, Message.
'   - Reports duplicates and missing references (From/To not found).
'   - Auto-fits columns at the end.
' -----------------------------------------------------------------------------------
Public Sub ValidateModel()
    On Error GoTo ErrHandler

    ' Collect IDs across Room sheets
    Dim idsDict As Scripting.Dictionary: Set idsDict = New Scripting.Dictionary
    Dim ws As Worksheet, issuesSheet As Worksheet, rowOut As Long
    Set issuesSheet = EnsureSheet("Validation")
    issuesSheet.Cells.Clear
    WriteHeaders issuesSheet, Array("Type", "Message")
    rowOut = 2

    ' Gather IDs
    Dim wsSrc As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If Left$(ws.Name, 4) = "Room" Then
            Dim headerCell As Range: Set headerCell = LocatePuzzleTable(ws)
            If Not headerCell Is Nothing Then
                Dim rowIndex As Long, lastRow As Long
                lastRow = ws.Cells(ws.Rows.Count, headerCell.column).End(xlUp).Row
                Dim colMap As Object: Set colMap = MapColumns(ws.Rows(headerCell.Row))
                For rowIndex = headerCell.Row + 1 To lastRow
                    Dim curId As String: curId = Trim$(CStr(ws.Cells(rowIndex, colMap("PuzzleID")).Value))
                    If Len(curId) > 0 Then
                        If idsDict.Exists(curId) Then
                            issuesSheet.Cells(rowOut, 1).Value = "Duplicate"
                            issuesSheet.Cells(rowOut, 2).Value = "PuzzleID appears multiple times: " & curId
                            rowOut = rowOut + 1
                        Else
                            idsDict(curId) = True
                        End If
                    End If
                Next rowIndex
            End If
        End If
    Next ws

    ' Check edges in "PDCData"
    Dim dataSheet As Worksheet: Set dataSheet = EnsureSheet("PDCData")
    Dim lastDataRow As Long: lastDataRow = dataSheet.Cells(dataSheet.Rows.Count, 1).End(xlUp).Row
    Dim edgeRow As Long
    For edgeRow = 2 To lastDataRow
        Dim fromId As String: fromId = Trim$(CStr(dataSheet.Cells(edgeRow, 2).Value))
        Dim toId As String: toId = Trim$(CStr(dataSheet.Cells(edgeRow, 3).Value))
        If Len(fromId) > 0 And Not idsDict.Exists(fromId) Then
            issuesSheet.Cells(rowOut, 1).Value = "MissingRef"
            issuesSheet.Cells(rowOut, 2).Value = "From not found: " & fromId
            rowOut = rowOut + 1
        End If
        If Len(toId) > 0 And Not idsDict.Exists(toId) Then
            issuesSheet.Cells(rowOut, 1).Value = "MissingRef"
            issuesSheet.Cells(rowOut, 2).Value = "To not found: " & toId
            rowOut = rowOut + 1
        End If
    Next edgeRow

    issuesSheet.Columns.AutoFit
    Exit Sub
ErrHandler:
    modErr.ReportError "ValidateModel", Err.Number, Erl, caption:=modMain.AppProjectName
End Sub

' ===== Private helpers ===============================================================

' -----------------------------------------------------------------------------------
' Function  : GetColorByType
' Purpose   : Returns an RGB color for a given node type string.
'
' Parameters:
'   nodeType [String]  - Node type, e.g. "story", "puzzle", "item".
'
' Returns   : Long (RGB color)
'
' Notes     :
'   - Case-insensitive match, defaults to gray (200,200,200).
' -----------------------------------------------------------------------------------
Private Function GetColorByType(ByVal nodeType As String) As Long
    Select Case LCase$(nodeType)
        Case "story":  GetColorByType = RGB(180, 167, 214)
        Case "puzzle": GetColorByType = RGB(255, 255, 153)
        Case "item":   GetColorByType = RGB(153, 255, 153)
        Case "main":   GetColorByType = RGB(255, 204, 153)
        Case "sub":    GetColorByType = RGB(153, 204, 255)
        Case Else:     GetColorByType = RGB(200, 200, 200)
    End Select
End Function

' -----------------------------------------------------------------------------------
' Function  : LocatePuzzleTable
' Purpose   : Finds the header row of the puzzle table by detecting required column
'             names within the first 50 rows and 50 columns.
'
' Parameters:
'   ws [Worksheet] - Worksheet to search.
'
' Returns   : Range (cell at detected header row, column 1), or Nothing.
'
' Notes     :
'   - Requires at least two of: "PuzzleID", "DependsOn", "Typ".
' -----------------------------------------------------------------------------------
Private Function LocatePuzzleTable(ByVal ws As Worksheet) As Range
    ' Finds the header row by required column names
    Dim requiredHeaders As Variant: requiredHeaders = Array("PuzzleID", "DependsOn", "Typ")
    Dim rowIndex As Long, colIndex As Long, hits As Long

    For rowIndex = 1 To 50
        hits = 0
        For colIndex = 1 To 50
            Dim cellText As String: cellText = CStr(ws.Cells(rowIndex, colIndex).Value)
            If UBound(Filter(requiredHeaders, cellText, True, vbTextCompare)) >= 0 Then hits = hits + 1
        Next colIndex
        If hits >= 2 Then
            Set LocatePuzzleTable = ws.Cells(rowIndex, 1)
            Exit Function
        End If
    Next rowIndex
End Function

' -----------------------------------------------------------------------------------
' Function  : MapColumns
' Purpose   : Builds a dictionary mapping header text to column index, starting at
'             the provided header row and scanning up to 50 columns to the right.
'
' Parameters:
'   headerRow [Range] - A single-row range containing headers.
'
' Returns   : Object (Scripting.Dictionary: key = trimmed header, value = column index)
'
' Notes     :
'   - Ignores empty header cells.
' -----------------------------------------------------------------------------------
Private Function MapColumns(ByVal headerRow As Range) As Scripting.Dictionary
    Dim map As Scripting.Dictionary: Set map = New Scripting.Dictionary
    Dim colIndex As Long
    For colIndex = headerRow.column To headerRow.column + 50
        Dim nameText As String: nameText = CStr(headerRow.Parent.Cells(headerRow.Row, colIndex).Value)
        If Len(nameText) > 0 Then
            map(Trim$(nameText)) = colIndex
        End If
    Next colIndex
    Set MapColumns = map
End Function

' -----------------------------------------------------------------------------------
' Procedure : WriteHeaders
' Purpose   : Writes a header array into row 1 of the target worksheet and bolds it.
'
' Parameters:
'   ws       [Worksheet] - Target worksheet.
'   headerArray  [Variant]   - 1D array of header captions.
'
' Returns   : (none)
'
' Notes     :
'   - Writes starting at cell (1,1), sequentially across columns.
' -----------------------------------------------------------------------------------
Private Sub WriteHeaders(ByVal ws As Worksheet, ByVal headerArray As Variant)
    Dim idx As Long
    For idx = LBound(headerArray) To UBound(headerArray)
        ws.Cells(1, 1 + idx).Value = headerArray(idx)
        ws.Cells(1, 1 + idx).Font.Bold = True
    Next idx
End Sub
