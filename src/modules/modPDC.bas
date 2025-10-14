Attribute VB_Name = "modPDC"
Option Explicit
Option Private Module

Sub GeneratePuzzleChart()
    Dim wksData As Worksheet, wksChart As Worksheet
    Dim dicNodes As Object, lngRow As Long
    Dim lngLastRow As Long, shp As Shape
    Dim dblX As Double, dblY As Double
    Dim strNodeId As String, strNodeName As String, strNodeType As String
    Dim strFromId As String, strToId As String
    Dim shpFrom As Shape, shpTo As Shape

    Set wksData = ThisWorkbook.Sheets("Daten")
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("Chart").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set wksChart = ThisWorkbook.Sheets.Add
    wksChart.name = "Chart"

    Set dicNodes = CreateObject("Scripting.Dictionary")

    ' Nodes erstellen
    lngLastRow = wksData.Cells(wksData.Rows.Count, 1).End(xlUp).Row
    dblX = 100: dblY = 100

    For lngRow = 2 To lngLastRow ' ab Zeile 2 (Überschrift)
        strNodeId = wksData.Cells(lngRow, 1).Value
        strNodeName = wksData.Cells(lngRow, 2).Value
        strNodeType = wksData.Cells(lngRow, 3).Value

        Set shp = wksChart.Shapes.AddShape(msoShapeRoundedRectangle, dblX, dblY, 120, 40)
        shp.TextFrame.Characters.Text = strNodeName
        shp.name = strNodeId
        shp.Fill.ForeColor.RGB = GetColorByType(strNodeType)

        dicNodes.Add strNodeId, shp

        dblX = dblX + 180
        If dblX > 1000 Then
            dblX = 100
            dblY = dblY + 100
        End If
    Next lngRow

    ' Verbindungen zeichnen
    lngLastRow = wksData.Cells(wksData.Rows.Count, 5).End(xlUp).Row
    For lngRow = 2 To lngLastRow
        strFromId = wksData.Cells(lngRow, 5).Value
        strToId = wksData.Cells(lngRow, 6).Value

        Set shpFrom = dicNodes(strFromId)
        Set shpTo = dicNodes(strToId)

        wksChart.Shapes.AddConnector(msoConnectorElbow, 0, 0, 100, 100).Select
        With Selection.ShapeRange
            .ConnectorFormat.BeginConnect shpFrom, 3
            .ConnectorFormat.EndConnect shpTo, 1
            .Line.ForeColor.RGB = RGB(50, 50, 50)
            .Line.EndArrowheadStyle = msoArrowheadTriangle
        End With
    Next lngRow
End Sub

Sub UpdatePuzzleChart()
    Dim wksData As Worksheet, wksChart As Worksheet
    Dim lngLastRow As Long, lngRow As Long
    Dim strNodeId As String, strNodeName As String, strNodeType As String
    Dim shp As Shape
    Dim strFromId As String, strToId As String
    Dim shpFrom As Shape, shpTo As Shape

    Set wksData = ThisWorkbook.Sheets("Daten")
    Set wksChart = ThisWorkbook.Sheets("Chart")

    ' 1. Texte & Farben aktualisieren
    lngLastRow = wksData.Cells(wksData.Rows.Count, 1).End(xlUp).Row
    For lngRow = 2 To lngLastRow
        strNodeId = wksData.Cells(lngRow, 1).Value
        strNodeName = wksData.Cells(lngRow, 2).Value
        strNodeType = wksData.Cells(lngRow, 3).Value

        On Error Resume Next
        Set shp = wksChart.Shapes(strNodeId)
        On Error GoTo 0

        If Not shp Is Nothing Then
            shp.TextFrame.Characters.Text = strNodeName
            shp.Fill.ForeColor.RGB = GetColorByType(strNodeType)
        End If
    Next lngRow

    ' 2. Alte Verbindungen löschen (alle Connectors entfernen)
    Dim iShp As Long
    For iShp = wksChart.Shapes.Count To 1 Step -1
        If wksChart.Shapes(iShp).Connector Then
            wksChart.Shapes(iShp).Delete
        End If
    Next iShp

    ' 3. Neue Verbindungen zeichnen laut From/To
    lngLastRow = wksData.Cells(wksData.Rows.Count, 5).End(xlUp).Row
    For lngRow = 2 To lngLastRow
        strFromId = wksData.Cells(lngRow, 5).Value
        strToId = wksData.Cells(lngRow, 6).Value

        On Error Resume Next
        Set shpFrom = wksChart.Shapes(strFromId)
        Set shpTo = wksChart.Shapes(strToId)
        On Error GoTo 0

        If Not shpFrom Is Nothing And Not shpTo Is Nothing Then
            Dim shpConn As Shape
            Set shpConn = wksChart.Shapes.AddConnector(msoConnectorElbow, 0, 0, 100, 100)
            With shpConn
                .ConnectorFormat.BeginConnect shpFrom, 3
                .ConnectorFormat.EndConnect shpTo, 1
                .Line.ForeColor.RGB = RGB(50, 50, 50)
                .Line.EndArrowheadStyle = msoArrowheadTriangle
            End With
        End If
    Next lngRow

    MsgBox "Diagramm wurde aktualisiert!", vbInformation
End Sub

Sub SyncPuzzleChart()
    Dim wksData As Worksheet, wksChart As Worksheet
    Dim lngLastRow As Long, lngRow As Long
    Dim strNodeId As String, strNodeName As String, strNodeType As String
    Dim shp As Shape
    Dim strFromId As String, strToId As String
    Dim shpFrom As Shape, shpTo As Shape
    Dim dicNodes As Object
    Dim dblX As Double, dblY As Double
    Dim lngShapeCount As Long

    Set wksData = ThisWorkbook.Sheets("Daten")
    Set wksChart = ThisWorkbook.Sheets("Chart")
    Set dicNodes = CreateObject("Scripting.Dictionary")

    ' 1. Shapes-Position initialisieren (für neue)
    dblX = 100: dblY = 100
    lngShapeCount = 0

    ' 2. Vorhandene Shapes erfassen
    For Each shp In wksChart.Shapes
        If Not shp.Connector Then
            On Error Resume Next
            If Not dicNodes.Exists(shp.name) Then
                dicNodes.Add shp.name, shp
            End If
            On Error GoTo 0
        End If
    Next shp

    lngShapeCount = dicNodes.Count
    
    ' 3. Alle Knoten aus Tabelle verarbeiten
    lngLastRow = wksData.Cells(wksData.Rows.Count, 1).End(xlUp).Row
    For lngRow = 2 To lngLastRow
        strNodeId = Trim(wksData.Cells(lngRow, 1).Value)
        strNodeName = wksData.Cells(lngRow, 2).Value
        strNodeType = wksData.Cells(lngRow, 3).Value

        If dicNodes.Exists(strNodeId) Then
            ' Vorhandene Shape ? aktualisieren
            Set shp = dicNodes(strNodeId)
            shp.TextFrame.Characters.Text = strNodeName
            shp.Fill.ForeColor.RGB = GetColorByType(strNodeType)
        Else
            ' Neue Shape erstellen
            Set shp = wksChart.Shapes.AddShape(msoShapeRoundedRectangle, dblX, dblY, 120, 40)
            shp.name = strNodeId
            shp.TextFrame.Characters.Text = strNodeName
            shp.Fill.ForeColor.RGB = GetColorByType(strNodeType)
            shp.TextFrame.HorizontalAlignment = xlHAlignCenter
            
            ' Neues Shape ins Dictionary
            If dicNodes.Exists(strNodeId) Then
                dicNodes.Remove strNodeId
            End If
            dicNodes.Add strNodeId, shp

            ' Position anpassen (einfaches Grid mit lngShapeCount)
            lngShapeCount = lngShapeCount + 1
            dblX = 100 + ((lngShapeCount - 1) Mod 5) * 180 ' 5 Shapes pro Reihe
            dblY = 100 + Int((lngShapeCount - 1) / 5) * 100
        End If
    Next lngRow

    ' 4. Alte Verbindungen (Connectoren) löschen
    Dim iShp As Long
    For iShp = wksChart.Shapes.Count To 1 Step -1
        If wksChart.Shapes(iShp).Connector Then
            wksChart.Shapes(iShp).Delete
        End If
    Next iShp

    ' 5. Neue Verbindungen zeichnen
    lngLastRow = wksData.Cells(wksData.Rows.Count, 5).End(xlUp).Row
    For lngRow = 2 To lngLastRow
        strFromId = Trim(wksData.Cells(lngRow, 5).Value)
        strToId = Trim(wksData.Cells(lngRow, 6).Value)

        If dicNodes.Exists(strFromId) And dicNodes.Exists(strToId) Then
            Set shpFrom = dicNodes(strFromId)
            Set shpTo = dicNodes(strToId)

            Dim shpConn As Shape
            Set shpConn = wksChart.Shapes.AddConnector(msoConnectorElbow, 0, 0, 100, 100)
            With shpConn
                .ConnectorFormat.BeginConnect shpFrom, 3
                .ConnectorFormat.EndConnect shpTo, 1
                .Line.ForeColor.RGB = RGB(50, 50, 50)
                .Line.EndArrowheadStyle = msoArrowheadTriangle
            End With
        End If
    Next lngRow

    MsgBox "Puzzle-Chart synchronisiert!", vbInformation
End Sub


Function GetColorByType(strNodeType As String) As Long
    Select Case LCase(strNodeType)
        Case "story": GetColorByType = RGB(180, 167, 214)
        Case "puzzle": GetColorByType = RGB(255, 255, 153)
        Case "item": GetColorByType = RGB(153, 255, 153)
        Case "main": GetColorByType = RGB(255, 204, 153)
        Case "sub": GetColorByType = RGB(153, 204, 255)
        Case Else: GetColorByType = RGB(200, 200, 200)
    End Select
End Function


Public Sub BuildPdcData()
    ' Scans all Room sheets, reads puzzle rows, expands DependsOn into edges
    Dim wks As Worksheet, wksTarget As Worksheet, lngRowOut As Long
    Dim dicHdr As Object: Set dicHdr = CreateObject("Scripting.Dictionary")
    Set wksTarget = EnsureSheet("Daten")
    wksTarget.Cells.Clear
    WriteHeaders wksTarget, Array("ID", "From", "To", "Type", "Condition", "Notes")
    lngRowOut = 2

    For Each wks In ThisWorkbook.Worksheets
        If Left$(wks.name, 4) = "Room" Then
            Dim rngHeader As Range: Set rngHeader = LocatePuzzleTable(wks) ' finds header row by signature
            If Not rngHeader Is Nothing Then
                Dim lngLastRow As Long: lngLastRow = wks.Cells(wks.Rows.Count, rngHeader.Column).End(xlUp).Row
                Dim dicCols As Object: Set dicCols = MapColumns(wks.Rows(rngHeader.Row))
                Dim lngRow As Long
                For lngRow = rngHeader.Row + 1 To lngLastRow
                    Dim strToId As String: strToId = Trim$(CStr(wks.Cells(lngRow, dicCols("PuzzleID")).Value))
                    If Len(strToId) = 0 Then GoTo NextRR
                    Dim strDep As String: strDep = CStr(wks.Cells(lngRow, dicCols("DependsOn")).Value)
                    Dim strTyp As String: strTyp = CStr(wks.Cells(lngRow, dicCols("Typ")).Value)
                    Dim strCond As String: strCond = CStr(wks.Cells(lngRow, dicCols("ErfordertItem")).Value)
                    Dim strNote As String: strNote = CStr(wks.Cells(lngRow, dicCols("Notes")).Value)
                    Dim astrDepParts() As String, lngIdx As Long
                    If Len(Trim$(strDep)) = 0 Then
                        ' orphan, still emit node info via empty From if you like
                    Else
                        astrDepParts = Split(strDep, ",")
                        For lngIdx = LBound(astrDepParts) To UBound(astrDepParts)
                            Dim strFromId As String: strFromId = Trim$(astrDepParts(lngIdx))
                            If Len(strFromId) > 0 Then
                                wksTarget.Cells(lngRowOut, 1).Value = lngRowOut - 1
                                wksTarget.Cells(lngRowOut, 2).Value = strFromId
                                wksTarget.Cells(lngRowOut, 3).Value = strToId
                                wksTarget.Cells(lngRowOut, 4).Value = IIf(Len(strTyp) > 0, strTyp, "requires")
                                wksTarget.Cells(lngRowOut, 5).Value = strCond
                                wksTarget.Cells(lngRowOut, 6).Value = strNote
                                lngRowOut = lngRowOut + 1
                            End If
                        Next lngIdx
                    End If
NextRR:
                Next lngRow
            End If
        End If
    Next wks
End Sub

Private Function LocatePuzzleTable(wks As Worksheet) As Range
    ' Finds the header row by required column names
    Dim vntFindSet As Variant: vntFindSet = Array("PuzzleID", "DependsOn", "Typ")
    Dim lngRow As Long, lngCol As Long, lngHits As Long
    For lngRow = 1 To 50
        lngHits = 0
        For lngCol = 1 To 50
            Dim strVal As String: strVal = CStr(wks.Cells(lngRow, lngCol).Value)
            If UBound(Filter(vntFindSet, strVal, True, vbTextCompare)) >= 0 Then lngHits = lngHits + 1
        Next lngCol
        If lngHits >= 2 Then
            Set LocatePuzzleTable = wks.Cells(lngRow, 1)
            Exit Function
        End If
    Next lngRow
End Function

Private Function MapColumns(rngHdrRow As Range) As Object
    ' Returns dictionary of column name to column index
    Dim dicMap As Object: Set dicMap = CreateObject("Scripting.Dictionary")
    Dim lngCol As Long
    For lngCol = rngHdrRow.Column To rngHdrRow.Column + 50
        Dim strName As String: strName = CStr(rngHdrRow.Parent.Cells(rngHdrRow.Row, lngCol).Value)
        If Len(strName) > 0 Then
            dicMap(Trim$(strName)) = lngCol
        End If
    Next lngCol
    Set MapColumns = dicMap
End Function

Private Sub WriteHeaders(wks As Worksheet, vntHeaders As Variant)
    Dim lngIndex As Long
    For lngIndex = LBound(vntHeaders) To UBound(vntHeaders)
        wks.Cells(1, 1 + lngIndex).Value = vntHeaders(lngIndex)
        wks.Cells(1, 1 + lngIndex).Font.Bold = True
    Next lngIndex
End Sub

Public Sub ValidateModel()
    ' Checks unique PuzzleIDs, missing references, cycles hint
    Dim dicIds As Object: Set dicIds = CreateObject("Scripting.Dictionary")
    Dim wks As Worksheet, wksIssues As Worksheet, lngRowOut As Long
    Set wksIssues = EnsureSheet("Validation")
    wksIssues.Cells.Clear
    WriteHeaders wksIssues, Array("Type", "Message")
    lngRowOut = 2

    ' Collect IDs
    Dim vntP As Variant, wksSrc As Worksheet
    For Each wks In ThisWorkbook.Worksheets
        If Left$(wks.name, 4) = "Room" Then
            Dim rngHeader As Range: Set rngHeader = LocatePuzzleTable(wks)
            If Not rngHeader Is Nothing Then
                Dim lngRow As Long, lngLastRow As Long
                lngLastRow = wks.Cells(wks.Rows.Count, rngHeader.Column).End(xlUp).Row
                Dim dicCols As Object: Set dicCols = MapColumns(wks.Rows(rngHeader.Row))
                For lngRow = rngHeader.Row + 1 To lngLastRow
                    Dim strId As String: strId = Trim$(CStr(wks.Cells(lngRow, dicCols("PuzzleID")).Value))
                    If Len(strId) > 0 Then
                        If dicIds.Exists(strId) Then
                            wksIssues.Cells(lngRowOut, 1).Value = "Duplicate"
                            wksIssues.Cells(lngRowOut, 2).Value = "PuzzleID appears multiple times: " & strId
                            lngRowOut = lngRowOut + 1
                        Else
                            dicIds(strId) = True
                        End If
                    End If
                Next lngRow
            End If
        End If
    Next wks

    ' Check edges in Daten
    Dim wksData As Worksheet: Set wksData = EnsureSheet("Daten")
    Dim lngLast As Long: lngLast = wksData.Cells(wksData.Rows.Count, 1).End(xlUp).Row
    Dim lngIndex As Long
    For lngIndex = 2 To lngLast
        Dim strFrom As String: strFrom = Trim$(CStr(wksData.Cells(lngIndex, 2).Value))
        Dim strTo As String: strTo = Trim$(CStr(wksData.Cells(lngIndex, 3).Value))
        If Len(strFrom) > 0 And Not dicIds.Exists(strFrom) Then
            wksIssues.Cells(lngRowOut, 1).Value = "MissingRef"
            wksIssues.Cells(lngRowOut, 2).Value = "From not found: " & strFrom
            lngRowOut = lngRowOut + 1
        End If
        If Len(strTo) > 0 And Not dicIds.Exists(strTo) Then
            wksIssues.Cells(lngRowOut, 1).Value = "MissingRef"
            wksIssues.Cells(lngRowOut, 2).Value = "To not found: " & strTo
            lngRowOut = lngRowOut + 1
        End If
    Next lngIndex

    ' Cycle check can be added via DFS, optional
    wksIssues.Columns.AutoFit
End Sub

