Attribute VB_Name = "modExport"
' -----------------------------------------------------------------------------------
' Module    : modExport
' Purpose   : Export functionality for RDD workbooks.
'             Provides PDF export (with cover page, TOC, room sheets, chart)
'             and CSV export (nodes + edges for graph software compatibility).
'
' Public API:
'   - ExportWorkbookToPdf   : Exports complete RDD to PDF with cover and TOC
'   - ExportPdcToCsv        : Exports PDC data as nodes.csv and edges.csv
'
' Dependencies:
'   - modRooms      : IsRoomSheet
'   - modMain       : AppProjectName
'   - modErr        : Error reporting
'
' Notes:
'   - PDF uses Excel's native ExportAsFixedFormat
'   - CSV format compatible with yEd, Graphviz, EdrawMax
'   - Called by modMain orchestration methods (user interaction handled there)
' -----------------------------------------------------------------------------------
Option Explicit
Option Private Module

' Print area for room sheets
Private Const ROOM_PRINT_AREA As String = "$A$1:$N$115"

' CSV column headers
Private Const CSV_NODES_HEADER As String = "ID,Label,Type,Room,Difficulty,Status"
Private Const CSV_EDGES_HEADER As String = "Source,Target,Type,Label"

' -----------------------------------------------------------------------------------
' Procedure : ExportWorkbookToPdf
' Purpose   : Exports complete RDD workbook to PDF including:
'             - Cover page with document properties
'             - Table of contents
'             - All room sheets (print area, fit to page, portrait)
'             - PDC Chart (if exists)
'
' Parameters:
'   exportBook  [Workbook] - The workbook to export
'   filePath    [String]   - Full path for the output PDF file
'
' Returns   : Boolean - True on success
'
' Notes:
'   - Creates temporary sheets for cover and TOC, deletes after export
'   - Room sheets are exported with defined print area
' -----------------------------------------------------------------------------------
Public Function ExportWorkbookToPdf( _
    ByVal exportBook As Workbook, _
    ByVal filePath As String) As Boolean
    
    On Error GoTo ErrHandler
    
    Dim sheetsToExport As Collection
    Dim ws As Worksheet
    Dim coverSheet As Worksheet
    Dim ToCSheet As Worksheet
    Dim chartSheet As Worksheet
    Dim roomSheets As Collection
    Dim sheetNames() As String
    Dim i As Long
    Dim originalActiveSheet As Worksheet
    
    ' Store original active sheet
    Set originalActiveSheet = exportBook.ActiveSheet
    
    ' Initialize collections
    Set sheetsToExport = New Collection
    Set roomSheets = New Collection
    
    ' Collect room sheets
    For Each ws In exportBook.Worksheets
        If modRooms.IsRoomSheet(ws) Then
            roomSheets.Add ws
        End If
    Next ws
    
    ' Check if Chart exists
    On Error Resume Next
    Set chartSheet = exportBook.Sheets("Chart")
    On Error GoTo ErrHandler
    
    ' Create cover page
    Set coverSheet = CreateCoverPage(exportBook)
    sheetsToExport.Add coverSheet
    
    ' Create TOC
    Set ToCSheet = CreateTableOfContents(exportBook, roomSheets, Not chartSheet Is Nothing)
    sheetsToExport.Add ToCSheet
    
    ' Add room sheets (with print setup)
    For Each ws In roomSheets
        SetupRoomSheetForPrint ws
        sheetsToExport.Add ws
    Next ws
    
    ' Add chart if exists
    If Not chartSheet Is Nothing Then
        SetupChartSheetForPrint chartSheet
        sheetsToExport.Add chartSheet
    End If
    
    ' Build array of sheet names for multi-sheet export
    ReDim sheetNames(1 To sheetsToExport.count)
    For i = 1 To sheetsToExport.count
        sheetNames(i) = sheetsToExport(i).name
    Next i
    
    ' Select sheets for export
    exportBook.Sheets(sheetNames).Select
    
    ' Export to PDF
    exportBook.Sheets(sheetNames(1)).ExportAsFixedFormat _
        Type:=xlTypePDF, _
        fileNAme:=filePath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False
    
    ' Cleanup: Delete temporary sheets
    Application.DisplayAlerts = False
    coverSheet.Delete
    ToCSheet.Delete
    Application.DisplayAlerts = True
    
    ' Restore original active sheet
    originalActiveSheet.Activate
    
    ExportWorkbookToPdf = True
    Exit Function
    
ErrHandler:
    ' Cleanup on error
    On Error Resume Next
    Application.DisplayAlerts = False
    If Not coverSheet Is Nothing Then coverSheet.Delete
    If Not ToCSheet Is Nothing Then ToCSheet.Delete
    Application.DisplayAlerts = True
    If Not originalActiveSheet Is Nothing Then originalActiveSheet.Activate
    On Error GoTo 0
    
    modErr.ReportError "ExportWorkbookToPdf", Err.Number, Erl, caption:=modMain.AppProjectName
    ExportWorkbookToPdf = False
End Function

' -----------------------------------------------------------------------------------
' Procedure : ExportPdcToCsv
' Purpose   : Exports PDC data as two CSV files for graph software compatibility.
'             Creates nodes.csv and edges.csv in the specified folder.
'
' Parameters:
'   exportBook         [Workbook] - The workbook containing PDCData sheet
'   folderPath [String]   - Folder path for output files
'
' Returns   : Boolean - True on success
'
' Notes:
'   - Format compatible with yEd, Graphviz, EdrawMax
'   - Uses UTF-8 encoding (via ADODB.Stream)
' -----------------------------------------------------------------------------------
Public Function ExportPdcToCsv( _
    ByVal exportBook As Workbook, _
    ByVal folderPath As String) As Boolean
    
    On Error GoTo ErrHandler
    
    Dim dataSheet As Worksheet
    Dim nodesPath As String
    Dim edgesPath As String
    
    ' Ensure trailing backslash
    If Right$(folderPath, 1) <> "\" Then
        folderPath = folderPath & "\"
    End If
    
    ' Get PDCData sheet
    Set dataSheet = exportBook.Sheets("PDCData")
    
    ' Build file paths
    nodesPath = folderPath & "nodes.csv"
    edgesPath = folderPath & "edges.csv"
    
    ' Export nodes
    ExportNodesToCsv dataSheet, nodesPath
    
    ' Export edges
    ExportEdgesToCsv dataSheet, edgesPath
    
    ExportPdcToCsv = True
    Exit Function
    
ErrHandler:
    modErr.ReportError "ExportPdcToCsv", Err.Number, Erl, caption:=modMain.AppProjectName
    ExportPdcToCsv = False
End Function

' ===== Private Helpers: PDF =======================================================

' -----------------------------------------------------------------------------------
' Function  : CreateCoverPage
' Purpose   : Creates a temporary cover page sheet with document properties.
' -----------------------------------------------------------------------------------
Private Function CreateCoverPage(ByVal exportBook As Workbook) As Worksheet
    Dim coverSheet As Worksheet
    Dim props As DocumentProperties
    Dim rowNum As Long
    
    ' Create new sheet at beginning
    Set coverSheet = exportBook.Sheets.Add(Before:=exportBook.Sheets(1))
    coverSheet.name = "_RDD_Cover_"
    
    Set props = exportBook.BuiltinDocumentProperties
    
    With coverSheet
        ' Title area
        .Range("B3:F3").Merge
        .Range("B3").value = "Room Design Document"
        .Range("B3").Font.Size = 28
        .Range("B3").Font.Bold = True
        .Range("B3").HorizontalAlignment = xlCenter
        
        ' Document title from properties
        .Range("B5:F5").Merge
        .Range("B5").value = GetBuiltinProperty(props, "Title")
        .Range("B5").Font.Size = 20
        .Range("B5").HorizontalAlignment = xlCenter
        
        ' Subject/Theme
        .Range("B7:F7").Merge
        .Range("B7").value = GetBuiltinProperty(props, "Subject")
        .Range("B7").Font.Size = 14
        .Range("B7").Font.Italic = True
        .Range("B7").HorizontalAlignment = xlCenter
        
        ' Properties table
        rowNum = 12
        
        ' Author
        .Cells(rowNum, 2).value = "Author:"
        .Cells(rowNum, 2).Font.Bold = True
        .Cells(rowNum, 3).value = GetBuiltinProperty(props, "Author")
        rowNum = rowNum + 1
        
        ' Manager
        .Cells(rowNum, 2).value = "Manager:"
        .Cells(rowNum, 2).Font.Bold = True
        .Cells(rowNum, 3).value = GetBuiltinProperty(props, "Manager")
        rowNum = rowNum + 1
        
        ' Company
        .Cells(rowNum, 2).value = "Company:"
        .Cells(rowNum, 2).Font.Bold = True
        .Cells(rowNum, 3).value = GetBuiltinProperty(props, "Company")
        rowNum = rowNum + 1
        
        ' Category
        .Cells(rowNum, 2).value = "Category:"
        .Cells(rowNum, 2).Font.Bold = True
        .Cells(rowNum, 3).value = GetBuiltinProperty(props, "Category")
        rowNum = rowNum + 1
        
        ' Keywords
        .Cells(rowNum, 2).value = "Keywords:"
        .Cells(rowNum, 2).Font.Bold = True
        .Cells(rowNum, 3).value = GetBuiltinProperty(props, "Keywords")
        rowNum = rowNum + 2
        
        ' Comments
        .Cells(rowNum, 2).value = "Comments:"
        .Cells(rowNum, 2).Font.Bold = True
        .Range(.Cells(rowNum, 3), .Cells(rowNum + 2, 5)).Merge
        .Cells(rowNum, 3).value = GetBuiltinProperty(props, "Comments")
        .Cells(rowNum, 3).WrapText = True
        .Cells(rowNum, 3).VerticalAlignment = xlTop
        rowNum = rowNum + 4
        
        ' Date
        .Cells(rowNum, 2).value = "Export Date:"
        .Cells(rowNum, 2).Font.Bold = True
        .Cells(rowNum, 3).value = Format$(Now, "yyyy-mm-dd hh:mm")
        
        ' Column widths
        .Columns("A").ColumnWidth = 5
        .Columns("B").ColumnWidth = 15
        .Columns("C:F").ColumnWidth = 20
        
        ' Print setup
        With .PageSetup
            .Orientation = xlPortrait
            .FitToPagesWide = 1
            .FitToPagesTall = 1
            .CenterHorizontally = True
            .CenterVertically = True
        End With
    End With
    
    Set CreateCoverPage = coverSheet
End Function

' -----------------------------------------------------------------------------------
' Function  : CreateTableOfContents
' Purpose   : Creates a temporary TOC sheet listing all room sheets and chart.
' -----------------------------------------------------------------------------------
Private Function CreateTableOfContents( _
    ByVal exportBook As Workbook, _
    ByVal roomSheets As Collection, _
    ByVal hasChart As Boolean) As Worksheet
    
    Dim ToCSheet As Worksheet
    Dim rowNum As Long
    Dim roomSheet As Worksheet
    Dim roomID As String
    Dim roomAlias As String
    Dim pageNum As Long
    
    ' Create new sheet after cover
    Set ToCSheet = exportBook.Sheets.Add(After:=exportBook.Sheets(1))
    ToCSheet.name = "_RDD_TOC_"
    
    With ToCSheet
        ' Header
        .Range("B2").value = "Table of Contents"
        .Range("B2").Font.Size = 18
        .Range("B2").Font.Bold = True
        
        ' Column headers
        rowNum = 5
        .Cells(rowNum, 2).value = "Section"
        .Cells(rowNum, 3).value = "Room ID"
        .Cells(rowNum, 4).value = "Room Alias"
        .Range(.Cells(rowNum, 2), .Cells(rowNum, 4)).Font.Bold = True
        .Range(.Cells(rowNum, 2), .Cells(rowNum, 4)).Borders(xlEdgeBottom).LineStyle = xlContinuous
        
        rowNum = rowNum + 2
        pageNum = 1
        
        ' Cover entry
        .Cells(rowNum, 2).value = "Cover"
        rowNum = rowNum + 1
        
        ' TOC entry
        .Cells(rowNum, 2).value = "Table of Contents"
        rowNum = rowNum + 1
        
        ' Room sheets
        .Cells(rowNum, 2).value = "--- Rooms ---"
        .Cells(rowNum, 2).Font.Italic = True
        rowNum = rowNum + 1
        
        For Each roomSheet In roomSheets
            ' Get room info from named ranges
            roomID = GetRoomId(roomSheet)
            roomAlias = GetRoomAlias(roomSheet)
            
            .Cells(rowNum, 2).value = roomSheet.name
            .Cells(rowNum, 3).value = roomID
            .Cells(rowNum, 4).value = roomAlias
            rowNum = rowNum + 1
        Next roomSheet
        
        ' Chart entry
        If hasChart Then
            rowNum = rowNum + 1
            .Cells(rowNum, 2).value = "--- Chart ---"
            .Cells(rowNum, 2).Font.Italic = True
            rowNum = rowNum + 1
            .Cells(rowNum, 2).value = "Puzzle Dependency Chart"
        End If
        
        ' Column widths
        .Columns("A").ColumnWidth = 5
        .Columns("B").ColumnWidth = 25
        .Columns("C").ColumnWidth = 15
        .Columns("D").ColumnWidth = 25
        
        ' Print setup
        With .PageSetup
            .Orientation = xlPortrait
            .FitToPagesWide = 1
            .FitToPagesTall = 1
        End With
    End With
    
    Set CreateTableOfContents = ToCSheet
End Function

' -----------------------------------------------------------------------------------
' Procedure : SetupRoomSheetForPrint
' Purpose   : Configures room sheet print settings (area, orientation, fit to page).
' -----------------------------------------------------------------------------------
Private Sub SetupRoomSheetForPrint(ByVal ws As Worksheet)
    With ws.PageSetup
        .PrintArea = ROOM_PRINT_AREA
        .Orientation = xlPortrait
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .CenterHorizontally = True
    End With
End Sub

' -----------------------------------------------------------------------------------
' Procedure : SetupChartSheetForPrint
' Purpose   : Configures chart sheet print settings.
' -----------------------------------------------------------------------------------
Private Sub SetupChartSheetForPrint(ByVal ws As Worksheet)
    With ws.PageSetup
        .PrintArea = ""  ' Clear to print all content
        .Orientation = xlLandscape  ' Chart typically wider
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .CenterHorizontally = True
        .CenterVertically = True
    End With
End Sub

' -----------------------------------------------------------------------------------
' Function  : GetBuiltinProperty
' Purpose   : Safely retrieves a built-in document property value.
' -----------------------------------------------------------------------------------
Private Function GetBuiltinProperty( _
    ByVal props As DocumentProperties, _
    ByVal propName As String) As String
    
    On Error Resume Next
    GetBuiltinProperty = props(propName).value
    If Err.Number <> 0 Then GetBuiltinProperty = ""
    On Error GoTo 0
End Function

' -----------------------------------------------------------------------------------
' Function  : GetRoomId
' Purpose   : Retrieves Room ID from sheet's named range.
' -----------------------------------------------------------------------------------
Private Function GetRoomId(ByVal ws As Worksheet) As String
    On Error Resume Next
    GetRoomId = ws.Range(modConst.NAME_CELL_ROOM_ID).value
    If Err.Number <> 0 Then GetRoomId = ""
    On Error GoTo 0
End Function

' -----------------------------------------------------------------------------------
' Function  : GetRoomAlias
' Purpose   : Retrieves Room Alias from sheet's named range.
' -----------------------------------------------------------------------------------
Private Function GetRoomAlias(ByVal ws As Worksheet) As String
    On Error Resume Next
    GetRoomAlias = ws.Range(modConst.NAME_CELL_ROOM_ALIAS).value
    If Err.Number <> 0 Then GetRoomAlias = ""
    On Error GoTo 0
End Function

' ===== Private Helpers: CSV =======================================================

' -----------------------------------------------------------------------------------
' Procedure : ExportNodesToCsv
' Purpose   : Exports nodes data to CSV file.
' -----------------------------------------------------------------------------------
Private Sub ExportNodesToCsv(ByVal dataSheet As Worksheet, ByVal filePath As String)
    Dim fileNum As Integer
    Dim lastRow As Long
    Dim rowIdx As Long
    Dim nodeID As String, nodeName As String, nodeType As String
    Dim nodeRoom As String, nodeDiff As String, nodeStatus As String
    Dim lineOut As String
    
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    
    ' Write header
    Print #fileNum, CSV_NODES_HEADER
    
    ' Find last row in nodes section (Column A)
    lastRow = dataSheet.Cells(dataSheet.Rows.count, 1).End(xlUp).Row
    
    ' Write data rows
    For rowIdx = 2 To lastRow
        nodeID = Trim$(CStr(dataSheet.Cells(rowIdx, 1).value))
        If LenB(nodeID) = 0 Then GoTo NextNode
        
        nodeName = CsvEscape(CStr(dataSheet.Cells(rowIdx, 2).value))
        nodeType = CStr(dataSheet.Cells(rowIdx, 3).value)
        nodeRoom = CStr(dataSheet.Cells(rowIdx, 4).value)
        nodeDiff = CStr(dataSheet.Cells(rowIdx, 5).value)
        nodeStatus = CStr(dataSheet.Cells(rowIdx, 6).value)
        
        lineOut = nodeID & "," & nodeName & "," & nodeType & "," & _
                  nodeRoom & "," & nodeDiff & "," & nodeStatus
        Print #fileNum, lineOut
NextNode:
    Next rowIdx
    
    Close #fileNum
End Sub

' -----------------------------------------------------------------------------------
' Procedure : ExportEdgesToCsv
' Purpose   : Exports edges data to CSV file.
' -----------------------------------------------------------------------------------
Private Sub ExportEdgesToCsv(ByVal dataSheet As Worksheet, ByVal filePath As String)
    Dim fileNum As Integer
    Dim lastRow As Long
    Dim rowIdx As Long
    Dim fromID As String, toID As String, edgeType As String, edgeNotes As String
    Dim lineOut As String
    
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    
    ' Write header
    Print #fileNum, CSV_EDGES_HEADER
    
    ' Find last row in edges section (Column H)
    lastRow = dataSheet.Cells(dataSheet.Rows.count, 8).End(xlUp).Row
    
    ' Write data rows (edges start at column H=8)
    For rowIdx = 2 To lastRow
        fromID = Trim$(CStr(dataSheet.Cells(rowIdx, 9).value))  ' Column I
        If LenB(fromID) = 0 Then GoTo NextEdge
        
        toID = CStr(dataSheet.Cells(rowIdx, 10).value)      ' Column J
        edgeType = CStr(dataSheet.Cells(rowIdx, 11).value)  ' Column K
        edgeNotes = CsvEscape(CStr(dataSheet.Cells(rowIdx, 12).value))  ' Column L
        
        lineOut = fromID & "," & toID & "," & edgeType & "," & edgeNotes
        Print #fileNum, lineOut
NextEdge:
    Next rowIdx
    
    Close #fileNum
End Sub

' -----------------------------------------------------------------------------------
' Function  : CsvEscape
' Purpose   : Escapes a string for CSV format (handles commas, quotes, newlines).
' -----------------------------------------------------------------------------------
Private Function CsvEscape(ByVal text As String) As String
    ' If contains comma, quote, or newline, wrap in quotes and escape internal quotes
    If InStr(text, ",") > 0 Or InStr(text, """") > 0 Or InStr(text, vbCr) > 0 Or InStr(text, vbLf) > 0 Then
        CsvEscape = """" & Replace(text, """", """""") & """"
    Else
        CsvEscape = text
    End If
End Function
