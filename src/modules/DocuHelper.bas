Attribute VB_Name = "DocuHelper"
Option Explicit

Sub ExportNamedRangesToCSV()
    Dim nm As name
    Dim tmpSheet As Worksheet
    Dim i As Long
    Dim filePath As String
    Dim fileNAme As String

    ' Create temporary worksheet
    Set tmpSheet = ActiveWorkbook.Worksheets.Add
    tmpSheet.name = "NamedRangesExport"

    ' Überschriften
    tmpSheet.Cells(1, 1).value = "Name"
    tmpSheet.Cells(1, 2).value = "RefersTo"
    tmpSheet.Cells(1, 3).value = "Value"

    i = 2
    For Each nm In ActiveWorkbook.Names
        tmpSheet.Cells(i, 1).value = nm.name
        tmpSheet.Cells(i, 2).value = "'" & nm.RefersTo
        On Error Resume Next
        tmpSheet.Cells(i, 3).value = nm.RefersToRange.value
        On Error GoTo 0
        i = i + 1
    Next nm

    ' Speicherort und Dateiname
    filePath = ActiveWorkbook.path
    If filePath = "" Then filePath = Application.DefaultFilePath
    fileNAme = filePath & "\NamedRangesExport.csv"

    ' Als CSV speichern
    Application.DisplayAlerts = False
    tmpSheet.Copy
    ActiveWorkbook.SaveAs fileNAme:=fileNAme, FileFormat:=xlCSV
    ActiveWorkbook.Close False
    Application.DisplayAlerts = True

    ' Temporäres Blatt löschen
    Application.DisplayAlerts = False
    tmpSheet.Delete
    Application.DisplayAlerts = True

    MsgBox "Export abgeschlossen: " & fileNAme, vbInformation
End Sub

Sub ExportCommentedCellsToCSV()
    Dim sheet As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim dict As Object
    Dim filePath As String, fileNAme As String
    Dim fso As Object, ts As Object
    Dim key As Variant
    Dim cellCmnt As String
    Dim cellValue As String
    Dim cellData As String

    Set sheet = ActiveSheet
    Set rng = sheet.UsedRange
    Set dict = CreateObject("Scripting.Dictionary")

    'Collect cells with classic comments (notes)
    For Each cell In rng
        If Not cell.Comment Is Nothing Then
            cellCmnt = ""
            On Error Resume Next
            cellCmnt = CStr(cell.Comment.text)
            If Err.Number <> 0 Then
                cellCmnt = ""   ' unexpected value occurs
                Err.Clear
            End If
            On Error GoTo 0
            cellValue = CStr(cell.value)
            dict(cell.Address) = cellValue & ";" & cellCmnt
        End If
    Next cell

    ' Path/file name
    filePath = ActiveWorkbook.path
    If filePath = "" Then filePath = Application.DefaultFilePath
    fileNAme = filePath & "\CommentedCellsExport.csv"

    ' Write Unicode CSV (third parameter = True)
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(fileNAme, True, True)

    ts.WriteLine "CellAddress;Value;Comment"

    For Each key In dict.Keys
        cellData = dict(key)
        ' Catch zero cleanly
        If IsNull(cellData) Then cellData = ";"
        ' Force as string
        cellData = CStr(cellData)
        ' Clean up line breaks/control characters
        cellData = Replace(cellData, vbCrLf, " ")
        cellData = Replace(cellData, vbCr, " ")
        cellData = Replace(cellData, vbLf, " ")
        ' Double quotation marks neutralize
        cellData = Replace(cellData, """", "'")
        ts.WriteLine key & ";" & cellData
    Next key

    ts.Close
    MsgBox "Export abgeschlossen: " & fileNAme, vbInformation
End Sub


