Attribute VB_Name = "DocuHelper"
Option Explicit

Sub ExportNamedRangesToCSV()
    Dim nm As Name
    Dim wsTemp As Worksheet
    Dim i As Long
    Dim strPath As String
    Dim strFileName As String

    ' Temporäres Arbeitsblatt erstellen
    Set wsTemp = ThisWorkbook.Worksheets.Add
    wsTemp.Name = "NamedRangesExport"

    ' Überschriften
    wsTemp.Cells(1, 1).value = "Name"
    wsTemp.Cells(1, 2).value = "RefersTo"
    wsTemp.Cells(1, 3).value = "Value"

    i = 2
    For Each nm In ThisWorkbook.Names
        wsTemp.Cells(i, 1).value = nm.Name
        wsTemp.Cells(i, 2).value = "'" & nm.RefersTo
        On Error Resume Next
        wsTemp.Cells(i, 3).value = nm.RefersToRange.value
        On Error GoTo 0
        i = i + 1
    Next nm

    ' Speicherort und Dateiname
    strPath = ThisWorkbook.path
    If strPath = "" Then strPath = Application.DefaultFilePath
    strFileName = strPath & "\NamedRangesExport.csv"

    ' Als CSV speichern
    Application.DisplayAlerts = False
    wsTemp.Copy
    ActiveWorkbook.SaveAs fileName:=strFileName, FileFormat:=xlCSV
    ActiveWorkbook.Close False
    Application.DisplayAlerts = True

    ' Temporäres Blatt löschen
    Application.DisplayAlerts = False
    wsTemp.Delete
    Application.DisplayAlerts = True

    MsgBox "Export abgeschlossen: " & strFileName, vbInformation
End Sub
