Attribute VB_Name = "modLists"
Option Explicit
Option Private Module

' -----------------------------------------------------------------------------------
' Function  : GetNamedOrHeaderValue
' Purpose   : Retrieves a value from a named range if available,
'             otherwise searches for a matching header in the worksheet and returns
'             the adjacent cell value to the right.
'
' Parameters:
'   wks          [Worksheet] - The worksheet to search
'   strLocalName [String]    - The name of the named range to check first
'   vntHeaders   [Variant]   - Array of header names to search for (case-insensitive)
'
' Returns   : The trimmed string value from the named range or the cell next to
'             the first matching header found. Returns an empty string if no match.
'
' Behavior  :
'   - Tries to read from the named range first
'   - If empty, scans the top 50 rows and 20 columns for a header match
'   - If a match is found, returns the value in the cell to the right
'
' Notes     :
'   - Header comparison is case-insensitive using StrComp
'   - Assumes headers are in row/column format and values are to the right
'   - No error raised if named range is missing or header not found
' -----------------------------------------------------------------------------------
Public Function GetNamedOrHeaderValue(ByVal wks As Worksheet, ByVal strLocalName As String, ByVal vntHeaders As Variant) As String
    On Error Resume Next
    Dim vntValue As Variant: vntValue = wks.Range(strLocalName).Value
    On Error GoTo 0
    If Len(Trim$(CStr(vntValue))) > 0 Then
        GetNamedOrHeaderValue = Trim$(CStr(vntValue))
        Exit Function
    End If
    Dim lngRow As Long, lngCol As Long
    For lngRow = 1 To 50
        For lngCol = 1 To 20
            Dim strHeader As String: strHeader = Trim$(CStr(wks.Cells(lngRow, lngCol).Value))
            If Len(strHeader) > 0 Then
                Dim lngIndex As Long
                For lngIndex = LBound(vntHeaders) To UBound(vntHeaders)
                    If StrComp(strHeader, vntHeaders(lngIndex), vbTextCompare) = 0 Then
                        GetNamedOrHeaderValue = Trim$(CStr(wks.Cells(lngRow, lngCol + 1).Value))
                        Exit Function
                    End If
                Next lngIndex
            End If
        Next lngCol
    Next lngRow
End Function

' -----------------------------------------------------------------------------------
' Procedure : CollectColumnValues
' Purpose   : Searches for a column with a matching header and collects all
'             non-empty values below it into a dictionary.
'
' Parameters:
'   wks             [Worksheet] - The worksheet to scan
'   vntHeaderNames  [Variant]   - Array of possible header names to match (case-insensitive)
'   dicValues       [Object]    - Dictionary object to store unique values as keys
'
' Behavior  :
'   - Scans the top 50 rows and columns to find a matching header
'   - Once found, reads downward from the header cell
'   - Stops reading after 10 consecutive empty cells
'   - Trims values and adds them as dictionary keys (duplicates ignored)
'
' Notes     :
'   - Header comparison is case-insensitive using StrComp
'   - Values are stored as keys with value = True
'   - Designed for performance-limited scans (50x50 cells, 10 empty streak cutoff)
'   - Assumes values are located directly below the header cell
' -----------------------------------------------------------------------------------
Public Sub CollectColumnValues(ByVal wks As Worksheet, ByVal vntHeaderNames As Variant, ByVal dicValues As Object)
    Dim lngRow As Long, lngCol As Long, lngFoundCol As Long
    For lngRow = 1 To 50
        For lngCol = 1 To 50
            Dim strHeader As String: strHeader = Trim$(CStr(wks.Cells(lngRow, lngCol).Value))
            If Len(strHeader) = 0 Then GoTo ContinueC
            Dim lngIndex As Long
            For lngIndex = LBound(vntHeaderNames) To UBound(vntHeaderNames)
                If StrComp(strHeader, vntHeaderNames(lngIndex), vbTextCompare) = 0 Then
                    lngFoundCol = lngCol
                    Exit For
                End If
            Next lngIndex
            If lngFoundCol > 0 Then Exit For
ContinueC:
        Next lngCol
        If lngFoundCol > 0 Then Exit For
    Next lngRow
    If lngFoundCol = 0 Then Exit Sub
    Dim lngReadRow As Long, lngEmptyStreak As Long
    lngReadRow = lngRow + 1
    Do While lngReadRow <= wks.Rows.Count And lngEmptyStreak < 10
        Dim strValue As String: strValue = Trim$(CStr(wks.Cells(lngReadRow, lngFoundCol).Value))
        If Len(strValue) = 0 Then
            lngEmptyStreak = lngEmptyStreak + 1
        Else
            lngEmptyStreak = 0
            dicValues(strValue) = True
        End If
        lngReadRow = lngReadRow + 1
    Loop
End Sub

' -----------------------------------------------------------------------------------
' Procedure : CollectColumnBlockGroupValues
' Purpose   : Extracts non-empty values from multiple categorized column blocks
'             in a worksheet and stores them in a dictionary, tagged by category.
'
' Parameters:
'   wks                   [Worksheet] - The worksheet containing the structured table
'   lngStartRow           [Long]      - The row number where category begins (Header row)
'   lngEndRow             [Long]      - The row number where category ends
'   vntCategoryHeaders    [Variant]   - Array of category names (e.g. "PICKUPABLE OBJECTS")
'   lngColumnsPerCategory [Long]      - Number of columns assigned to each category block
'   dicValues             [Object]    - Dictionary to store values as keys, with category name as value
'
' Behavior:
'   - Iterates through each category block, defined by its header and column span
'   - Scans downward from the header row, collecting all non-empty cell values
'   - Each value is trimmed and added to the dictionary with its corresponding category
'
' Notes:
'   - Assumes categories are laid out horizontally in adjacent column blocks
'   - Designed for structured layouts like interaction tables (e.g. VR, UI design)
'   - Duplicate values across categories will overwrite previous entries
' -----------------------------------------------------------------------------------
Public Sub CollectColumnBlockGroupValues( _
    ByVal wks As Worksheet, _
    ByVal lngStartRow As Long, _
    ByVal lngEndRow As Long, _
    ByVal vntCategoryHeaders As Variant, _
    ByVal lngColumnsPerCategory As Long, _
    ByVal dicValues As Object)

    Dim lngCatIndex As Long, lngColOffset As Long
    Dim lngRow As Long, lngCol As Long
    Dim lngHeaderRow As Long: lngHeaderRow = lngStartRow

    If lngEndRow <= lngStartRow Then Exit Sub
    
    ' Loop over all categories
    For lngCatIndex = LBound(vntCategoryHeaders) To UBound(vntCategoryHeaders)
        lngColOffset = (lngCatIndex * lngColumnsPerCategory) + 1 ' Start column for category

        ' Loop over all columns in the category
        For lngCol = lngColOffset To lngColOffset + lngColumnsPerCategory - 1
            ' Loop over the lines below the header line until endRow
            For lngRow = lngHeaderRow + 1 To lngEndRow
                Dim strValue As String: strValue = Trim$(CStr(wks.Cells(lngRow, lngCol).Value))
                If Len(strValue) > 0 Then
                    dicValues(strValue) = vntCategoryHeaders(lngCatIndex)
                End If
            Next lngRow
        Next lngCol
    Next lngCatIndex
End Sub

' -----------------------------------------------------------------------------------
' Procedure : UpdateNamedListRange
' Purpose   : Updates or creates a named range in the workbook referring to a
'             column of values starting from row 2 to the last used row.
'
' Parameters:
'   strName [String]     - The name of the named range to update/create
'   wks     [Worksheet]  - The worksheet containing the range
'   lngCol  [Long]       - The column number containing the list
'
' Notes:
'   - If the named range already exists, it will be updated
'   - If it doesn't exist, it will be created
' -----------------------------------------------------------------------------------
Public Sub UpdateNamedListRange(ByVal strName As String, ByVal wks As Worksheet, ByVal lngCol As Long)
    Dim lngLastRow As Long
    Dim wbActive As Workbook
    
    Set wbActive = ActiveWorkbook
    lngLastRow = wks.Cells(wks.Rows.Count, lngCol).End(xlUp).Row
    If lngLastRow < 2 Then lngLastRow = 2
    Dim strRef As String
    strRef = "='" & wks.Name & "'!" & wks.Range(wks.Cells(2, lngCol), wks.Cells(lngLastRow, lngCol)).Address
    On Error Resume Next
    wbActive.names(strName).RefersTo = strRef
    If Err.Number <> 0 Then
        wbActive.names.Add Name:=strName, RefersTo:=strRef
    End If
    On Error GoTo 0
End Sub


' -----------------------------------------------------------------------------------
' Procedure : WriteDictSetToColumn
' Purpose   : Writes the contents of a dictionary (set) to a specified column in
'             the worksheet, sorted alphabetically.
'
' Parameters:
'   wks        [Worksheet] - The worksheet to write into
'   dicSet     [Object]    - Dictionary containing the values (keys only)
'   lngStartRow[Long]      - The row number where writing begins
'   lngCol     [Long]      - The column number where values are written
'
' Notes:
'   - Performs a simple bubble sort for ordering
'   - Assumes dictionary keys are strings
' -----------------------------------------------------------------------------------
Public Sub WriteDictSetToColumn(ByVal wks As Worksheet, ByVal dicSet As Object, ByVal lngStartRow As Long, ByVal lngCol As Long)
    Dim astrValues() As String, vntKey As Variant, lngIndex As Long
    If dicSet.Count = 0 Then Exit Sub
    ReDim astrValues(1 To dicSet.Count)
    lngIndex = 1
    For Each vntKey In dicSet.Keys
        astrValues(lngIndex) = CStr(vntKey)
        lngIndex = lngIndex + 1
    Next vntKey
    ' simple bubble sort
    Dim blnSwapped As Boolean, strTemp As String
    Do
        blnSwapped = False
        For lngIndex = LBound(astrValues) To UBound(astrValues) - 1
            If astrValues(lngIndex) > astrValues(lngIndex + 1) Then
                strTemp = astrValues(lngIndex): astrValues(lngIndex) = astrValues(lngIndex + 1): astrValues(lngIndex + 1) = strTemp
                blnSwapped = True
            End If
        Next lngIndex
    Loop While blnSwapped
    For lngIndex = 1 To UBound(astrValues)
        wks.Cells(lngStartRow + lngIndex - 1, lngCol).Value = astrValues(lngIndex)
    Next lngIndex
End Sub

' -----------------------------------------------------------------------------------
' Procedure : AppendMissingDictKeysToColumn
' Purpose   : Appends keys from newKeys to a column if they do not already exist
'             in the supplied existing dictionary.
'
' Parameters:
'   wks         [Worksheet]       - Target worksheet.
'   lngCol      [Long]            - Target column number to append into.
'   dicExisting [Object]          - Dictionary containing existing values (keys only).
'   dicNewKeys  [Object]          - Dictionary containing new values to append (keys only).
'   lngStartRow [Long]            - Optional, First data row. (default 2).
' -----------------------------------------------------------------------------------
Public Sub AppendMissingDictKeysToColumn( _
    ByVal wks As Worksheet, _
    ByVal lngCol As Long, _
    ByVal dicExisting As Object, _
    ByVal dicNewKeys As Object, _
    Optional ByVal lngStartRow As Long = 2)

    Dim lngLastRow As Long, lngNextRow As Long
    Dim vntKey As Variant, strValue As String

    lngLastRow = wks.Cells(wks.Rows.Count, lngCol).End(xlUp).Row
    If lngLastRow < lngStartRow Then
        lngNextRow = lngStartRow
    Else
        lngNextRow = lngLastRow + 1
    End If

    For Each vntKey In dicNewKeys.Keys
        strValue = Trim$(CStr(vntKey))
        If Len(strValue) > 0 Then
            If Not dicExisting.Exists(strValue) Then
                wks.Cells(lngNextRow, lngCol).Value = strValue
                dicExisting(strValue) = True
                lngNextRow = lngNextRow + 1
            End If
        End If
    Next vntKey
End Sub

