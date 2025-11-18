Attribute VB_Name = "modLists"
' -----------------------------------------------------------------------------------
' Module: modLists
' Purpose: Helper functions for collecting, normalizing, and writing list-like
'          data from worksheets, named ranges, and structured tables into
'          dictionaries and ranges. Intended mainly for validation lists and
'          configuration-style data.
'
' Public API:
'   - GetNamedOrHeaderValue
'   - CollectColumnValues
'   - CollectColumnValuesFromRange
'   - CollectColumnBlockGroupValues
'   - CollectNamedRangeValues
'   - CollectTableColumnValues
'   - UpdateNamedListRange
'   - WriteDictSetToColumn
'   - WriteDictSetToTableColumn
'   - AppendMissingDictKeysToColumn
'   - AppendMissingDictKeysToTableColumn
'   - ClearTableColumn
'
' Notes:
'   - Option Private Module keeps the module internal to the project.
' -----------------------------------------------------------------------------------
Option Explicit
Option Private Module

' ===== Public API ===================================================================

' -----------------------------------------------------------------------------------
' Function  : GetNamedOrHeaderValue
' Purpose   : Retrieves a value from a named range if available,
'             otherwise searches for a matching header in the worksheet and returns
'             the adjacent cell value to the right.
'
' Parameters:
'   sheet          [Worksheet] - The worksheet to search
'   localName      [String]    - The name of the named range to check first
'   headerArray    [Variant]   - Array of header names to search for (case-insensitive)
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
Public Function GetNamedOrHeaderValue(ByVal sheet As Worksheet, ByVal localName As String, ByVal headerArray As Variant) As String
    On Error GoTo ErrHandler

    Dim namedValue As Variant
    Dim searchRange As Range
    Dim headerCell As Range
    Dim bestHeaderCell As Range
    Dim headerIndex As Long
    Dim firstAddress As String

    ' Try named range first (silent failure if missing)
    On Error Resume Next
    namedValue = sheet.Range(localName).Value
    On Error GoTo ErrHandler

    If Len(Trim$(CStr(namedValue))) > 0 Then
        GetNamedOrHeaderValue = Trim$(CStr(namedValue))
        GoTo CleanExit
    End If

    ' Search for headers in A1:T50
    Set searchRange = sheet.Range(sheet.Cells(1, 1), sheet.Cells(50, 20))

    For headerIndex = LBound(headerArray) To UBound(headerArray)
        Set headerCell = searchRange.Find(What:=headerArray(headerIndex), _
                                          LookIn:=xlValues, _
                                          LookAt:=xlWhole, _
                                          MatchCase:=False, _
                                          SearchOrder:=xlByRows, _
                                          SearchDirection:=xlNext)
        If Not headerCell Is Nothing Then
            firstAddress = headerCell.Address
            Do
                ' Track the top-most, then left-most match across all headers
                If bestHeaderCell Is Nothing Then
                    Set bestHeaderCell = headerCell
                ElseIf headerCell.Row < bestHeaderCell.Row Or _
                       (headerCell.Row = bestHeaderCell.Row And headerCell.column < bestHeaderCell.column) Then
                    Set bestHeaderCell = headerCell
                End If
                Set headerCell = searchRange.FindNext(headerCell)
            Loop While Not headerCell Is Nothing And headerCell.Address <> firstAddress
        End If
    Next headerIndex

    If Not bestHeaderCell Is Nothing Then
        GetNamedOrHeaderValue = Trim$(CStr(bestHeaderCell.Offset(0, 1).Value))
    End If

CleanExit:
    Exit Function
ErrHandler:
    modErr.ReportError "GetNamedOrHeaderValue", Err.Number, Erl, caption:=modMain.AppProjectName
    Resume CleanExit
End Function

' -----------------------------------------------------------------------------------
' Procedure : CollectColumnValues
' Purpose   : Searches for a column with a matching header and collects all
'             non-empty values below it into a dictionary.
'
' Parameters:
'   sheet             [Worksheet] - The worksheet to scan
'   headerArray    [Variant]   - Array of possible header names to match (case-insensitive)
'   valuesDict      [Scripting.Dictionary] - Dictionary object to store unique values as keys
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
Public Sub CollectColumnValues(ByVal sheet As Worksheet, ByVal headerArray As Variant, ByVal valuesDict As Scripting.Dictionary)
    On Error GoTo ErrHandler

    Dim searchRange As Range
    Dim headerCell As Range
    Dim bestHeaderCell As Range
    Dim headerIndex As Long
    Dim firstAddress As String
    Dim headerRow As Long
    Dim foundColumnIndex As Long
    Dim columnRange As Range

    ' Search for headers in A1:AX50 (50 columns)
    Set searchRange = sheet.Range(sheet.Cells(1, 1), sheet.Cells(50, 50))

    For headerIndex = LBound(headerArray) To UBound(headerArray)
        Set headerCell = searchRange.Find(What:=headerArray(headerIndex), _
                                          LookIn:=xlValues, _
                                          LookAt:=xlWhole, _
                                          MatchCase:=False, _
                                          SearchOrder:=xlByRows, _
                                          SearchDirection:=xlNext)
        If Not headerCell Is Nothing Then
            firstAddress = headerCell.Address
            Do
                ' Track the top-most, then left-most match across all headers
                If bestHeaderCell Is Nothing Then
                    Set bestHeaderCell = headerCell
                ElseIf headerCell.Row < bestHeaderCell.Row Or _
                       (headerCell.Row = bestHeaderCell.Row And headerCell.column < bestHeaderCell.column) Then
                    Set bestHeaderCell = headerCell
                End If
                Set headerCell = searchRange.FindNext(headerCell)
            Loop While Not headerCell Is Nothing And headerCell.Address <> firstAddress
        End If
    Next headerIndex

    If bestHeaderCell Is Nothing Then GoTo CleanExit

    headerRow = bestHeaderCell.Row
    foundColumnIndex = bestHeaderCell.column

    ' Build range from first data row below header down to the end of the sheet
    Set columnRange = sheet.Range( _
        sheet.Cells(headerRow + 1, foundColumnIndex), _
        sheet.Cells(sheet.Rows.Count, foundColumnIndex) _
    )

     ExtractColumnValues columnRange, valuesDict, 10

CleanExit:
    Exit Sub
ErrHandler:
    modErr.ReportError "CollectColumnValues", Err.Number, Erl, caption:=modMain.AppProjectName
    Resume CleanExit
End Sub

' -----------------------------------------------------------------------------------
' Procedure : CollectColumnValuesFromRange
' Purpose   : Reads all non-empty values from a given column range
'             and stores them as unique keys in the dictionary.
'
' Parameters:
'   column  [Range]  - The column range (e.g., from GetTableColumnRange)
'   valuesDict   [Scripting.Dictionary] - Dictionary for storing the unique values
'
' Behavior  :
'   - Reads all cells in the range from top to bottom
'   - Ignores empty cells
'   - Inserts trimmed values as keys with value = True into the dictionary
'   - Stops after 10 consecutive empty cells
'
' Notes     :
'   - No header search necessary
'   - Range can be, for example, a table column (DataBodyRange)
' ------------------------------------------------------- ----------------------------
Public Sub CollectColumnValuesFromRange(ByVal column As Range, ByVal valuesDict As Scripting.Dictionary)
    On Error GoTo ErrHandler

    ExtractColumnValues column, valuesDict

CleanExit:
    Exit Sub
ErrHandler:
    modErr.ReportError "CollectColumnValuesFromRange", Err.Number, Erl, caption:=modMain.AppProjectName
    Resume CleanExit
End Sub

' -----------------------------------------------------------------------------------
' Procedure : CollectColumnBlockGroupValues
' Purpose   : Extracts non-empty values from multiple categorized column blocks
'             in a worksheet and stores them in a dictionary, tagged by category.
'
' Parameters:
'   sheet                   [Worksheet] - The worksheet containing the structured table
'   startRow              [Long]      - The row number where category begins (Header row)
'   endRow                [Long]      - The row number where category ends
'   categoryHeadersArray  [Variant]   - Array of category names (e.g. "PICKUPABLE OBJECTS")
'   columnsPerCategory    [Long]      - Number of columns assigned to each category block
'   valuesDict            [Scripting.Dictionary]    - Dictionary to store values as keys, with category name as value
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
    ByVal sheet As Worksheet, _
    ByVal startRow As Long, _
    ByVal endRow As Long, _
    ByVal categoryHeadersArray As Variant, _
    ByVal columnsPerCategory As Long, _
    ByVal valuesDict As Scripting.Dictionary)

    On Error GoTo ErrHandler

    Dim categoryIndex As Long
    Dim categoryColumnOffset As Long
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim headerRow As Long
    Dim cellValue As String

    headerRow = startRow

    If endRow <= startRow Then GoTo CleanExit

    ' Loop over all categories
    For categoryIndex = LBound(categoryHeadersArray) To UBound(categoryHeadersArray)
        categoryColumnOffset = (categoryIndex * columnsPerCategory) + 1 ' Start column for category

        ' Loop over all columns in the category
        For colIndex = categoryColumnOffset To categoryColumnOffset + columnsPerCategory - 1
            ' Loop over the lines below the header line until endRow
            For rowIndex = headerRow + 1 To endRow
                cellValue = Trim$(CStr(sheet.Cells(rowIndex, colIndex).Value))
                If Len(cellValue) > 0 Then
                    valuesDict(cellValue) = categoryHeadersArray(categoryIndex)
                End If
            Next rowIndex
        Next colIndex
    Next categoryIndex

CleanExit:
    Exit Sub
ErrHandler:
    modErr.ReportError "CollectColumnBlockGroupValues", Err.Number, Erl, caption:=modMain.AppProjectName
    Resume CleanExit
End Sub

' -----------------------------------------------------------------------------------
' Procedure : CollectNamedRangeValues
' Purpose   : Collects all non-empty values from a named range on a worksheet
'             and stores them as unique keys in a dictionary.
'
' Parameters:
'   sheet          [Worksheet] - The worksheet containing the named range.
'   rangeName    [String]    - The name of the range to resolve.
'   valuesDict   [Scripting.Dictionary]    - Dictionary receiving the unique values.
'
' Behavior  :
'   - Resolves the named range on the given worksheet.
'   - Ignores empty cells.
'   - Adds each value as a key to the dictionary (duplicates ignored).
'
' Notes     :
'   - If the named range is missing or invalid, the procedure exits silently.
' -----------------------------------------------------------------------------------
Public Sub CollectNamedRangeValues(ByVal sheet As Worksheet, rangeName As String, ByVal valuesDict As Scripting.Dictionary)
    On Error GoTo ErrHandler

    Dim rngNamed As Range
    Dim cell As Range

    If Len(rangeName) = 0 Then GoTo CleanExit

    On Error Resume Next
    Set rngNamed = sheet.Range(rangeName)
    On Error GoTo ErrHandler

    If rngNamed Is Nothing Then GoTo CleanExit

    For Each cell In rngNamed.Cells
        If Not IsEmpty(cell.Value) Then
            If Not valuesDict.Exists(cell.Value) Then
                valuesDict.Add cell.Value, True
            Else
                valuesDict(cell.Value) = True
            End If
        End If
    Next cell

CleanExit:
    Exit Sub
ErrHandler:
    modErr.ReportError "CollectNamedRangeValues", Err.Number, Erl, caption:=modMain.AppProjectName
    Resume CleanExit
End Sub

' -----------------------------------------------------------------------------------
' Procedure : CollectTableColumnValues
' Purpose   : Collects all non-empty values from a structured table column into a dictionary.
'
' Parameters:
'   sourceTable        [ListObject] - The structured table containing the column.
'   columnName         [String]     - The name of the column to read from.
'   valuesDict         [Scripting.Dictionary]     - Dictionary object to store unique values as keys.
'
' Behavior  :
'   - Reads all cells in the specified column's DataBodyRange
'   - Trims values and adds them as dictionary keys (duplicates ignored)
'   - Stops reading after 10 consecutive empty cells
'
' Notes     :
'   - Values are stored as keys with value = True
'   - Assumes values are strings
' -----------------------------------------------------------------------------------
Public Sub CollectTableColumnValues(ByVal sourceTable As ListObject, ByVal columnName As String, ByVal valuesDict As Scripting.Dictionary)
    On Error GoTo ErrHandler

    Dim column As Range
    Dim cell As Range
    Dim emptyStreak As Long
    Dim cellValue As String

    Set column = sourceTable.ListColumns(columnName).DataBodyRange

    ExtractColumnValues column, valuesDict

CleanExit:
    Exit Sub
ErrHandler:
    modErr.ReportError "CollectTableColumnValues", Err.Number, Erl, caption:=modMain.AppProjectName
    Resume CleanExit
End Sub


' -----------------------------------------------------------------------------------
' Procedure : UpdateNamedListRange
' Purpose   : Updates or creates a named range in the workbook referring to a
'             column of values starting from row 2 to the last used row.
'
' Parameters:
'   rangeName  [String]     - The name of the named range to update/create
'   sheet        [Worksheet]  - The worksheet containing the range
'   columnIdx  [Long]       - The column number containing the list
'
' Notes:
'   - If the named range already exists, it will be updated
'   - If it doesn't exist, it will be created
' -----------------------------------------------------------------------------------
Public Sub UpdateNamedListRange(ByVal rangeName As String, ByVal sheet As Worksheet, ByVal columnIdx As Long)
    On Error GoTo ErrHandler

    Dim lastRow As Long
    Dim wb As Workbook
    Dim reference As String

    Set wb = ActiveWorkbook
    lastRow = sheet.Cells(sheet.Rows.Count, columnIdx).End(xlUp).Row
    If lastRow < 2 Then lastRow = 2

    reference = "='" & sheet.Name & "'!" & sheet.Range(sheet.Cells(2, columnIdx), sheet.Cells(lastRow, columnIdx)).Address

    On Error Resume Next
    wb.Names(rangeName).RefersTo = reference
    If Err.Number <> 0 Then
        wb.Names.Add Name:=rangeName, RefersTo:=reference
    End If
    On Error GoTo ErrHandler

CleanExit:
    Exit Sub
ErrHandler:
    MsgBox "Error (" & Err.Number & "): " & Err.Description, vbCritical, "UpdateNamedListRange"
    Resume CleanExit
End Sub


' -----------------------------------------------------------------------------------
' Procedure : WriteDictSetToColumn
' Purpose   : Writes the contents of a dictionary (set) to a specified column in
'             the worksheet, sorted alphabetically.
'
' Parameters:
'   sheet                     [Worksheet] - The worksheet to write into
'   valuesDict              [Scripting.Dictionary]    - Dictionary containing the keys/values
'   startRow                [Long]      - The row number where writing begins
'   columnIdx               [Long]      - The column number where keys are written (values added as cell comment)
'   writeValuesToNextColumn [Boolean]   - (optional) Writes the Values of valuesDict to the next column. default = False
' Notes:
'   - Performs a simple quick sort for ordering
'   - Assumes dictionary keys and values are strings
' -----------------------------------------------------------------------------------
Public Sub WriteDictSetToColumn(ByVal sheet As Worksheet, ByVal valuesDict As Scripting.Dictionary, ByVal startRow As Long, ByVal columnIdx As Long, Optional ByVal writeValuesToNextColumn As Boolean = False)
    On Error GoTo ErrHandler

    Dim sortedKeys As Collection
    Dim keyArray() As String
    Dim currentKey As Variant
    Dim index As Long

    Set sortedKeys = New Collection

    ' Copy keys to array
    ReDim keyArray(0 To valuesDict.Count - 1)
    index = 0
    For Each currentKey In valuesDict.Keys
        keyArray(index) = CStr(currentKey)
        index = index + 1
    Next currentKey

    ' Sort array
    modUtil.QuickSortStringArray keyArray, LBound(keyArray), UBound(keyArray)

    ' Load sorted keys into collection
    For index = LBound(keyArray) To UBound(keyArray)
        sortedKeys.Add keyArray(index)
    Next index

    ' Add to lists
    For index = 1 To sortedKeys.Count
        sheet.Cells(startRow + index - 1, columnIdx).Value = sortedKeys(index)
        If writeValuesToNextColumn Then
            sheet.Cells(startRow + index - 1, columnIdx + 1).Value = valuesDict(sortedKeys(index))
        End If
    Next index

CleanExit:
    Exit Sub
ErrHandler:
    MsgBox "Error (" & Err.Number & "): " & Err.Description, vbCritical, "WriteDictSetToColumn"
    Resume CleanExit
End Sub

' -----------------------------------------------------------------------------------
' Procedure : WriteDictSetToTableColumn
' Purpose   : Writes the contents of a dictionary (set) into a structured table column,
'             sorted alphabetically, and expands the table if needed.
'
' Parameters:
'   targetTable              [ListObject] - The structured table to write into
'   columnName               [String]     - Name of the column to write into
'   valuesDict               [Scripting.Dictionary]     - Dictionary containing the keys/values
'   writeValuesToNextColumn  [Boolean]    - (optional) Writes the Values of valuesDict to the next column. default = False
'
' Notes:
'   - Automatically adds rows to the table if valuesDict.Count > current row count
'   - Assumes dictionary keys and values are strings
' -----------------------------------------------------------------------------------
Public Sub WriteDictSetToTableColumn(ByVal targetTable As ListObject, ByVal columnName As String, ByVal valuesDict As Scripting.Dictionary, Optional ByVal writeValuesToNextColumn As Boolean = False)
    On Error GoTo ErrHandler

    Dim sortedKeys As Collection
    Dim keyArray() As String
    Dim currentKey As Variant
    Dim index As Long
    Dim targetColumn As Range

    Set sortedKeys = New Collection

    ' Copy keys to array
    ReDim keyArray(0 To valuesDict.Count - 1)
    index = 0
    For Each currentKey In valuesDict.Keys
        keyArray(index) = CStr(currentKey)
        index = index + 1
    Next currentKey

    ' Sort array
    modUtil.QuickSortStringArray keyArray, LBound(keyArray), UBound(keyArray)

    ' Load sorted keys into collection
    For index = LBound(keyArray) To UBound(keyArray)
        sortedKeys.Add keyArray(index)
    Next index

    ' Ensure table has enough rows
    Do While targetTable.ListRows.Count < sortedKeys.Count
        targetTable.ListRows.Add
    Loop

    ' Write values into table column
    Set targetColumn = targetTable.ListColumns(columnName).DataBodyRange

    For index = 1 To sortedKeys.Count
        targetColumn.Cells(index, 1).Value = sortedKeys(index)
        If writeValuesToNextColumn Then
            targetColumn.Cells(index, 1).Offset(0, 1).Value = valuesDict(sortedKeys(index))
        End If
    Next index

CleanExit:
    Exit Sub
ErrHandler:
    MsgBox "Error (" & Err.Number & "): " & Err.Description, vbCritical, "WriteDictSetToTableColumn"
    Resume CleanExit
End Sub


' -----------------------------------------------------------------------------------
' Procedure : AppendMissingDictKeysToColumn
' Purpose   : Appends keys from newKeys to a column if they do not already exist
'             in the supplied existing dictionary.
'
' Parameters:
'   sheet              [Worksheet]          - Target worksheet.
'   columnIdx        [Long]               - Target column number to append into.
'   existingKeysDict [Scripting.Dictionary]          - Dictionary containing existing values (keys only).
'   newKeysDict      [Scripting.Dictionary]          - Dictionary containing new values to append (keys only).
'   StartRow         [Long]                  - Optional, First data row. (default 2).
' -----------------------------------------------------------------------------------
Public Sub AppendMissingDictKeysToColumn( _
    ByVal sheet As Worksheet, _
    ByVal columnIdx As Long, _
    ByVal existingKeysDict As Scripting.Dictionary, _
    ByVal newKeysDict As Scripting.Dictionary, _
    Optional ByVal startRow As Long = 2)

    On Error GoTo ErrHandler

    Dim lastRow As Long
    Dim nextRow As Long
    Dim currentKey As Variant
    Dim valueText As String

    lastRow = sheet.Cells(sheet.Rows.Count, columnIdx).End(xlUp).Row
    If lastRow < startRow Then
        nextRow = startRow
    Else
        nextRow = lastRow + 1
    End If

    For Each currentKey In newKeysDict.Keys
        valueText = Trim$(CStr(currentKey))
        If Len(valueText) > 0 Then
            If Not existingKeysDict.Exists(valueText) Then
                sheet.Cells(nextRow, columnIdx).Value = valueText
                existingKeysDict(valueText) = True
                nextRow = nextRow + 1
            End If
        End If
    Next currentKey

CleanExit:
    Exit Sub
ErrHandler:
    MsgBox "Error (" & Err.Number & "): " & Err.Description, vbCritical, "AppendMissingDictKeysToColumn"
    Resume CleanExit
End Sub

' -----------------------------------------------------------------------------------
' Procedure : AppendMissingDictKeysToTableColumn
' Purpose   : Appends keys from newKeysDict to a structured table column if they
'             do not already exist in existingKeysDict. Expands the table as needed.
'
' Parameters:
'   targetTable      [ListObject]               - The structured table to write into.
'   columnName       [String]                   - The name of the column to append into.
'   existingKeysDict [Scripting.Dictionary]     - Dictionary containing existing values (keys only).
'   newKeysDict      [Scripting.Dictionary]     - Dictionary containing new values to append (keys only).
' -----------------------------------------------------------------------------------
Public Sub AppendMissingDictKeysToTableColumn( _
    ByVal targetTable As ListObject, _
    ByVal columnName As String, _
    ByVal existingKeysDict As Scripting.Dictionary, _
    ByVal newKeysDict As Scripting.Dictionary)

    On Error GoTo ErrHandler

    Dim currentKey As Variant
    Dim valueText As String
    Dim targetColumn As Range
    Dim nextRowIndex As Long

    Set targetColumn = targetTable.ListColumns(columnName).DataBodyRange

    nextRowIndex = targetColumn.Rows.Count + 1 ' Start writing after last row

    For Each currentKey In newKeysDict.Keys
        valueText = Trim$(CStr(currentKey))
        If Len(valueText) > 0 Then
            If Not existingKeysDict.Exists(valueText) Then
                targetTable.ListRows.Add ' Add new row to table
                targetTable.DataBodyRange.Cells(nextRowIndex, targetTable.ListColumns(columnName).index).Value = valueText
                existingKeysDict(valueText) = True
                nextRowIndex = nextRowIndex + 1
            End If
        End If
    Next currentKey

CleanExit:
    Exit Sub
ErrHandler:
    MsgBox "Error (" & Err.Number & "): " & Err.Description, vbCritical, "AppendMissingDictKeysToTableColumn"
    Resume CleanExit
End Sub

' -----------------------------------------------------------------------------------
' Function  : ClearTableColumn
' Purpose   : Clears the contents of a structured table column that is used
'             as a list/source column.
' Parameters:
'   targetTable        [ListObject] - The table containing the column.
'   columnName         [String]     - The name of the column to clear.
'
' Returns   :
'   Boolean - True if the column was cleared; False if the table is Nothing.
'
' Notes     :
'   - Does not handle errors for invalid column names explicitly; such misuse
'     will raise a runtime error as usual.
' -----------------------------------------------------------------------------------
Public Function ClearTableColumn(ByVal targetTable As ListObject, ByVal columnName As String) As Boolean
    If targetTable Is Nothing Then Exit Function

    targetTable.ListColumns(columnName).DataBodyRange.ClearContents
    ClearTableColumn = True
End Function

' ===== Private Helpers =============================================================

' -----------------------------------------------------------------------------------
' Procedure : ExtractColumnValues
' Purpose   : Internal helper to collect non-empty values from a column-like range
'             into a dictionary, stopping after a configurable number of consecutive
'             empty cells.
'
' Parameters:
'   column    [Range]  - The one-dimensional range (typically a column).
'   valuesDict     [Object] - Dictionary to receive unique values as keys.
'   maxEmptyStreak [Long]   - Optional; number of consecutive empty cells before
'                              stopping (default 10).
'
' Notes     :
'   - Trims cell values before storing.
'   - Values are stored as keys with value = True.
' -----------------------------------------------------------------------------------
Private Sub ExtractColumnValues( _
    ByVal column As Range, _
    ByVal valuesDict As Scripting.Dictionary, _
    Optional ByVal maxEmptyStreak As Long = 10)

    Dim cell As Range
    Dim emptyStreak As Long
    Dim cellValue As String

    emptyStreak = 0

    For Each cell In column.Cells
        cellValue = Trim$(CStr(cell.Value))

        If Len(cellValue) = 0 Then
            emptyStreak = emptyStreak + 1
            If emptyStreak >= maxEmptyStreak Then Exit For
        Else
            emptyStreak = 0
            valuesDict(cellValue) = True
        End If
    Next cell
End Sub
