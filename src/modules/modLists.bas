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
' Procedure : CollectNamedRangePairs
' Purpose   : Collects paired values from two named ranges on a worksheet and stores
'             them in a dictionary (first range = Key, second range = Item).
'
' Parameters:
'   sheet           [Worksheet] - The worksheet containing the named ranges.
'   rangeNameKeys   [String]    - The name of the range containing the keys.
'   rangeNameItems  [String]    - The name of the range containing the items.
'   pairsDict       [Scripting.Dictionary] - Dictionary receiving the key-value pairs.
'
' Behavior  :
'   - Resolves both named ranges on the given worksheet.
'   - Reads values from both ranges as arrays (high performance).
'   - Ignores pairs where the key is empty.
'   - If key already exists, updates the item value.
'   - If ranges have different sizes, uses the smaller count.
'
' Notes     :
'   - If either named range is missing or invalid, the procedure exits silently.
'   - Uses array access for optimal performance with large ranges.
'   - Single-cell ranges are handled separately to avoid array issues.
' -----------------------------------------------------------------------------------
Public Sub CollectNamedRangePairs(ByVal sheet As Worksheet, _
                                  ByVal rangeNameKeys As String, _
                                  ByVal rangeNameItems As String, _
                                  ByVal pairsDict As Scripting.Dictionary)
    On Error GoTo ErrHandler
    
    Dim rngKeys As Range
    Dim rngItems As Range
    Dim arrKeys As Variant
    Dim arrItems As Variant
    Dim i As Long
    Dim maxCount As Long
    Dim keyVal As Variant
    Dim itemVal As Variant
    
    ' Validate input
    If Len(rangeNameKeys) = 0 Or Len(rangeNameItems) = 0 Then GoTo CleanExit
    
    ' Resolve named ranges
    On Error Resume Next
    Set rngKeys = sheet.Range(rangeNameKeys)
    Set rngItems = sheet.Range(rangeNameItems)
    On Error GoTo ErrHandler
    
    If rngKeys Is Nothing Or rngItems Is Nothing Then GoTo CleanExit
    
    ' Determine the minimum count to process
    maxCount = Application.WorksheetFunction.Min(rngKeys.Cells.Count, rngItems.Cells.Count)
    If maxCount = 0 Then GoTo CleanExit
    
    ' Load values into arrays for performance
    ' Single cell ranges need special handling
    If rngKeys.Cells.Count = 1 Then
        ReDim arrKeys(1 To 1, 1 To 1)
        arrKeys(1, 1) = rngKeys.Value
    Else
        arrKeys = rngKeys.Value
    End If
    
    If rngItems.Cells.Count = 1 Then
        ReDim arrItems(1 To 1, 1 To 1)
        arrItems(1, 1) = rngItems.Value
    Else
        arrItems = rngItems.Value
    End If
    
    ' Process pairs
    For i = 1 To maxCount
        ' Handle both single-column and multi-column ranges
        If IsArray(arrKeys) Then
            If UBound(arrKeys, 2) > 1 Then
                ' Multi-column range - take first column
                keyVal = arrKeys(i, 1)
            Else
                ' Single-column range
                keyVal = arrKeys(i, 1)
            End If
        Else
            keyVal = arrKeys
        End If
        
        If IsArray(arrItems) Then
            If UBound(arrItems, 2) > 1 Then
                itemVal = arrItems(i, 1)
            Else
                itemVal = arrItems(i, 1)
            End If
        Else
            itemVal = arrItems
        End If
        
        ' Add or update dictionary entry (skip empty keys)
        If Not IsEmpty(keyVal) And Len(CStr(keyVal)) > 0 Then
            pairsDict(keyVal) = itemVal
        End If
    Next i
    
CleanExit:
    Exit Sub
ErrHandler:
    modErr.ReportError "CollectNamedRangePairs", Err.Number, Erl, caption:=modMain.AppProjectName
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
' Procedure : CollectTableColumnPairs
' Purpose   : Collects key-value pairs from two columns of a structured table into a dictionary.
'
' Parameters:
'   sourceTable        [ListObject]               - The structured table containing the columns.
'   keyColumnName      [String]                   - The name of the column containing keys.
'   itemColumnName     [String]                   - The name of the column containing items/values.
'   pairsDict          [Scripting.Dictionary]     - Dictionary to store key-value pairs.
'
' Behavior  :
'   - Reads values from both columns simultaneously
'   - Processes up to the minimum row count of both columns
'   - Skips rows where the key is empty or whitespace
'   - Stores key as dictionary key and item as dictionary value
'   - If key already exists, the value is overwritten
'
' Notes     :
'   - Uses array access for optimal performance with large tables
'   - Trims key values before storing
'   - Item values can be any type (string, number, date, etc.)
'   - If either column doesn't exist, the procedure exits silently
'   - Empty item values are allowed if the key is non-empty
' -----------------------------------------------------------------------------------
Public Sub CollectTableColumnPairs( _
    ByVal sourceTable As ListObject, _
    ByVal keyColumnName As String, _
    ByVal itemColumnName As String, _
    ByVal pairsDict As Scripting.Dictionary)
    
    On Error GoTo ErrHandler
    
    Dim keyColumn As Range
    Dim itemColumn As Range
    Dim arrKeys As Variant
    Dim arrItems As Variant
    Dim i As Long
    Dim maxCount As Long
    Dim keyVal As Variant
    Dim itemVal As Variant
    Dim keyText As String
    
    ' Validate input
    If sourceTable Is Nothing Then GoTo CleanExit
    If Len(keyColumnName) = 0 Or Len(itemColumnName) = 0 Then GoTo CleanExit
    
    ' Resolve table columns
    On Error Resume Next
    Set keyColumn = sourceTable.ListColumns(keyColumnName).DataBodyRange
    Set itemColumn = sourceTable.ListColumns(itemColumnName).DataBodyRange
    On Error GoTo ErrHandler
    
    If keyColumn Is Nothing Or itemColumn Is Nothing Then GoTo CleanExit
    
    ' Check if table has data rows
    If sourceTable.ListRows.Count = 0 Then GoTo CleanExit
    
    ' Determine the minimum count to process
    maxCount = Application.WorksheetFunction.Min(keyColumn.Rows.Count, itemColumn.Rows.Count)
    If maxCount = 0 Then GoTo CleanExit
    
    ' Load values into arrays for performance
    ' Single cell ranges need special handling
    If keyColumn.Cells.Count = 1 Then
        ReDim arrKeys(1 To 1, 1 To 1)
        arrKeys(1, 1) = keyColumn.Value
    Else
        arrKeys = keyColumn.Value
    End If
    
    If itemColumn.Cells.Count = 1 Then
        ReDim arrItems(1 To 1, 1 To 1)
        arrItems(1, 1) = itemColumn.Value
    Else
        arrItems = itemColumn.Value
    End If
    
    ' Process pairs
    For i = 1 To maxCount
        ' Extract key value (always from first column of the range)
        If IsArray(arrKeys) Then
            keyVal = arrKeys(i, 1)
        Else
            keyVal = arrKeys
        End If
        
        ' Extract item value (always from first column of the range)
        If IsArray(arrItems) Then
            itemVal = arrItems(i, 1)
        Else
            itemVal = arrItems
        End If
        
        ' Add or update dictionary entry (skip empty keys)
        If Not IsEmpty(keyVal) Then
            keyText = Trim$(CStr(keyVal))
            If Len(keyText) > 0 Then
                pairsDict(keyText) = itemVal
            End If
        End If
    Next i
    
CleanExit:
    Exit Sub
ErrHandler:
    modErr.ReportError "CollectTableColumnPairs", Err.Number, Erl, caption:=modMain.AppProjectName
    Resume CleanExit
End Sub


' -----------------------------------------------------------------------------------
' Procedure : UpdateNamedListRange
' Purpose   : Updates or creates a named range in the workbook referring to a
'             column of values starting from row 2 to the last used row.
'
' Parameters:
'   rangeName  [String]     - The name of the named range to update/create
'   sheet      [Worksheet]  - The worksheet containing the range
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
'   sheet                   [Worksheet] - The worksheet to write into
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
' Purpose   : Writes the contents of a dictionary into structured table columns,
'             sorted alphabetically by keys, and expands the table if needed.
'
' Parameters:
'   targetTable        [ListObject]            - The structured table to write into
'   keyColumnName      [String]                - Name of the column for dictionary keys
'   valuesDict         [Scripting.Dictionary]  - Dictionary containing the keys/values
'   itemColumnName     [String]                - (Optional) Name of the column for dictionary items.
'                                                 If omitted, only keys are written.
'
' Notes:
'   - Automatically adds rows to the table if valuesDict.Count > current row count
'   - Keys are sorted alphabetically before writing
'   - If itemColumnName is provided, writes key-value pairs to both columns
' -----------------------------------------------------------------------------------
Public Sub WriteDictSetToTableColumn(ByVal targetTable As ListObject, _
                                     ByVal keyColumnName As String, _
                                     ByVal valuesDict As Scripting.Dictionary, _
                                     Optional ByVal itemColumnName As String = "")
    On Error GoTo ErrHandler
    
    ' exit if dict is empty
    If valuesDict.Count = 0 Then
        Exit Sub
    End If
    
    Dim sortedKeys As Collection
    Dim keyArray() As String
    Dim currentKey As Variant
    Dim index As Long
    Dim keyColumn As Range
    Dim itemColumn As Range
    Dim writeItems As Boolean
    
    Set sortedKeys = New Collection
    
    ' Determine if we need to write items
    writeItems = (Len(itemColumnName) > 0)
    
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
    
    ' Get column references
    Set keyColumn = targetTable.ListColumns(keyColumnName).DataBodyRange
    If writeItems Then
        Set itemColumn = targetTable.ListColumns(itemColumnName).DataBodyRange
    End If
    
    ' Write values into table columns
    For index = 1 To sortedKeys.Count
        keyColumn.Cells(index, 1).Value = sortedKeys(index)
        If writeItems Then
            itemColumn.Cells(index, 1).Value = valuesDict(sortedKeys(index))
        End If
    Next index
    
CleanExit:
    Exit Sub
ErrHandler:
    modErr.ReportError "WriteDictSetToTableColumn", Err.Number, Erl, caption:=modMain.AppProjectName
    Resume CleanExit
End Sub


' -----------------------------------------------------------------------------------
' Procedure : AppendMissingDictSetToColumns
' Purpose   : Appends key-value pairs from newDictSet to columns if the keys do
'             not already exist in the supplied existing dictionary.
'
' Parameters:
'   sheet            [Worksheet]             - Target worksheet.
'   keyColumnIdx     [Long]                  - Target column number for keys.
'   existingDictSet  [Scripting.Dictionary]  - Dictionary containing existing keys/values.
'   newDictSet       [Scripting.Dictionary]  - Dictionary containing new keys/value to append.
'   startRow         [Long]                  - Optional, First data row (default 2).
'   itemColumnIdx    [Long]                  - Optional, Target column number for items.
'                                              If 0 or omitted, only keys are written.
'
' Notes:
'   - If itemColumnIdx is provided, both keys and items are written
'   - existingDictSet is updated with newly added keys
' -----------------------------------------------------------------------------------
Public Sub AppendMissingDictSetToColumns( _
    ByVal sheet As Worksheet, _
    ByVal keyColumnIdx As Long, _
    ByVal existingDictSet As Scripting.Dictionary, _
    ByVal newDictSet As Scripting.Dictionary, _
    Optional ByVal startRow As Long = 2, _
    Optional ByVal itemColumnIdx As Long = 0)
    On Error GoTo ErrHandler
    
    Dim lastRow As Long
    Dim nextRow As Long
    Dim currentKey As Variant
    Dim keyText As String
    Dim itemValue As Variant
    Dim writeItems As Boolean
    
    ' Determine if we need to write items
    writeItems = (itemColumnIdx > 0)
    
    ' Find last row in key column
    lastRow = sheet.Cells(sheet.Rows.Count, keyColumnIdx).End(xlUp).Row
    If lastRow < startRow Then
        nextRow = startRow
    Else
        nextRow = lastRow + 1
    End If
    
    ' Iterate through new pairs
    For Each currentKey In newDictSet.Keys
        keyText = Trim$(CStr(currentKey))
        
        If Len(keyText) > 0 Then
            If Not existingDictSet.Exists(keyText) Then
                ' Write key
                sheet.Cells(nextRow, keyColumnIdx).Value = keyText
                
                ' Write item if requested
                If writeItems Then
                    itemValue = newDictSet(currentKey)
                    sheet.Cells(nextRow, itemColumnIdx).Value = itemValue
                End If
                
                ' Update existing keys dictionary
                existingDictSet(keyText) = True
                
                nextRow = nextRow + 1
            End If
        End If
    Next currentKey
    
CleanExit:
    Exit Sub
ErrHandler:
    modErr.ReportError "AppendMissingDictSetToColumns", Err.Number, Erl, caption:=modMain.AppProjectName
    Resume CleanExit
End Sub

' -----------------------------------------------------------------------------------
' Procedure : AppendMissingDictSetToTableColumns
' Purpose   : Appends key-value pairs from newDictSet to structured table columns
'             if the keys do not already exist in existingDictSet. Expands the table
'             as needed.
'
' Parameters:
'   targetTable      [ListObject]               - The structured table to write into.
'   keyColumnName    [String]                   - The name of the column for keys.
'   existingDictSet  [Scripting.Dictionary]     - Dictionary containing existing keys/values
'   newDictSet       [Scripting.Dictionary]     - Dictionary containing new keys/values to append.
'   itemColumnName   [String]                   - (Optional) The name of the column for items.
'                                                  If omitted, only keys are written.
'
' Notes:
'   - If itemColumnName is provided, both keys and items are written
'   - existingDictSet is updated with newly added keys
' -----------------------------------------------------------------------------------
Public Sub AppendMissingDictSetToTableColumns( _
    ByVal targetTable As ListObject, _
    ByVal keyColumnName As String, _
    ByVal existingDictSet As Scripting.Dictionary, _
    ByVal newDictSet As Scripting.Dictionary, _
    Optional ByVal itemColumnName As String = "")
    On Error GoTo ErrHandler
    
    Dim currentKey As Variant
    Dim keyText As String
    Dim itemValue As Variant
    Dim keyColumn As Range
    Dim itemColumn As Range
    Dim nextRowIndex As Long
    Dim writeItems As Boolean
    Dim keyColIndex As Long
    Dim itemColIndex As Long
    
    ' Determine if we need to write items
    writeItems = (Len(itemColumnName) > 0)
    
    ' Get column references and indices
    Set keyColumn = targetTable.ListColumns(keyColumnName).DataBodyRange
    keyColIndex = targetTable.ListColumns(keyColumnName).index
    
    If writeItems Then
        Set itemColumn = targetTable.ListColumns(itemColumnName).DataBodyRange
        itemColIndex = targetTable.ListColumns(itemColumnName).index
    End If
    
    ' Start writing after last row
    nextRowIndex = keyColumn.Rows.Count + 1
    
    ' Iterate through new pairs
    For Each currentKey In newDictSet.Keys
        keyText = Trim$(CStr(currentKey))
        
        If Len(keyText) > 0 Then
            If Not existingDictSet.Exists(keyText) Then
                ' Add new row to table
                targetTable.ListRows.Add
                
                ' Write key
                targetTable.DataBodyRange.Cells(nextRowIndex, keyColIndex).Value = keyText
                
                ' Write item if requested
                If writeItems Then
                    itemValue = newDictSet(currentKey)
                    targetTable.DataBodyRange.Cells(nextRowIndex, itemColIndex).Value = itemValue
                End If
                
                ' Update existing keys dictionary
                existingDictSet(keyText) = True
                
                nextRowIndex = nextRowIndex + 1
            End If
        End If
    Next currentKey
    
CleanExit:
    Exit Sub
ErrHandler:
    modErr.ReportError "AppendMissingDictSetToTableColumns", Err.Number, Erl, caption:=modMain.AppProjectName
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
