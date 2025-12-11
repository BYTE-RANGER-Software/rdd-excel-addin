Attribute VB_Name = "modLists"
' -----------------------------------------------------------------------------------
' Module: modLists
' Purpose: Helper functions for collecting, normalizing, and writing list-like
'          data from worksheets, named ranges, and structured tables into
'          dictionaries and ranges. Intended mainly for validation lists and
'          configuration-style data.
'
' Public API:
'   - AppendMissingDictSetToTableColumns
'   - GetNamedOrHeaderValue
'   - CollectNamedRangePairsToDict
'   - CollectNamedRangeValuesToDict
'   - CollectTableColumnsToDict
'   - WriteDictToTableColumns
'   - ClearTableColumn
'
' Notes:
'   - Option Private Module keeps the module internal to the project.
' -----------------------------------------------------------------------------------
Option Explicit
Option Private Module

Public Const DICT_VALUE_SEPERATOR  As String = "|"

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
'   headerArray    [Variant]   - (Optional) Array of header names to search for (case-insensitive)
'
' Returns   : The trimmed string value from the named range or the cell next to
'             the first matching header found. Returns an empty string if no match.
'
' Behavior  :
'   - Tries to read from the named range first
'   - If empty and headerArray is present, scans the Used Range for a header match
'   - If a match is found, returns the value in the cell to the right
'
' Notes     :
'   - Header comparison is case-insensitive using StrComp
'   - Assumes headers are in row/column format and values are to the right
'   - No error raised if named range is missing or header not found
' -----------------------------------------------------------------------------------
Public Function GetNamedOrHeaderValue(ByVal sheet As Worksheet, ByVal localName As String, Optional ByVal headerArray As Variant) As String
    On Error GoTo ErrHandler

    Dim namedValue As Variant
    Dim searchRange As Range
    Dim headerCell As Range
    Dim bestHeaderCell As Range
    Dim headerIndex As Long
    Dim firstAddress As String

    ' Try named range first (silent failure if missing)
    On Error Resume Next
    namedValue = sheet.Range(localName).value
    On Error GoTo ErrHandler

    If Len(Trim$(CStr(namedValue))) > 0 Then
        GetNamedOrHeaderValue = Trim$(CStr(namedValue))
        GoTo CleanExit
    End If

    ' Search for headers in A1:N115
    Dim lastCell As Range
    Set lastCell = sheet.Cells.SpecialCells(xlCellTypeLastCell)
    Set searchRange = sheet.Range(sheet.Cells(1, 1), lastCell)
    
    If IsMissing(headerArray) Then GoTo CleanExit
    
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
        GetNamedOrHeaderValue = Trim$(CStr(bestHeaderCell.offset(0, 1).value))
    End If

CleanExit:
    Exit Function
ErrHandler:
    modErr.ReportError "GetNamedOrHeaderValue", Err.Number, Erl, caption:=modMain.AppProjectName
End Function

' -----------------------------------------------------------------------------------
' Procedure : CollectTableColumnsToDict
' Purpose   : Reads values from structured table columns into a dictionary,
'             with the key column and optional value columns.
'
' Parameters:
'   sourceTable      [ListObject]            - The structured table to read from
'   keyColumnName    [String]                - Name of the column containing dictionary keys
'   valuesDict       [Scripting.Dictionary]  - Dictionary to store the key-value pairs
'   value1ColumnName [String]                - (Optional) Name of the column for dictionary value 1.
'                                              If omitted, only keys are stored (with empty string values).
'   value2ColumnName [String]                - (Optional) Name of the column for dictionary value 2.
'                                              If provided, values are concatenated with DICT_VALUE_SEPERATOR delimiter.
'
' Behavior:
'   - Reads values from table columns as arrays (high performance).
'   - Ignores rows where the key is empty.
'   - If key already exists, updates the value.
'   - If value2ColumnName is provided, concatenates value1 and value2 with DICT_VALUE_SEPERATOR delimiter.
'   - If only value2ColumnName is provided (value1 empty), uses value2 as the sole value.
'
' Notes:
'   - If table is empty or columns don't exist, exits silently.
'   - Uses array access for optimal performance with large tables.
'   - Single-row tables are handled separately to avoid array issues.
' -----------------------------------------------------------------------------------
Public Sub CollectTableColumnsToDict(ByVal sourceTable As ListObject, _
                                        ByVal keyColumnName As String, _
                                        ByVal valuesDict As Scripting.Dictionary, _
                                        Optional ByVal value1ColumnName As String = "", _
                                        Optional ByVal value2ColumnName As String = "")
    On Error GoTo ErrHandler
    
    Dim arrKeys As Variant
    Dim arrValue1 As Variant
    Dim arrValue2 As Variant
    Dim i As Long
    Dim rowCount As Long
    Dim key As Variant
    Dim value As String
    Dim hasValue1 As Boolean
    Dim hasValue2 As Boolean
    
    ' Validate input
    If sourceTable Is Nothing Then GoTo CleanExit
    If Len(keyColumnName) = 0 Then GoTo CleanExit
    
    ' Check if table has data
    rowCount = sourceTable.ListRows.count
    If rowCount = 0 Then GoTo CleanExit
    
    ' Determine which value columns to read
    If Len(value1ColumnName) = 0 And Len(value2ColumnName) > 0 Then
        value1ColumnName = value2ColumnName
        value2ColumnName = ""
    End If
    
    hasValue1 = (Len(value1ColumnName) > 0)
    hasValue2 = (Len(value2ColumnName) > 0)
    
    ' Validate column existence
    On Error Resume Next
    Dim keyCol As ListColumn
    Dim val1Col As ListColumn
    Dim val2Col As ListColumn
    
    Set keyCol = sourceTable.ListColumns(keyColumnName)
    If keyCol Is Nothing Then GoTo CleanExit
    
    If hasValue1 Then
        Set val1Col = sourceTable.ListColumns(value1ColumnName)
        If val1Col Is Nothing Then GoTo CleanExit
    End If
    
    If hasValue2 Then
        Set val2Col = sourceTable.ListColumns(value2ColumnName)
        If val2Col Is Nothing Then GoTo CleanExit
    End If
    On Error GoTo ErrHandler
    
    ' Load values into arrays for performance
    If rowCount = 1 Then
        ' Single row - handle separately to avoid array issues
        ReDim arrKeys(1 To 1, 1 To 1)
        arrKeys(1, 1) = keyCol.DataBodyRange.value
        
        If hasValue1 Then
            ReDim arrValue1(1 To 1, 1 To 1)
            arrValue1(1, 1) = val1Col.DataBodyRange.value
        End If
        
        If hasValue2 Then
            ReDim arrValue2(1 To 1, 1 To 1)
            arrValue2(1, 1) = val2Col.DataBodyRange.value
        End If
    Else
        ' Multiple rows - read as arrays
        arrKeys = keyCol.DataBodyRange.value
        
        If hasValue1 Then
            arrValue1 = val1Col.DataBodyRange.value
        End If
        
        If hasValue2 Then
            arrValue2 = val2Col.DataBodyRange.value
        End If
    End If
    
    ' Process rows and populate dictionary
    For i = 1 To rowCount
        key = arrKeys(i, 1)
        
        ' Skip empty keys
        If Not IsEmpty(key) And Len(CStr(key)) > 0 Then
            value = ""
            
            If hasValue1 Then
                value = CStr(arrValue1(i, 1))
            End If
            
            If hasValue2 Then
                If Len(value) > 0 Then
                    value = value & DICT_VALUE_SEPERATOR & CStr(arrValue2(i, 1))
                Else
                    value = CStr(arrValue2(i, 1))
                End If
            End If
            
            ' Add or update dictionary entry
            valuesDict(CStr(key)) = value
        End If
    Next i
    
CleanExit:
    Exit Sub
ErrHandler:
    modErr.ReportError "CollectTableColumnsToDict", Err.Number, Erl, caption:=modMain.AppProjectName
End Sub

' -----------------------------------------------------------------------------------
' Procedure : CollectNamedRangePairsToDict
' Purpose   : Collects paired values from two named ranges on a worksheet and stores
'             them in a dictionary (first range = Key, second range = Value).
'
' Parameters:
'   sheet           [Worksheet] - The worksheet containing the named ranges.
'   rangeNameKeys   [String]    - The name of the range containing the keys.
'   rangeNameValues  [String]    - The name of the range containing the Values.
'   pairsDict       [Scripting.Dictionary] - Dictionary receiving the key-value pairs.
'
' Behavior  :
'   - Resolves both named ranges on the given worksheet.
'   - Reads values from both ranges as arrays (high performance).
'   - Ignores pairs where the key is empty.
'   - If key already exists, updates the Value.
'   - If ranges have different sizes, uses the smaller count.
'
' Notes     :
'   - If either named range is missing or invalid, the procedure exits silently.
'   - Uses array access for optimal performance with large ranges.
'   - Single-cell ranges are handled separately to avoid array issues.
' -----------------------------------------------------------------------------------
Public Sub CollectNamedRangePairsToDict(ByVal sheet As Worksheet, _
                                  ByVal rangeNameKeys As String, _
                                  ByVal rangeNameValues As String, _
                                  ByVal pairsDict As Scripting.Dictionary)
    On Error GoTo ErrHandler
    
    Dim rngKeys As Range
    Dim rngValues As Range
    Dim arrKeys As Variant
    Dim arrValues As Variant
    Dim i As Long
    Dim maxCount As Long
    Dim key As Variant
    Dim value As Variant
    
    ' Validate input
    If Len(rangeNameKeys) = 0 Or Len(rangeNameValues) = 0 Then GoTo CleanExit
    
    ' Resolve named ranges
    On Error Resume Next
    Set rngKeys = sheet.Range(rangeNameKeys)
    Set rngValues = sheet.Range(rangeNameValues)
    On Error GoTo ErrHandler
    
    If rngKeys Is Nothing Or rngValues Is Nothing Then GoTo CleanExit
    
    ' Determine the minimum count to process
    maxCount = Application.WorksheetFunction.Min(rngKeys.Cells.count, rngValues.Cells.count)
    If maxCount = 0 Then GoTo CleanExit
    
    ' Load values into arrays for performance
    ' Single cell ranges need special handling
    If rngKeys.Cells.count = 1 Then
        ReDim arrKeys(1 To 1, 1 To 1)
        arrKeys(1, 1) = rngKeys.value
    Else
        arrKeys = rngKeys.value
    End If
    
    If rngValues.Cells.count = 1 Then
        ReDim arrItems(1 To 1, 1 To 1)
        arrItems(1, 1) = rngValues.value
    Else
        arrItems = rngValues.value
    End If
    
    ' Process pairs
    For i = 1 To maxCount
        ' Handle both single-column and multi-column ranges
        If IsArray(arrKeys) Then
            If UBound(arrKeys, 2) > 1 Then
                ' Multi-column range - take first column
                key = arrKeys(i, 1)
            Else
                ' Single-column range
                key = arrKeys(i, 1)
            End If
        Else
            key = arrKeys
        End If
        
        If IsArray(arrItems) Then
            If UBound(arrItems, 2) > 1 Then
                value = arrItems(i, 1)
            Else
                value = arrItems(i, 1)
            End If
        Else
            value = arrItems
        End If
        
        ' Add or update dictionary entry (skip empty keys)
        If Not IsEmpty(key) And Len(CStr(key)) > 0 Then
            pairsDict(key) = value
        End If
    Next i
    
CleanExit:
    Exit Sub
ErrHandler:
    modErr.ReportError "CollectNamedRangePairsToDict", Err.Number, Erl, caption:=modMain.AppProjectName
End Sub

' -----------------------------------------------------------------------------------
' Procedure : CollectNamedRangeValuesToDict
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
Public Sub CollectNamedRangeValuesToDict(ByVal sheet As Worksheet, rangeName As String, ByVal valuesDict As Scripting.Dictionary)
    On Error GoTo ErrHandler

    Dim rngNamed As Range
    Dim cell As Range

    If Len(rangeName) = 0 Then GoTo CleanExit

    On Error Resume Next
    Set rngNamed = sheet.Range(rangeName)
    On Error GoTo ErrHandler

    If rngNamed Is Nothing Then GoTo CleanExit

    For Each cell In rngNamed.Cells
        If Not IsEmpty(cell.value) Then
            If Not valuesDict.Exists(cell.value) Then
                valuesDict.Add cell.value, True
            Else
                valuesDict(cell.value) = True
            End If
        End If
    Next cell

CleanExit:
    Exit Sub
ErrHandler:
    modErr.ReportError "CollectNamedRangeValuesToDict", Err.Number, Erl, caption:=modMain.AppProjectName
End Sub

' -----------------------------------------------------------------------------------
' Procedure : WriteDictToTableColumns
' Purpose   : Writes the contents of a dictionary into structured table columns,
'             sorted alphabetically by keys, and expands the table if needed.
'
' Parameters:
'   targetTable        [ListObject]            - The structured table to write into
'   keyColumnName      [String]                - Name of the column for dictionary keys
'   valuesDict         [Scripting.Dictionary]  - Dictionary containing the keys/values
'   value1ColumnName   [String]                - (Optional) Name of the column for dictionary value / value 1.
'                                                If omitted, only keys are written.
'   value2ColumnName   [String]                - (Optional) Name of the column for dictionary value 2.
'                                                If omitted, only keys and value1 are written.
'
' Notes:
'   - Automatically adds rows to the table if valuesDict.Count > current row count
'   - Keys are sorted alphabetically before writing
'   - If value1ColumnName is provided, writes key-value pairs to key and value1 columns
'   - If value2ColumnName is provided without value1ColumnName, value2ColumnName is used as value1ColumnName
'   - If dictionary values contain DICT_VALUE_SEPERATOR delimiter and both value columns are specified,
'     the value is split and written to both value1 and value2 columns
'   - Dictionary values without DICT_VALUE_SEPERATOR delimiter are written entirely to value1Column
' -----------------------------------------------------------------------------------
Public Sub WriteDictToTableColumns(ByVal targetTable As ListObject, _
                                     ByVal keyColumnName As String, _
                                     ByVal valuesDict As Scripting.Dictionary, _
                                     Optional ByVal value1ColumnName As String = "", _
                                     Optional ByVal value2ColumnName As String = "")
    On Error GoTo ErrHandler
    
    ' Exit if dict is empty
    If valuesDict.count = 0 Then Exit Sub
    
    Dim sortedKeys As Collection
    Dim keyArray() As String
    Dim currentKey As Variant
    Dim index As Long
    Dim writeValues As Long
    Dim values() As Variant
    Dim hasTwoColumns As Boolean
    
    ' Determine if we need to write values
    If Len(value1ColumnName) = 0 And Len(value2ColumnName) > 0 Then
        value1ColumnName = value2ColumnName
        value2ColumnName = ""
    End If
    
    writeValues = Abs((Len(value1ColumnName) > 0) + (Len(value2ColumnName) > 0))
    hasTwoColumns = (writeValues = 2)
    
    ' Copy keys to array and sort
    ReDim keyArray(0 To valuesDict.count - 1)
    index = 0
    For Each currentKey In valuesDict.Keys
        keyArray(index) = CStr(currentKey)
        index = index + 1
    Next currentKey
    
    modUtil.QuickSortStringArray keyArray, LBound(keyArray), UBound(keyArray)
    
    ' Load sorted keys into collection
    Set sortedKeys = New Collection
    For index = LBound(keyArray) To UBound(keyArray)
        sortedKeys.Add keyArray(index)
    Next index
    
    ' Ensure table has enough rows
    Do While targetTable.ListRows.count < sortedKeys.count
        targetTable.ListRows.Add
    Loop
    
    ' Prepare arrays for batch writing
    Dim keyData() As Variant
    Dim value1Data() As Variant
    Dim value2Data() As Variant
    Dim splitValues() As String
    Dim currentValue As Variant
    
    ReDim keyData(1 To sortedKeys.count, 1 To 1)
    If writeValues > 0 Then ReDim value1Data(1 To sortedKeys.count, 1 To 1)
    If hasTwoColumns Then ReDim value2Data(1 To sortedKeys.count, 1 To 1)
    
    ' Fill arrays
    For index = 1 To sortedKeys.count
        keyData(index, 1) = sortedKeys(index)
        
        If writeValues > 0 Then
            currentValue = valuesDict(sortedKeys(index))
            
            If hasTwoColumns And InStr(currentValue, DICT_VALUE_SEPERATOR) > 0 Then
                splitValues = Split(currentValue, DICT_VALUE_SEPERATOR, 2, vbTextCompare)
                value1Data(index, 1) = CStr(splitValues(0))
                value2Data(index, 1) = CStr(splitValues(1))
            Else
                value1Data(index, 1) = currentValue
                If hasTwoColumns Then value2Data(index, 1) = ""
            End If
        End If
    Next index
    
    ' Write arrays to table
    targetTable.ListColumns(keyColumnName).DataBodyRange.Resize(sortedKeys.count, 1).value = keyData
    
    If writeValues > 0 Then
        targetTable.ListColumns(value1ColumnName).DataBodyRange.Resize(sortedKeys.count, 1).value = value1Data
    End If
    
    If hasTwoColumns Then
        targetTable.ListColumns(value2ColumnName).DataBodyRange.Resize(sortedKeys.count, 1).value = value2Data
    End If
    
CleanExit:
    Exit Sub
ErrHandler:
    modErr.ReportError "WriteDictToTableColumns", Err.Number, Erl, caption:=modMain.AppProjectName
End Sub

' -----------------------------------------------------------------------------------
' Procedure : AppendMissingDictSetToTableColumns
' Purpose   : Appends key-value pairs from newDictSet to structured table columns
'             if the keys do not already exist in existingDictSet. Expands the table
'             as needed. Supports one or two value columns with DICT_VALUE_SEPERATOR.
'
' Parameters:
'   targetTable      [ListObject]               - The structured table to write into.
'   keyColumnName    [String]                   - The name of the column for keys.
'   existingDictSet  [Scripting.Dictionary]     - Dictionary containing existing keys/values
'   newDictSet       [Scripting.Dictionary]     - Dictionary containing new keys/values to append.
'   value1ColumnName [String]                   - (Optional) The name of the column for value 1.
'                                                  If omitted, only keys are written.
'   value2ColumnName [String]                   - (Optional) The name of the column for value 2.
'                                                  If omitted, only keys and value1 are written.
'
' Behavior:
'   - If value1ColumnName is provided, writes key-value pairs to key and value1 columns
'   - If value2ColumnName is provided without value1ColumnName, value2ColumnName is used as value1ColumnName
'   - If dictionary values contain DICT_VALUE_SEPERATOR delimiter and both value columns are specified,
'     the value is split and written to both value1 and value2 columns
'   - Dictionary values without DICT_VALUE_SEPERATOR delimiter are written entirely to value1Column
'   - existingDictSet is updated with newly added keys
'   - Only keys that do NOT exist in existingDictSet are appended
'
' Notes:
'   - Compatible with WriteDictToTableColumns and CollectTableColumnsToDict
'
' Example Usage:
'   ' Two-column append (Room ID => Room No | Room Alias)
'   modLists.AppendMissingDictSetToTableColumns dataList, LISTS_HEADER_ROOM_ID, _
'       existingKeysDict, newRoomsDict, LISTS_HEADER_ROOM_NO, LISTS_HEADER_ROOM_ALIAS
'
'   ' Single-column append (Item ID => Item Name)
'   modLists.AppendMissingDictSetToTableColumns dataList, LISTS_HEADER_ITEM_ID, _
'       existingKeysDict, newItemsDict, LISTS_HEADER_ITEM_NAME
'
'   ' Keys only
'   modLists.AppendMissingDictSetToTableColumns dataList, LISTS_HEADER_SCENE_ID, _
'       existingKeysDict, newScenesDict
' -----------------------------------------------------------------------------------
Public Sub AppendMissingDictSetToTableColumns( _
    ByVal targetTable As ListObject, _
    ByVal keyColumnName As String, _
    ByVal existingDictSet As Scripting.Dictionary, _
    ByVal newDictSet As Scripting.Dictionary, _
    Optional ByVal value1ColumnName As String = "", _
    Optional ByVal value2ColumnName As String = "")
    On Error GoTo ErrHandler
    
    Dim currentKey As Variant
    Dim keyText As String
    Dim currentValue As Variant
    Dim splitValues() As String
    Dim keyColumn As Range
    Dim value1Column As Range
    Dim value2Column As Range
    Dim nextRowIndex As Long
    Dim writeValues As Long
    Dim hasTwoColumns As Boolean
    Dim keyColIndex As Long
    Dim value1ColIndex As Long
    Dim value2ColIndex As Long
    
    ' Determine if we need to write values
    If Len(value1ColumnName) = 0 And Len(value2ColumnName) > 0 Then
        value1ColumnName = value2ColumnName
        value2ColumnName = ""
    End If
    
    writeValues = Abs((Len(value1ColumnName) > 0) + (Len(value2ColumnName) > 0))
    hasTwoColumns = (writeValues = 2)
    
    ' Get column references and indices
    Set keyColumn = targetTable.ListColumns(keyColumnName).DataBodyRange
    keyColIndex = targetTable.ListColumns(keyColumnName).index
    
    If writeValues > 0 Then
        Set value1Column = targetTable.ListColumns(value1ColumnName).DataBodyRange
        value1ColIndex = targetTable.ListColumns(value1ColumnName).index
    End If
    
    If hasTwoColumns Then
        Set value2Column = targetTable.ListColumns(value2ColumnName).DataBodyRange
        value2ColIndex = targetTable.ListColumns(value2ColumnName).index
    End If
    
    ' Start writing after last row
    nextRowIndex = keyColumn.Rows.count + 1
    
    ' Iterate through new pairs
    For Each currentKey In newDictSet.Keys
        keyText = Trim$(CStr(currentKey))
        
        If Len(keyText) > 0 Then
            If Not existingDictSet.Exists(keyText) Then
                ' Add new row to table
                targetTable.ListRows.Add
                
                ' Write key
                targetTable.DataBodyRange.Cells(nextRowIndex, keyColIndex).value = keyText
                
                ' Write values if requested
                If writeValues > 0 Then
                    currentValue = newDictSet(currentKey)
                    
                    If hasTwoColumns And InStr(currentValue, DICT_VALUE_SEPERATOR) > 0 Then
                        ' Split value into two columns (e.g., "RoomNo|RoomAlias")
                        splitValues = Split(currentValue, DICT_VALUE_SEPERATOR, 2, vbTextCompare)
                        targetTable.DataBodyRange.Cells(nextRowIndex, value1ColIndex).value = CStr(splitValues(0))
                        targetTable.DataBodyRange.Cells(nextRowIndex, value2ColIndex).value = CStr(splitValues(1))
                    Else
                        ' Write entire value to value1 column
                        targetTable.DataBodyRange.Cells(nextRowIndex, value1ColIndex).value = currentValue
                        If hasTwoColumns Then
                            targetTable.DataBodyRange.Cells(nextRowIndex, value2ColIndex).value = ""
                        End If
                    End If
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

' -----------------------------------------------------------------------------------
' Function: CollectTableColumnValuesToArray
' Purpose: Extracts the non-empty values of a ListObject column into an array.
'
' Parameters:
'   lo          [ListObject] - The ListObject
'   columnName  [String]     - The column header
'
' Returns: Variant array with the values (or empty array)
' -----------------------------------------------------------------------------------
Public Function CollectTableColumnValuesToArray(ByVal lo As ListObject, _
    ByVal columnName As String, Optional breakOnEmptyLine As Boolean = False) As Variant
    
    On Error Resume Next
    
    Dim col As ListColumn
    Dim cell As Range
    Dim result() As String
    Dim count As Long
    Dim cellValue As String
    
    ' Find column
    Set col = lo.ListColumns(columnName)
    If col Is Nothing Then
        CollectTableColumnValuesToArray = Array()
        Exit Function
    End If
    
    If col.DataBodyRange Is Nothing Then
        CollectTableColumnValuesToArray = Array()
        Exit Function
    End If
    
    ' Count non-empty values
    count = 0
    For Each cell In col.DataBodyRange
        cellValue = Trim$(CStr(cell.value))
        If LenB(cellValue) > 0 Then
            count = count + 1
        ElseIf breakOnEmptyLine Then
            Exit For
        End If
    Next cell
    
    If count = 0 Then
        CollectTableColumnValuesToArray = Array()
        Exit Function
    End If
    
    ' Fill array
    ReDim result(0 To count - 1)
    count = 0
    For Each cell In col.DataBodyRange
        cellValue = Trim$(CStr(cell.value))
        If LenB(cellValue) > 0 Then
            result(count) = cellValue
            count = count + 1
        ElseIf breakOnEmptyLine Then
            Exit For
        End If
    Next cell
    
    CollectTableColumnValuesToArray = result
    On Error GoTo 0
End Function


' ===== Private Helpers =============================================================



