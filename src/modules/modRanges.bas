Attribute VB_Name = "modRanges"
' -----------------------------------------------------------------------------------
' Module    : modRanges
' Purpose   : Helper utilities for detecting framed ranges, column ranges by headers,
'             and checking ranges for specific formulas/values.
'
' Public API:
'   - FindFramedRangeByHeading : Returns a framed rectangular range around a heading.
'   - GetColumnRangeByHeader   : Returns the vertical range from a header down to a strong bottom border.
'   - RangeHasFormula2         : Checks if any cell's Formula2 matches the given text.
'   - RangeHasValue            : Checks if any cell equals/contains a given value.
'   - GetTableColumnRange      : Returns the DataBodyRange for a specific ListObject column.
'   - GetCellNameByPattern     : Returns the Name object of a cell if it has a defined name
'   - GetTable                 : Returns the ListObject with name and worksheet
'
' Dependencies:
'   - modErr : error reporting
' Notes     :
' -----------------------------------------------------------------------------------
Option Explicit
Option Private Module

' ===== Public API ===================================================================

' -----------------------------------------------------------------------------------
' Function  : FindFramedRangeByHeading
' Purpose   : Finds the rectangular, border-framed range that contains the specified
'             heading text and returns that range.
'
' Parameters:
'   targetSheet     [Worksheet]        - Worksheet to search in.
'   headingText     [String]           - Heading text to locate.
'   isWholeMatch    [Boolean]          - Optional. If True, matches the whole cell (xlWhole);
'                                        if False, allows partial match (xlPart). Default True.
'   excludeHeader   [Boolean]          - Optional. If True, excludes the header cell from the
'                                        returned range. Default True.
'
' Returns   : [Range] - The framed range bounded by left/right/top/bottom borders that
'                       encloses the heading; Nothing if not found or if the frame is incomplete.
'
' Notes     :
'   - Uses Cells.Find (xlValues, case-insensitive) to locate the heading.
'   - If the heading cell is merged, uses the first cell of the MergeArea.
'   - Scans horizontally to detect left/right borders, scans upward to find the top
'     border, and scans downward until both side cells have bottom borders.
' -----------------------------------------------------------------------------------
Public Function FindFramedRangeByHeading( _
    ByVal targetSheet As Worksheet, _
    ByVal headingText As String, _
    Optional ByVal isWholeMatch As Boolean = True, _
    Optional ByVal excludeHeader As Boolean = True) As Range

    Dim foundCell As Range
    Dim topRow As Long, bottomRow As Long
    Dim leftCol As Long, rightCol As Long
    Dim scanRow As Long, scanCol As Long
    Dim lookAtMode As XlLookAt
    Dim resultRng As Range
    
    On Error GoTo ErrHandler

    lookAtMode = IIf(isWholeMatch, xlWhole, xlPart)

    Set foundCell = targetSheet.Cells.Find(What:=headingText, LookIn:=xlValues, LookAt:=lookAtMode, _
                                           MatchCase:=False, SearchOrder:=xlByRows, SearchDirection:=xlNext)
    If foundCell Is Nothing Then Exit Function

    ' Work with the first cell of a merged heading area
    If foundCell.MergeCells Then Set foundCell = foundCell.MergeArea.Cells(1, 1)

    ' Find left edge: move left until a left border exists
    scanCol = foundCell.column
    Do While scanCol > 1 And Not CellHasBorder(targetSheet.Cells(foundCell.Row, scanCol), xlEdgeLeft)
        scanCol = scanCol - 1
    Loop
    If Not CellHasBorder(targetSheet.Cells(foundCell.Row, scanCol), xlEdgeLeft) Then Exit Function
    leftCol = scanCol

    ' Find right edge: move right until a right border exists
    scanCol = foundCell.column
    Do While scanCol < targetSheet.Columns.Count And Not CellHasBorder(targetSheet.Cells(foundCell.Row, scanCol), xlEdgeRight)
        scanCol = scanCol + 1
    Loop
    If Not CellHasBorder(targetSheet.Cells(foundCell.Row, scanCol), xlEdgeRight) Then Exit Function
    rightCol = scanCol

    ' Find top edge: move up until a top border exists
    scanRow = foundCell.Row
    Do While scanRow > 1 And Not CellHasBorder(targetSheet.Cells(scanRow, foundCell.column), xlEdgeTop)
        scanRow = scanRow - 1
    Loop
    If Not CellHasBorder(targetSheet.Cells(scanRow, foundCell.column), xlEdgeTop) Then Exit Function
    topRow = scanRow

    ' Find bottom edge: scan down until BOTH left & right edge cells have bottom borders
    scanRow = foundCell.Row
    Do While scanRow < targetSheet.Rows.Count
        If CellHasBorder(targetSheet.Cells(scanRow, leftCol), xlEdgeBottom) And _
           CellHasBorder(targetSheet.Cells(scanRow, rightCol), xlEdgeBottom) Then
            bottomRow = scanRow
            Exit Do
        End If
        scanRow = scanRow + 1
    Loop
    If bottomRow = 0 Then Exit Function

    Set resultRng = targetSheet.Range( _
        targetSheet.Cells(topRow + CInt(excludeHeader), leftCol), _
        targetSheet.Cells(bottomRow, rightCol))
        
CleanExit:
    Set FindFramedRangeByHeading = resultRng
    Exit Function

ErrHandler:
    modErr.ReportError "FindFramedRangeByHeading", Err.Number, Erl, caption:=modMain.AppProjectName
    Resume CleanExit
End Function

' -----------------------------------------------------------------------------------
' Function  : GetColumnRangeByHeader
' Purpose   : Finds the column range starting at a header cell and ending at the
'             first cell whose bottom border matches the expected weight.
'
' Parameters:
'   targetSheet     [Worksheet]       - Worksheet to search in.
'   headerText      [String]          - Header text to locate.
'   isWholeMatch    [Boolean]         - Optional. If True, xlWhole; if False, xlPart. Default True.
'   borderWeight    [XlBorderWeight]  - Optional. Expected border weight. Default xlMedium.
'   excludeHeader   [Boolean]         - Optional. If True, excludes the header cell. Default True.
'
' Returns   : [Range] - From the header cell down to the detected bottom edge; Nothing if not found.
'
' Notes     :
'   - Uses Cells.Find with case-insensitive search; respects merged header cells.
'   - Detects the bottom edge via CellHasBorder(..., xlEdgeBottom, ...).
' -----------------------------------------------------------------------------------
Public Function GetColumnRangeByHeader( _
    ByVal targetSheet As Worksheet, _
    ByVal headerText As String, _
    Optional ByVal isWholeMatch As Boolean = True, _
    Optional ByVal borderWeight As XlBorderWeight = xlMedium, _
    Optional ByVal excludeHeader As Boolean = True) As Range

    Dim headerCell As Range
    Dim headerCol As Long, rowIdx As Long, lastRow As Long
    Dim lookAtMode As XlLookAt: lookAtMode = IIf(isWholeMatch, xlWhole, xlPart)
    Dim resultRng As Range
    
    On Error GoTo ErrHandler

    ' Locate the header cell
    Set headerCell = targetSheet.Cells.Find(What:=headerText, LookIn:=xlValues, LookAt:=lookAtMode, _
                                            MatchCase:=False, SearchOrder:=xlByRows, SearchDirection:=xlNext)
    If headerCell Is Nothing Then Exit Function
    If headerCell.MergeCells Then Set headerCell = headerCell.MergeArea.Cells(1, 1)

    headerCol = headerCell.column

    ' Scan down until a matching bottom border is found
    For rowIdx = headerCell.Row To targetSheet.Rows.Count
        If CellHasBorder(targetSheet.Cells(rowIdx, headerCol), xlEdgeBottom, borderWeight) Then
            lastRow = rowIdx
            Exit For
        End If
    Next rowIdx
    If lastRow = 0 Then Exit Function

    Set resultRng = targetSheet.Range( _
        targetSheet.Cells(headerCell.Row + CInt(excludeHeader), headerCol), _
        targetSheet.Cells(lastRow, headerCol))
        
CleanExit:
    Set GetColumnRangeByHeader = resultRng
    Exit Function

ErrHandler:
    modErr.ReportError "GetColumnRangeByHeader", Err.Number, Erl, caption:=modMain.AppProjectName
    Resume CleanExit
End Function

' -----------------------------------------------------------------------------------
' Function  : RangeHasFormula2
' Purpose   : Checks whether any cell in the given range has a Formula2 equal to the
'             provided text (without the leading "="), with optional exact/substring
'             and case-sensitive/insensitive comparison.
'
' Parameters:
'   searchRange     [Range]   - Range to search within.
'   formulaTextNoEq [String]  - Formula text without the leading "=".
'   isExactMatch    [Boolean] - Optional. If True, exact match; if False, substring match. Default True.
'   isCaseSensitive [Boolean] - Optional. If True, case-sensitive; otherwise case-insensitive. Default False.
'
' Returns   : [Boolean] - True if a matching Formula2 is found; otherwise False.
'
' Notes     :
'   - Limits scan to formula cells using SpecialCells(xlCellTypeFormulas); returns False if none.
'   - Normalizes both sides via NormalizeFormula2Text (trims, removes "=", unifies separators, strips spaces/CR/LF).
' -----------------------------------------------------------------------------------
Public Function RangeHasFormula2( _
    ByVal searchRange As Range, _
    ByVal formulaTextNoEq As String, _
    Optional ByVal isExactMatch As Boolean = True, _
    Optional ByVal isCaseSensitive As Boolean = False) As Boolean

    Dim formulaCells As Range
    Dim cellItem As Range
    Dim targetText As String, cellText As String
    Dim hasValue As Boolean
    
    On Error GoTo ErrHandler

    If searchRange Is Nothing Then Exit Function

    ' Limit to formula cells for speed; handle "no formulas" case.
    On Error Resume Next
    Set formulaCells = searchRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    If formulaCells Is Nothing Then Exit Function

    targetText = NormalizeFormula2Text(formulaTextNoEq, isCaseSensitive)

    For Each cellItem In formulaCells.Cells
        cellText = NormalizeFormula2Text(cellItem.Formula2, isCaseSensitive)

        If isExactMatch Then
            If cellText = targetText Then
                hasValue = True: GoTo CleanExit
            End If
        Else
            ' Substring match if needed
            If InStr(1, cellText, targetText, IIf(isCaseSensitive, vbBinaryCompare, vbTextCompare)) > 0 Then
                hasValue = True: GoTo CleanExit
            End If
        End If
    Next cellItem
    
CleanExit:
    RangeHasFormula2 = hasValue
    Exit Function

ErrHandler:
    modErr.ReportError "RangeHasFormula2", Err.Number, Erl, caption:=modMain.AppProjectName
    Resume CleanExit
End Function

' -----------------------------------------------------------------------------------
' Function  : RangeHasValue
' Purpose   : Checks whether any cell in the given range contains the specified value.
'             Text search can be exact or substring match; numeric and date checks are exact.
'
' Parameters:
'   searchRange     [Range]   - Range to search within.
'   targetValue     [Variant] - Value to search for.
'   isExactMatch    [Boolean] - Optional. If True, exact match; if False, substring match. Default True.
'   isCaseSensitive [Boolean] - Optional. If True, case-sensitive; otherwise case-insensitive. Default False.
'
' Returns   : [Boolean] - True if any cell in searchRange contains the given value.
'
' Notes     :
'   - Uses Value2 for array extraction; handles 1x1 and multi-cell ranges.
' -----------------------------------------------------------------------------------
Public Function RangeHasValue( _
    ByVal searchRange As Range, _
    ByVal targetValue As Variant, _
    Optional ByVal isExactMatch As Boolean = True, _
    Optional ByVal isCaseSensitive As Boolean = False) As Boolean

    Dim hasValue As Boolean
    
    On Error GoTo ErrHandler

    If searchRange Is Nothing Then Exit Function

    Dim dataValues As Variant
    dataValues = searchRange.value2  ' Could be scalar for 1x1

    Dim cmpMode As VbCompareMethod
    cmpMode = IIf(isCaseSensitive, vbBinaryCompare, vbTextCompare)

    Dim rowIdx As Long, colIdx As Long
    Dim rowLBound As Long, rowUBound As Long, colLBound As Long, colUBound As Long

    If IsArray(dataValues) Then
        rowLBound = LBound(dataValues, 1): rowUBound = UBound(dataValues, 1)
        colLBound = LBound(dataValues, 2): colUBound = UBound(dataValues, 2)

        For rowIdx = rowLBound To rowUBound
            For colIdx = colLBound To colUBound
                If CellMatches(dataValues(rowIdx, colIdx), targetValue, isExactMatch, cmpMode) Then
                    hasValue = True: GoTo CleanExit
                End If
            Next colIdx
        Next rowIdx
    Else
        ' 1x1 range: dataValues is scalar
        hasValue = CellMatches(dataValues, targetValue, isExactMatch, cmpMode)
    End If

CleanExit:
    RangeHasValue = hasValue
    Exit Function
    
ErrHandler:
    modErr.ReportError "RangeHasValue", Err.Number, Erl, caption:=modMain.AppProjectName
    Resume CleanExit
End Function

' -----------------------------------------------------------------------------------
' Function  : GetTableColumnRange
' Purpose   : Returns the DataBodyRange for a given ListObject column by name.
'
' Parameters:
'   targetSheet     [Worksheet] - Worksheet that hosts the ListObject.
'   listObjectName  [String]    - Name of the table (ListObject).
'   columnName      [String]    - Display name of the table column.
'
' Returns   : [Range] - The DataBodyRange for the column; Nothing on error or if not found.
'
' Notes     :
'   - Errors are ignored to allow a Nothing return when objects are missing.
' -----------------------------------------------------------------------------------
Public Function GetTableColumnRange( _
    ByVal targetSheet As Worksheet, _
    ByVal listObjectName As String, _
    ByVal columnName As String) As Range

    On Error Resume Next
    Set GetTableColumnRange = targetSheet.ListObjects(listObjectName).ListColumns(columnName).DataBodyRange
    On Error GoTo 0
End Function

' -----------------------------------------------------------------------------------
' Function  : GetCellNameByPattern
' Purpose   : Returns the Name object of a cell if it has a defined name whose
'             Name property matches the given pattern.
'
' Parameters:
'   cell        [Range]  - Target cell to inspect.
'   pattern     [String] - Pattern to match against Name.Name (e.g. "DropDownCell*").
'
' Returns   : [Name] - The matching Name object if found; Nothing otherwise.
'
' Notes     :
'   - Expects a single-cell range; multi-cell ranges are ignored.
'   - Uses VBA Like operator for pattern matching.
'   - Safely ignores errors if the cell has no associated Name.
' -----------------------------------------------------------------------------------
Public Function GetCellNameByPattern( _
    ByVal cell As Range, _
    ByVal pattern As String) As Name

    Dim cellName As Name
    
    If cell Is Nothing Then Exit Function
    If cell.CountLarge <> 1 Then Exit Function

    On Error Resume Next
    If cell.Name.Name Like pattern Then
        Set cellName = cell.Name
    End If
    On Error GoTo 0

    Set GetCellNameByPattern = cellName
End Function

' ===== Private Helpers =============================================================

' -----------------------------------------------------------------------------------
' Function  : NormalizeFormula2Text
' Purpose   : Normalizes a Formula2 string for comparison by removing cosmetic
'             differences and harmonizing list separators.
'
' Parameters:
'   formulaText     [String]   - Formula text (with or without leading "=").
'   isCaseSensitive [Boolean]  - If False, folds to lower-case; if True, preserves case.
'
' Returns   : [String] - Normalized formula text (no leading "=", unified separators,
'                        no spaces/CR/LF).
'
' Notes     :
'   - Uses Application.International(xlListSeparator) to harmonize "," and ";".
'   - Trims whitespace, removes leading "=", removes spaces and CR/LF characters.
' -----------------------------------------------------------------------------------
Private Function NormalizeFormula2Text( _
    ByVal formulaText As String, _
    ByVal isCaseSensitive As Boolean) As String

    Dim listSep As String
    listSep = Application.International(xlListSeparator)

    formulaText = Trim$(formulaText)
    If Left$(formulaText, 1) = "=" Then formulaText = Mid$(formulaText, 2)

    ' Unify list separators; user may type "," while UI uses ";", or vice versa
    formulaText = Replace(formulaText, ",", listSep)
    formulaText = Replace(formulaText, ";", listSep)

    ' Remove spaces and CR/LF to avoid cosmetic mismatches
    formulaText = Replace(formulaText, " ", vbNullString)
    formulaText = Replace(formulaText, vbCr, vbNullString)
    formulaText = Replace(formulaText, vbLf, vbNullString)

    If Not isCaseSensitive Then formulaText = LCase$(formulaText)

    NormalizeFormula2Text = formulaText
End Function


' -----------------------------------------------------------------------------------
' Function  : CellHasBorder
' Purpose   : Checks whether a specific edge of a cell has a continuous border
'             with the specified weight.
'
' Parameters:
'   targetCell   [Range]           - Target cell to inspect.
'   edgeIndex    [XlBordersIndex]  - Edge to check (xlEdgeLeft, xlEdgeRight, xlEdgeTop, xlEdgeBottom).
'   borderWeight [XlBorderWeight]  - Optional. Expected border weight; default xlMedium.
'
' Returns   : [Boolean] - True if the edge uses LineStyle xlContinuous and matches the given weight.
'
' Notes     :
'   - Intended for framed-range detection used by FindFramedRangeByHeading.
' -----------------------------------------------------------------------------------
Private Function CellHasBorder( _
    ByVal targetCell As Range, _
    ByVal edgeIndex As XlBordersIndex, _
    Optional ByVal borderWeight As XlBorderWeight = xlMedium) As Boolean

    With targetCell.Borders(edgeIndex)
        CellHasBorder = (.LineStyle = xlContinuous And .Weight = borderWeight)
    End With
End Function


' -----------------------------------------------------------------------------------
' Function  : CellMatches
' Purpose   : Compares a single cell value against a target using exact or substring
'             semantics for text; exact comparison for numbers and dates.
'
' Parameters:
'   cellValue     [Variant]          - Value from a worksheet cell.
'   targetValue   [Variant]          - Expected value to compare against.
'   isExactMatch  [Boolean]          - If True, exact match; if False, substring match for text.
'   cmpMode       [VbCompareMethod]  - Comparison mode for text checks.
'
' Returns   : [Boolean] - True if the cell value matches according to the given rules.
' -----------------------------------------------------------------------------------
Private Function CellMatches( _
    ByVal cellValue As Variant, _
    ByVal targetValue As Variant, _
    ByVal isExactMatch As Boolean, _
    ByVal cmpMode As VbCompareMethod) As Boolean

    If IsEmpty(cellValue) Then Exit Function

    Select Case VarType(targetValue)
        Case vbString
            If VarType(cellValue) = vbString Then
                If isExactMatch Then
                    CellMatches = (StrComp(cellValue, targetValue, cmpMode) = 0)
                Else
                    CellMatches = (InStr(1, cellValue, targetValue, cmpMode) > 0)
                End If
            End If

        Case vbDouble, vbSingle, vbCurrency, vbInteger, vbLong, vbDecimal
            If IsNumeric(cellValue) Then
                CellMatches = (CDbl(cellValue) = CDbl(targetValue))
            End If

        Case vbDate
            If IsDate(cellValue) Then
                ' Compare Excel serials (Value2 is serial for dates)
                CellMatches = (CLng(CDbl(cellValue)) = CLng(CDbl(targetValue)))
            End If

        Case Else
            ' Fallback as string compare
            If isExactMatch Then
                CellMatches = (StrComp(CStr(cellValue), CStr(targetValue), cmpMode) = 0)
            Else
                CellMatches = (InStr(1, CStr(cellValue), CStr(targetValue), cmpMode) > 0)
            End If
    End Select
End Function

' -----------------------------------------------------------------------------------
' Function  : GetTable
' Purpose   : Returns the ListObject with the given name from the specified worksheet.
'
' Parameters:
'   ws          [Worksheet] - Worksheet that hosts the table.
'   tableName   [String]    - Name of the ListObject to return.
'
' Returns   : [ListObject] - The matching ListObject; Nothing if not found.
'
' Notes     :
'   - Uses On Error Resume Next to allow a Nothing return when the table
'     does not exist on the given worksheet.
' -----------------------------------------------------------------------------------
Public Function GetTable( _
    ByVal ws As Worksheet, _
    ByVal tableName As String) As ListObject

    On Error Resume Next
    Set GetTable = ws.ListObjects(tableName)
    On Error GoTo 0
End Function


