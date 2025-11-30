Attribute VB_Name = "modRanges"
' -----------------------------------------------------------------------------------
' Module    : modRanges
' Purpose   : Helper utilities for detecting framed ranges, column ranges by headers,
'             and checking ranges for specific formulas/values.
'
' Public API:
'   - RangeHasValue            : Checks if any cell equals/contains a given value.
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


