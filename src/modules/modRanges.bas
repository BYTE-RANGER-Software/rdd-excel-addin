Attribute VB_Name = "modRanges"
Option Explicit
Option Private Module

' -----------------------------------------------------------------------------------
' Function  : FindFramedRangeByHeading
' Purpose   : Finds the rectangular, border-framed range that contains the specified
'             heading text and returns that range.
'
' Parameters:
'   wks             [Worksheet]        - Worksheet to search in
'   strHeading      [String]           - Heading text to locate
'   blnMatchWhole   [Boolean]          - Optional, If True, matches the whole cell (xlWhole). (default True);
'                                        if False, allows partial match (xlPart). (Default True.)
'   blnWithoutHeader[Boolean]          - Optional, If True, does not add the header cell to the return range. (default True).
'
' Returns:
'   Range - The framed range bounded by left/right/top/bottom borders that encloses
'           the heading; returns Nothing if not found or if the frame is incomplete.
'
' Behavior  :
'   - Uses Cells.Find (xlValues, case-insensitive) to locate the heading.
'   - If the heading cell is merged, uses the first cell of the MergeArea.
'   - Scans left/right from the heading row to detect edge borders,
'     scans up to find the top edge, and scans down until both left & right
'     edge cells have bottom borders to define the bottom edge.
'   - Exits with Nothing if any edge cannot be determined.
'
' Notes     :
'   - Designed for bordered "boxes" in structured sheets.
'   - Helper function CellHasBorder is used to test specific cell edges.
' -----------------------------------------------------------------------------------
Public Function FindFramedRangeByHeading(wks As Worksheet, strHeading As String, _
                                         Optional blnMatchWhole As Boolean = True, _
                                         Optional ByVal blnWithoutHeader As Boolean = True) As Range
    Dim rngFound As Range, lngTopRow As Long, lngBottomRow As Long
    Dim lngLeftCol As Long, lngRightCol As Long, lngRow As Long, lngCol As Long
    Dim enuLookAtMode As XlLookAt
    enuLookAtMode = IIf(blnMatchWhole, xlWhole, xlPart)
    
    Set rngFound = wks.Cells.Find(What:=strHeading, LookIn:=xlValues, lookat:=enuLookAtMode, _
                                  MatchCase:=False, SearchOrder:=xlByRows, SearchDirection:=xlNext)
    If rngFound Is Nothing Then Exit Function
    
    ' Work with the first cell of a merged heading area
    If rngFound.MergeCells Then Set rngFound = rngFound.MergeArea.Cells(1, 1)
    
    ' Find left edge: move left until a left border exists
    lngCol = rngFound.Column
    Do While lngCol > 1 And Not CellHasBorder(wks.Cells(rngFound.Row, lngCol), xlEdgeLeft)
        lngCol = lngCol - 1
    Loop
    If Not CellHasBorder(wks.Cells(rngFound.Row, lngCol), xlEdgeLeft) Then Exit Function
    lngLeftCol = lngCol
    
    ' Find right edge: move right until a right border exists
    lngCol = rngFound.Column
    Do While lngCol < wks.Columns.Count And Not CellHasBorder(wks.Cells(rngFound.Row, lngCol), xlEdgeRight)
        lngCol = lngCol + 1
    Loop
    If Not CellHasBorder(wks.Cells(rngFound.Row, lngCol), xlEdgeRight) Then Exit Function
    lngRightCol = lngCol
    
    ' Find top edge: move up until a top border exists
    lngRow = rngFound.Row
    Do While lngRow > 1 And Not CellHasBorder(wks.Cells(lngRow, rngFound.Column), xlEdgeTop)
        lngRow = lngRow - 1
    Loop
    If Not CellHasBorder(wks.Cells(lngRow, rngFound.Column), xlEdgeTop) Then Exit Function
    lngTopRow = lngRow
    
    ' Find bottom edge: scan down until BOTH left & right edge cells have bottom borders
    lngRow = rngFound.Row
    Do While lngRow < wks.Rows.Count
        If CellHasBorder(wks.Cells(lngRow, lngLeftCol), xlEdgeBottom) And _
           CellHasBorder(wks.Cells(lngRow, lngRightCol), xlEdgeBottom) Then
            lngBottomRow = lngRow
            Exit Do
        End If
        lngRow = lngRow + 1
    Loop
    If lngBottomRow = 0 Then Exit Function
    
    Set FindFramedRangeByHeading = wks.Range(wks.Cells(lngTopRow + CInt(blnWithoutHeader), lngLeftCol), wks.Cells(lngBottomRow, lngRightCol))
End Function

' -----------------------------------------------------------------------------------
' Function  : GetColumnRangeByHeader
' Purpose   : Finds the column range starting at a header cell and ending at the
'             first cell whose bottom border is considered "strong".
'
' Parameters:
'   ws              [Worksheet]       - Worksheet to search in.
'   strHeader       [String]          - Header text to locate.
'   blnMatchWhole   [Boolean]         - Optional, If True, xlWhole match; if False, xlPart (default True).
'   xlBrdWeight     [XlBorderWeight]  - Optional, expected border weight (default xlMedium)
'   blnWithoutHeader[Boolean]         - Optional, If True, does not add the header cell to the return range (default True).
'
' Returns   : Range - From the header cell down to the detected bottom edge; Nothing if not found.
'
' Notes     :
'   - Uses Cells.Find with case-insensitive search; respects merged header cells.
'   - Detects the bottom edge via CellHasBorder(..., xlEdgeBottom, ...).
' -----------------------------------------------------------------------------------
Public Function GetColumnRangeByHeader( _
        ByVal ws As Worksheet, _
        ByVal strHeader As String, _
        Optional ByVal blnMatchWhole As Boolean = True, _
        Optional ByVal xlBrdWeight As XlBorderWeight = xlMedium, _
        Optional ByVal blnWithoutHeader As Boolean = True _
    ) As Range
    
    Dim rngHeader As Range
    Dim lCol As Long, lRow As Long, lLast As Long
    Dim eLook As XlLookAt: eLook = IIf(blnMatchWhole, xlWhole, xlPart)
    
    ' Locate the header cell
    Set rngHeader = ws.Cells.Find(What:=strHeader, LookIn:=xlValues, lookat:=eLook, _
                                  MatchCase:=False, SearchOrder:=xlByRows, SearchDirection:=xlNext)
    If rngHeader Is Nothing Then Exit Function
    If rngHeader.MergeCells Then Set rngHeader = rngHeader.MergeArea.Cells(1, 1)
    
    lCol = rngHeader.Column
    
    ' Scan down until a strong bottom border is found
    For lRow = rngHeader.Row To ws.Rows.Count
        If CellHasBorder(ws.Cells(lRow, lCol), xlEdgeBottom, xlBrdWeight) Then
            lLast = lRow
            Exit For
        End If
    Next lRow
    If lLast = 0 Then Exit Function
    
    Set GetColumnRangeByHeader = ws.Range(ws.Cells(rngHeader.Row + CInt(blnWithoutHeader), lCol), ws.Cells(lLast, lCol))
End Function

' -----------------------------------------------------------------------------------
' Function  : RangeHasFormula2
' Purpose   : Checks whether any cell in the given range has a Formula2 equal to the
'             provided text (without the leading "="), with optional exact/substring
'             and case-sensitive/insensitive comparison.
'
' Parameters:
'   rngSearch        [Range]   - Range to search within.
'   strFormulaNoEq   [String]  - Formula text without the leading "=".
'   blnExact         [Boolean] - If True, exact match; if False, substring match (default True).
'   blnCaseSensitive [Boolean] - If True, case-sensitive; otherwise case-insensitive (default False).
'
' Returns   : Boolean - True if a matching Formula2 is found; otherwise False.
'
' Notes     :
'   - Limits scan to formula cells using SpecialCells(xlCellTypeFormulas); returns False if none.
'   - Normalizes both sides via NormalizeFormula2Text (trims, removes "=", unifies separators, strips spaces/CR/LF).
' -----------------------------------------------------------------------------------
Public Function RangeHasFormula2( _
        ByVal rngSearch As Range, _
        ByVal strFormulaNoEq As String, _
        Optional ByVal blnExact As Boolean = True, _
        Optional ByVal blnCaseSensitive As Boolean = False _
    ) As Boolean
    
    Dim rngFormulas As Range
    Dim rngCell As Range
    Dim strTarget As String, strCell As String
    
    If rngSearch Is Nothing Then Exit Function
    
    ' Limit to formula cells for speed; handle "no formulas" case.
    On Error Resume Next
    Set rngFormulas = rngSearch.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    If rngFormulas Is Nothing Then Exit Function
    
    strTarget = NormalizeFormula2Text(strFormulaNoEq, blnCaseSensitive)
    
    For Each rngCell In rngFormulas.Cells
        strCell = NormalizeFormula2Text(rngCell.Formula2, blnCaseSensitive)
        
        If blnExact Then
            If strCell = strTarget Then
                RangeHasFormula2 = True
                Exit Function
            End If
        Else
            ' substring match if you ever need it
            If InStr(1, strCell, strTarget, IIf(blnCaseSensitive, vbBinaryCompare, vbTextCompare)) > 0 Then
                RangeHasFormula2 = True
                Exit Function
            End If
        End If
    Next rngCell
End Function

' -----------------------------------------------------------------------------------
' Function  : RangeHasValue
' Purpose   : Text search can be exact or substring match, numeric and date are exact.
'
' Parameters:
'   rngSearch        [Range]   - Range to search within.
'   vntValue         [Variant] - value to search for.
'   blnExact         [Boolean] - If True, exact match; if False, substring match (default True).
'   blnCaseSensitive [Boolean] - If True, case-sensitive; otherwise case-insensitive (default False).
'
' Returns   : Boolean - True if any cell in rngSearch contains the given value.
'
' Notes     :
' -----------------------------------------------------------------------------------
Public Function RangeHasValue( _
        ByVal rngSearch As Range, _
        ByVal vntValue As Variant, _
        Optional ByVal blnExact As Boolean = True, _
        Optional ByVal blnCaseSensitive As Boolean = False _
    ) As Boolean
    
    Dim c As Range
    Dim vCell As Variant
    Dim cmpType As VbCompareMethod
    cmpType = IIf(blnCaseSensitive, vbBinaryCompare, vbTextCompare)
    
    If rngSearch Is Nothing Then Exit Function
    
    For Each c In rngSearch.Cells
        vCell = c.Value
        
        If IsEmpty(vCell) Then GoTo NextCell
        
        Select Case VarType(vntValue)
            Case vbString
                If VarType(vCell) = vbString Then
                    If blnExact Then
                        If StrComp(vCell, vntValue, cmpType) = 0 Then
                            RangeHasValue = True
                            Exit Function
                        End If
                    Else
                        If InStr(1, vCell, vntValue, cmpType) > 0 Then
                            RangeHasValue = True
                            Exit Function
                        End If
                    End If
                End If
            
            Case vbDouble, vbSingle, vbCurrency, vbInteger, vbLong, vbDecimal
                If IsNumeric(vCell) Then
                    If CDbl(vCell) = CDbl(vntValue) Then
                        RangeHasValue = True
                        Exit Function
                    End If
                End If
            
            Case vbDate
                If IsDate(vCell) Then
                    If CLng(vCell) = CLng(vntValue) Then
                        RangeHasValue = True
                        Exit Function
                    End If
                End If
            
            Case Else
                ' fallback to string compare
                If blnExact Then
                    If StrComp(CStr(vCell), CStr(vntValue), cmpType) = 0 Then
                        RangeHasValue = True
                        Exit Function
                    End If
                Else
                    If InStr(1, CStr(vCell), CStr(vntValue), cmpType) > 0 Then
                        RangeHasValue = True
                        Exit Function
                    End If
                End If
        End Select
NextCell:
    Next c
End Function


'--- helpers ---

' -----------------------------------------------------------------------------------
' Function  : NormalizeFormula2Text
' Purpose   : Normalizes a Formula2 string for comparison by removing cosmetic
'             differences and harmonizing list separators.
'
' Parameters:
'   strFormula       [String]   - Formula text (with or without leading "=").
'   blnCaseSensitive [Boolean]  - If False, folds to lower-case; if True, preserves case.
'
' Returns   : String - Normalized formula text (no leading "=", unified separators, no spaces/CR/LF).
'
' Notes     :
'   - Uses Application.International(xlListSeparator) to harmonize "," and ";".
'   - Trims whitespace, removes leading "=", removes spaces and CR/LF characters.
' -----------------------------------------------------------------------------------
Private Function NormalizeFormula2Text(ByVal strFormula As String, ByVal blnCaseSensitive As Boolean) As String
    Dim strSep As String
    strSep = Application.International(xlListSeparator)
    
    strFormula = Trim$(strFormula)
    If Left$(strFormula, 1) = "=" Then strFormula = Mid$(strFormula, 2)
    
    ' unify list separators, user may type "," while UI uses ";", or vice versa
    strFormula = Replace(strFormula, ",", strSep)
    strFormula = Replace(strFormula, ";", strSep)
    
    ' remove spaces and CR/LF to avoid cosmetic mismatches
    strFormula = Replace(strFormula, " ", vbNullString)
    strFormula = Replace(strFormula, vbCr, vbNullString)
    strFormula = Replace(strFormula, vbLf, vbNullString)
    
    If Not blnCaseSensitive Then strFormula = LCase$(strFormula)
    
    NormalizeFormula2Text = strFormula
End Function

' -----------------------------------------------------------------------------------
' Function  : CellHasBorder
' Purpose   : Checks whether a specific edge of a cell has a continuous border
'             with the specified weight.
'
' Parameters:
'   rngCell [Range]           - Target cell to inspect
'   xlEdge  [XlBordersIndex]  - Edge to check (xlEdgeLeft, xlEdgeRight, xlEdgeTop, xlEdgeBottom)
'   xlWeight[XlBorderWeight]  - Optional, expected border weight, default xlMedium
'
' Returns:
'   Boolean - True if the edge uses LineStyle xlContinuous and matches the given weight,
'             otherwise False.
'
' Notes:
'   - Intended for framed-range detection used by FindFramedRangeByHeading.
' -----------------------------------------------------------------------------------
Private Function CellHasBorder(rngCell As Range, xlEdge As XlBordersIndex, Optional xlWeight As XlBorderWeight = xlMedium) As Boolean
    With rngCell.Borders(xlEdge)
        CellHasBorder = (.LineStyle = xlContinuous And .Weight = xlWeight)
    End With
End Function

