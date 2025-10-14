Attribute VB_Name = "modUtil"
' modUtil
Option Explicit
Option Private Module

' -----------------------------------------------------------------------------------
' Function  : EnsureSheet
' Purpose   : Ensures that a worksheet with the given 'name' exists in the specified
'             workbook. If not found create a new worksheet with 'name'.
'
' Parameters:
'   strName   [String]    - Name of the worksheet to ensure
'   wbTarget  [Workbook]  - (Optional) Target workbook; defaults to ActiveWorkbook
'
' Returns   : Worksheet object of the found or newly created sheet
'
' Notes     :
' -----------------------------------------------------------------------------------
Public Function EnsureSheet(strName As String, Optional wbTarget As Workbook = Nothing) As Worksheet
    On Error Resume Next
    If wbTarget Is Nothing Then Set wbTarget = ActiveWorkbook
    Set EnsureSheet = wbTarget.Worksheets(strName)
    On Error GoTo 0
    If EnsureSheet Is Nothing Then
        Set EnsureSheet = wbTarget.Worksheets.Add(After:=Sheets(Sheets.Count))
        EnsureSheet.name = strName
    End If
End Function

' -----------------------------------------------------------------------------------
' Function  : SheetExists
' Purpose   : Checks whether a worksheet with the specified name exists
'             in the given workbook (or ActiveWorkbook by default).
'
' Parameters:
'   strName      [String]     - The name of the worksheet to search for
'   wbTarget     [Workbook]   - (Optional) The workbook to search in
'
' Returns:
'   Boolean - True if the sheet exists, False otherwise
'
' Notes:
'   - Uses error handling to prevent runtime error if sheet doesn't exist
' -----------------------------------------------------------------------------------
Public Function SheetExists(strName As String, Optional wbTarget As Workbook = Nothing) As Boolean
    Dim wksSheet As Worksheet
    
    On Error Resume Next
    If wbTarget Is Nothing Then Set wbTarget = ActiveWorkbook
    Set wksSheet = wbTarget.Worksheets(strName)
    SheetExists = Not wksSheet Is Nothing
    On Error GoTo 0
End Function

' -----------------------------------------------------------------------------------
' Function  : SheetCodeNameExists
' Purpose   : Checks if a worksheet with the specified **code name** exists
'             in the given workbook (or ActiveWorkbook by default).
'
' Parameters:
'   strCodeName [String]     - The code name of the worksheet to look for
'   wbTarget    [Workbook]   - (Optional) The workbook to search in
'
' Returns:
'   Boolean - True if a matching code name is found, False otherwise
'
' Notes:
'   - Code names are set in the VBA editor, not via sheet tabs
'   - Comparison is case-sensitive
' -----------------------------------------------------------------------------------
Public Function SheetCodeNameExists(strCodeName As String, Optional wbTarget As Workbook = Nothing) As Boolean
    Dim wksSheet As Worksheet
    
    If wbTarget Is Nothing Then Set wbTarget = ActiveWorkbook
    For Each wksSheet In wbTarget.Worksheets
        If wksSheet.codeName = strCodeName Then
            SheetCodeNameExists = True
            Exit Function
        End If
    Next wksSheet
    SheetCodeNameExists = False
End Function

' -----------------------------------------------------------------------------------
' Function  : GetSheetByCodeName
' Purpose   : Returns the worksheet object that matches the specified code name
'             from the given workbook (or ActiveWorkbook by default).
'
' Parameters:
'   strCodeName [String]     - The code name to search for
'   wbTarget    [Workbook]   - (Optional) The workbook to search in
'
' Returns:
'   Worksheet - The matching worksheet object, or Nothing if not found
'
' Notes:
'   - Code names are those seen in the VBA editor (e.g., "Sheet1"), not tab names
'   - Comparison is case-sensitive
' -----------------------------------------------------------------------------------
Public Function GetSheetByCodeName(strCodeName As String, Optional wbTarget As Workbook = Nothing) As Worksheet
    Dim wksSheet As Worksheet
    
    If wbTarget Is Nothing Then Set wbTarget = ActiveWorkbook
    For Each wksSheet In wbTarget.Worksheets
        If wksSheet.codeName = strCodeName Then
            Set GetSheetByCodeName = wksSheet
            Exit Function
        End If
    Next wksSheet
    Set GetSheetByCodeName = Nothing
End Function

' ---------------------------------------------------------------
' Procedure : HideOpMode
' Purpose   : Enables or disables Excel's interactive features
'             to optimize performance and suppress distractions
'             during automated operations.
'
' Parameters:
'   blnEnable [Boolean] - If True, disables events, screen updates,
'                         animations, and alerts (silent mode).
'                         If False, restores normal behavior.
'
' Usage     : Call HideOpMode True before automation,
'             and HideOpMode False afterwards.
' ---------------------------------------------------------------
Public Sub HideOpMode(ByVal blnEnable As Boolean)
    Application.EnableEvents = Not blnEnable
    Application.ScreenUpdating = Not blnEnable
    Application.EnableAnimations = Not blnEnable
    Application.DisplayAlerts = Not blnEnable
End Sub

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
    strRef = "='" & wks.name & "'!" & wks.Range(wks.Cells(2, lngCol), wks.Cells(lngLastRow, lngCol)).Address
    On Error Resume Next
    wbActive.Names(strName).RefersTo = strRef
    If Err.Number <> 0 Then
        wbActive.Names.Add name:=strName, RefersTo:=strRef
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
'   lngStartRow [Long, Optional]  - First data row, default 2.
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

' -----------------------------------------------------------------------------------
' Function  : FindFramedRangeByHeading
' Purpose   : Finds the rectangular, border-framed range that contains the specified
'             heading text and returns that range.
'
' Parameters:
'   wks          [Worksheet]        - Worksheet to search in
'   strHeading   [String]           - Heading text to locate
'   blnMatchWhole[Boolean, Optional]- If True, matches the whole cell (xlWhole);
'                                     if False, allows partial match (xlPart). Default True.
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
                                         Optional blnMatchWhole As Boolean = True) As Range
    Dim rngFound As Range, lngTopRow As Long, lngBottomRow As Long
    Dim lngLeftCol As Long, lngRightCol As Long, lngRow As Long, lngCol As Long
    Dim enuLookAtMode As XlLookAt
    enuLookAtMode = IIf(blnMatchWhole, xlWhole, xlPart)
    
    Set rngFound = wks.Cells.Find(What:=strHeading, LookIn:=xlValues, LookAt:=enuLookAtMode, _
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
    
    Set FindFramedRangeByHeading = wks.Range(wks.Cells(lngTopRow, lngLeftCol), wks.Cells(lngBottomRow, lngRightCol))
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
        CellHasBorder = (.LineStyle = xlContinuous And .weight = xlWeight)
    End With
End Function
