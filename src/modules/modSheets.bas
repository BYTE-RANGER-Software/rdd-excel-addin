Attribute VB_Name = "modSheets"
Option Explicit
Option Private Module

' -----------------------------------------------------------------------------------
' Function  : EnsureSheet
' Purpose   : Ensures that a worksheet with the given 'name' exists in the specified
'             workbook. If not found create a new worksheet with 'name'.
'
' Parameters:
'   strName   [String]    - Name of the worksheet to ensure
'   wbTarget  [Workbook]  - Optional, Target workbook. (defaults to ActiveWorkbook)
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
        EnsureSheet.Name = strName
    End If
End Function

' -----------------------------------------------------------------------------------
' Function  : SheetExists
' Purpose   : Checks whether a worksheet with the specified name exists
'             in the given workbook (or ActiveWorkbook by default).
'
' Parameters:
'   strName      [String]     - The name of the worksheet to search for
'   wbTarget     [Workbook]   - Optional, The workbook to search in (defaults to ActiveWorkbook)
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
'   wbTarget    [Workbook]   - Optional, The workbook to search in (defaults to ActiveWorkbook)
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
'   wbTarget    [Workbook]   - Optional, The workbook to search in
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

' Build a dictionary of sheets filtered by name.
' Key = sheet name, Item = Worksheet (or Chart if includeCharts=True).
Public Function BuildDictFromSheetsByName( _
        ByVal wb As Workbook, _
        ByVal strPattern As String, _
        Optional ByVal mode As SheetNameMatchMode = SNMM_Exact, _
        Optional ByVal wksExclude As Worksheet = Nothing, _
        Optional ByVal blnIgnoreCase As Boolean = True, _
        Optional ByVal includeCharts As Boolean = False, _
        Optional ByVal excludeHidden As Boolean = False _
    ) As Object
    
    Dim dic As Object: Set dic = CreateObject("Scripting.Dictionary")
    Dim strPat As String: strPat = IIf(blnIgnoreCase, LCase$(strPattern), strPattern)
    Dim strName As String
    Dim wks As Worksheet
    Dim cht As Chart
    Dim objRegEx As Object, useRegex As Boolean
    
    ' Prepare RegExp if requested
    If mode = SNMM_Regex Then
        On Error Resume Next
        Set objRegEx = CreateObject("VBScript.RegExp")
        On Error GoTo 0
        If Not objRegEx Is Nothing Then
            objRegEx.pattern = strPattern
            objRegEx.blnIgnoreCase = blnIgnoreCase
            objRegEx.Global = False
            useRegex = True
        Else
            ' Fallback to substring if RegExp is unavailable
            mode = SNMM_Contains
        End If
    End If
    
    ' Worksheets
    For Each wks In wb.Worksheets
        If Not wksExclude Is Nothing Then
            If wks.Name = wksExclude.Name Then GoTo NextWS
        End If
        If excludeHidden And wks.Visible = xlSheetHidden Or wks.Visible = xlSheetVeryHidden Then GoTo NextWS
        
        strName = IIf(blnIgnoreCase, LCase$(wks.Name), wks.Name)
        If SheetNameMatches(strName, strPat, mode, objRegEx, useRegex) Then dic(wks.Name) = wks
NextWS:
    Next wks
    
    ' Optional: Charts
    If includeCharts Then
        For Each cht In wb.Charts
            strName = IIf(blnIgnoreCase, LCase$(cht.Name), cht.Name)
            If SheetNameMatches(strName, strPat, mode, objRegEx, useRegex) Then dic(cht.Name) = cht
        Next cht
    End If
    
    Set BuildDictFromSheetsByName = dic
End Function

' Helper: compare a name with pattern using the selected mode
Private Function SheetNameMatches(ByVal strName As String, ByVal strPat As String, _
                             ByVal mode As SheetNameMatchMode, _
                             ByVal objRegEx As Object, ByVal useRegex As Boolean) As Boolean
    Select Case mode
        Case SNMM_Exact:    SheetNameMatches = (strName = strPat)
        Case SNMM_Prefix:   SheetNameMatches = (Left$(strName, Len(strPat)) = strPat)
        Case SNMM_Suffix:   SheetNameMatches = (Right$(strName, Len(strPat)) = strPat)
        Case SNMM_Contains: SheetNameMatches = (InStr(1, strName, strPat, vbTextCompare) > 0)
        Case SNMM_Wildcard: SheetNameMatches = (strName Like strPat)
        Case SNMM_Regex:    SheetNameMatches = (useRegex And objRegEx.Test(strName))
        Case Else:       SheetNameMatches = False
    End Select
End Function
