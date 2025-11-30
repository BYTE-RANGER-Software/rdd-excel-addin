Attribute VB_Name = "modSheets"
' -----------------------------------------------------------------------------------
' Module    : modSheets
' Purpose   : Provides helper routines for working with worksheets and charts,
'             including lookup, creation, and tag-based selection.
'
' Public API:
'   - SheetNameMatchMode          : Enum defining name matching strategies.
'   - EnsureSheet                 : Ensures a sheet exists, creating it if necessary.
'   - SheetCodeNameExists         : Checks if a sheet exists by VBA CodeName.
'   - GetSheetByCodeName          : Retrieves a sheet by VBA CodeName.
'
' Private Helpers:
'   - SheetNameMatches            : Internal helper for sheet name matching.
'
' Dependencies:
'
' Notes     :
'   - Keep this module focused on sheet/chart lookup and selection utilities.
' -----------------------------------------------------------------------------------
Option Explicit
Option Private Module

' ===== Public API ====================================================================
' Public enum and helper functions for sheet and chart lookup.

' Match modes for sheet name selection
Public Enum SheetNameMatchMode
    SNMM_Exact = 0      ' exact match
    SNMM_Prefix = 1     ' starts with pattern
    SNMM_Suffix = 2     ' ends with pattern
    SNMM_Contains = 3   ' substring
    SNMM_Wildcard = 4   ' VBA Like pattern, e.g., "Room*"
    SNMM_Regex = 5      ' VBScript.RegExp (optional)
End Enum

' -----------------------------------------------------------------------------------
' Function  : EnsureSheet
' Purpose   : Ensures that a worksheet with the given 'name' exists in the specified
'             workbook. If not found create a new worksheet with 'name'.
'
' Parameters:
'   sheetName       [String]    - Name of the worksheet to ensure
'   targetWorkbook  [Workbook]  - Optional, Target workbook. (defaults to ActiveWorkbook)
'
' Returns   : Worksheet object of the found or newly created sheet
'
' Notes     :
' -----------------------------------------------------------------------------------
Public Function EnsureSheet(sheetName As String, Optional targetWorkbook As Workbook = Nothing) As Worksheet
    On Error Resume Next
    If targetWorkbook Is Nothing Then Set targetWorkbook = ActiveWorkbook
    Set EnsureSheet = targetWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If EnsureSheet Is Nothing Then
        Set EnsureSheet = targetWorkbook.Worksheets.Add(After:=Sheets(Sheets.Count))
        EnsureSheet.Name = sheetName
    End If
End Function

' -----------------------------------------------------------------------------------
' Function  : SheetCodeNameExists
' Purpose   : Checks if a worksheet with the specified **code name** exists
'             in the given workbook (or ActiveWorkbook by default).
'
' Parameters:
'   sheetCodeName [String]     - The code name of the worksheet to look for
'   targetWorkbook    [Workbook]   - Optional, The workbook to search in (defaults to ActiveWorkbook)
'
' Returns:
'   Boolean - True if a matching code name is found, False otherwise
'
' Notes:
'   - Code names are set in the VBA editor, not via sheet tabs
'   - Comparison is case-sensitive
' -----------------------------------------------------------------------------------
Public Function SheetCodeNameExists(sheetCodeName As String, Optional targetWorkbook As Workbook = Nothing) As Boolean
    Dim foundSheet As Worksheet
    
    If targetWorkbook Is Nothing Then Set targetWorkbook = ActiveWorkbook
    For Each foundSheet In targetWorkbook.Worksheets
        If foundSheet.CodeName = sheetCodeName Then
            SheetCodeNameExists = True
            Exit Function
        End If
    Next foundSheet
    SheetCodeNameExists = False
End Function

' -----------------------------------------------------------------------------------
' Function  : GetSheetByCodeName
' Purpose   : Returns the worksheet object that matches the specified code name
'             from the given workbook (or ActiveWorkbook by default).
'
' Parameters:
'   sheetCodeName     [String]     - The code name to search for
'   targetWorkbook    [Workbook]   - Optional, The workbook to search in
'
' Returns:
'   Worksheet - The matching worksheet object, or Nothing if not found
'
' Notes:
'   - Code names are those seen in the VBA editor (e.g., "Sheet1"), not tab names
'   - Comparison is case-sensitive
' -----------------------------------------------------------------------------------
Public Function GetSheetByCodeName(sheetCodeName As String, Optional targetWorkbook As Workbook = Nothing) As Worksheet
    Dim foundSheet As Worksheet
    
    If targetWorkbook Is Nothing Then Set targetWorkbook = ActiveWorkbook
    For Each foundSheet In targetWorkbook.Worksheets
        If foundSheet.CodeName = sheetCodeName Then
            Set GetSheetByCodeName = foundSheet
            Exit Function
        End If
    Next foundSheet
    Set GetSheetByCodeName = Nothing
End Function

' -----------------------------------------------------------------------------------
' Function  : SheetNameMatches
' Purpose   : Evaluates whether a sheet name matches a pattern based on the
'             provided SheetNameMatchMode and optional regular expression.
'
' Parameters:
'   sheetName   [String]             - Name to test.
'   patternValue    [String]             - Pattern to match against.
'   matchMode       [SheetNameMatchMode] - Matching strategy.
'   regExObject     [Object]             - VBScript.RegExp instance (may be Nothing).
'   useRegex        [Boolean]            - Indicates whether regExObject is valid.
'
' Returns   : Boolean - True if the name matches; otherwise False.
'
' Notes     :
'   - Internal helper; not intended as public API.
' -----------------------------------------------------------------------------------
Private Function SheetNameMatches(ByVal sheetName As String, _
                                  ByVal patternValue As String, _
                                  ByVal matchMode As SheetNameMatchMode, _
                                  ByVal regExObject As Object, _
                                  ByVal useRegex As Boolean) As Boolean
    Select Case matchMode
        Case SNMM_Exact
            SheetNameMatches = (sheetName = patternValue)
        Case SNMM_Prefix
            SheetNameMatches = (Left$(sheetName, Len(patternValue)) = patternValue)
        Case SNMM_Suffix
            SheetNameMatches = (Right$(sheetName, Len(patternValue)) = patternValue)
        Case SNMM_Contains
            SheetNameMatches = (InStr(1, sheetName, patternValue, vbTextCompare) > 0)
        Case SNMM_Wildcard
            SheetNameMatches = (sheetName Like patternValue)
        Case SNMM_Regex
            SheetNameMatches = (useRegex And regExObject.Test(sheetName))
        Case Else
            SheetNameMatches = False
    End Select
End Function
