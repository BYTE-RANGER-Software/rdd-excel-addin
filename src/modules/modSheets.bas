Attribute VB_Name = "modSheets"
' -----------------------------------------------------------------------------------
' Module    : modSheets
' Purpose   : Provides helper routines for working with worksheets and charts,
'             including lookup, creation, and tag-based selection.
'
' Public API:
'   - SheetNameMatchMode          : Enum defining name matching strategies.
'   - EnsureSheet                 : Ensures a sheet exists, creating it if necessary.
'   - SheetExists                 : Checks if a sheet exists by name.
'   - SheetCodeNameExists         : Checks if a sheet exists by VBA CodeName.
'   - GetSheetByCodeName          : Retrieves a sheet by VBA CodeName.
'   - BuildDictFromSheetsByName   : Builds a dictionary of sheets/charts by name.
'   - BuildDictFromSheetsByTag    : Builds a dictionary of sheets by tag.
'
' Private Helpers:
'   - SheetNameMatches            : Internal helper for sheet name matching.
'
' Dependencies:
'   - modTags (for BuildDictFromSheetsByTag)
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
' Function  : SheetExists
' Purpose   : Checks whether a worksheet with the specified name exists
'             in the given workbook (or ActiveWorkbook by default).
'
' Parameters:
'   sheetName      [String]     - The name of the worksheet to search for
'   targetWorkbook [Workbook]   - Optional, The workbook to search in (defaults to ActiveWorkbook)
'
' Returns:
'   Boolean - True if the sheet exists, False otherwise
'
' Notes:
'   - Uses error handling to prevent runtime error if sheet doesn't exist
' -----------------------------------------------------------------------------------
Public Function SheetExists(sheetName As String, Optional targetWorkbook As Workbook = Nothing) As Boolean
    Dim foundSheet As Worksheet
    
    On Error Resume Next
    If targetWorkbook Is Nothing Then Set targetWorkbook = ActiveWorkbook
    Set foundSheet = targetWorkbook.Worksheets(sheetName)
    SheetExists = Not foundSheet Is Nothing
    On Error GoTo 0
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
' Function  : BuildDictFromSheetsByName
' Purpose   : Builds a dictionary of sheets (and optionally charts) whose names match
'             a given pattern using the specified match mode.
'
' Parameters:
'   srcWorkbook     [Workbook]           - Workbook to scan.
'   namePattern     [String]             - Name or pattern used for matching.
'   mode            [SheetNameMatchMode] - Matching strategy (exact, prefix, etc.).
'   sheetToExclude  [Worksheet]          - Optional, sheet to exclude from results.
'   ignoreCase      [Boolean]            - Optional, if True compare ignoring case.
'   includeCharts   [Boolean]            - Optional, if True include charts as well.
'   excludeHidden   [Boolean]            - Optional, if True exclude hidden sheets.
'
' Returns   : Scripting.Dictionary - with Keys = sheet/chart name,
'                                    Items = Worksheet or Chart objects.
'
' Notes     :
'   - Never returns Nothing; returns an empty dictionary when there are no matches.
' -----------------------------------------------------------------------------------
Public Function BuildDictFromSheetsByName( _
        ByVal srcWorkbook As Workbook, _
        ByVal namePattern As String, _
        Optional ByVal mode As SheetNameMatchMode = SNMM_Exact, _
        Optional ByVal sheetToExclude As Worksheet = Nothing, _
        Optional ByVal ignoreCase As Boolean = True, _
        Optional ByVal includeCharts As Boolean = False, _
        Optional ByVal excludeHidden As Boolean = False _
    ) As Scripting.Dictionary
    
    Dim sheetDict As Scripting.Dictionary: Set sheetDict = New Scripting.Dictionary
    Dim patternValue As String: patternValue = IIf(ignoreCase, LCase$(namePattern), namePattern)
    Dim sheetName As String
    Dim Sh As Worksheet
    Dim cht As Chart
    Dim regExObject As Object
    Dim useRegex As Boolean
    
    ' Prepare RegExp if requested
    If mode = SNMM_Regex Then
        On Error Resume Next
        Set regExObject = CreateObject("VBScript.RegExp")
        On Error GoTo 0
        If Not regExObject Is Nothing Then
            regExObject.pattern = namePattern
            regExObject.ignoreCase = ignoreCase
            regExObject.Global = False
            useRegex = True
        Else
            ' Fallback to substring if RegExp is unavailable
            mode = SNMM_Contains
        End If
    End If
    
    ' Worksheets
    For Each Sh In srcWorkbook.Worksheets
        If Not sheetToExclude Is Nothing Then
            If Sh.Name = sheetToExclude.Name Then GoTo NextWS
        End If
        If excludeHidden And (Sh.Visible = xlSheetHidden Or Sh.Visible = xlSheetVeryHidden) Then GoTo NextWS
        
        sheetName = IIf(ignoreCase, LCase$(Sh.Name), Sh.Name)
        If SheetNameMatches(sheetName, patternValue, mode, regExObject, useRegex) Then sheetDict(Sh.Name) = Sh
NextWS:
    Next Sh
    
    ' Optional: Charts
    If includeCharts Then
        For Each cht In srcWorkbook.Charts
            sheetName = IIf(ignoreCase, LCase$(cht.Name), cht.Name)
            If SheetNameMatches(sheetName, patternValue, mode, regExObject, useRegex) Then sheetDict(cht.Name) = cht
        Next cht
    End If
    
    Set BuildDictFromSheetsByName = sheetDict
End Function

' -----------------------------------------------------------------------------------
' Function  : BuildDictFromSheetsByTag
' Purpose   : Build a dictionary of worksheets in a workbook that carry a given tag
'             (stored as a worksheet CustomProperty via your tagging system).
'
' Parameters:
'   srcWorkbook           [Workbook] - Source workbook to scan.
'   tagName               [String]   - Tag name (free-text as used with modTags.*).
'   sheetToExclude        [Worksheet]- (Optional) Sheet to exclude from results.
'   excludeHidden         [Boolean]  - (Optional) Exclude hidden/very hidden sheets.
'   tagValueFilter        [String]   - (Optional) When provided, only include sheets
'                                      whose tag value equals this string.
'   valueCaseSensitive    [Boolean]  - (Optional) Case-sensitive value comparison.
'
' Returns   :Scripting.Dictionary - with Keys  = Worksheet.Name,
'                                        Items = Worksheet object. Never returns Nothing (empty on no matches).
'
' Notes     :
'   - Uses modTags.HasSheetTag(sheet, tagName, rValue) to probe for tags.
'   - If tagValueFilter is empty, only the presence of the tag is required.
' -----------------------------------------------------------------------------------
Public Function BuildDictFromSheetsByTag( _
        ByVal srcWorkbook As Workbook, _
        ByVal tagName As String, _
        Optional ByVal sheetToExclude As Worksheet = Nothing, _
        Optional ByVal excludeHidden As Boolean = False, _
        Optional ByVal tagValueFilter As String = vbNullString, _
        Optional ByVal valueCaseSensitive As Boolean = False _
    ) As Scripting.Dictionary

    On Error GoTo ErrHandler

    Dim sheetDict As Scripting.Dictionary
    Set sheetDict = New Scripting.Dictionary

    If srcWorkbook Is Nothing Then
        Set BuildDictFromSheetsByTag = sheetDict
        Exit Function
    End If

    Dim sheet As Worksheet
    Dim tagValue As String
    Dim cmpMode As VbCompareMethod
    cmpMode = IIf(valueCaseSensitive, vbBinaryCompare, vbTextCompare)

    For Each sheet In srcWorkbook.Worksheets
        ' exclude a specific sheet if requested
        If Not sheetToExclude Is Nothing Then
            If sheet.Name = sheetToExclude.Name Then GoTo NextSheet
        End If

        ' optionally skip hidden sheets
        If excludeHidden Then
            If (sheet.Visible = xlSheetHidden) Or (sheet.Visible = xlSheetVeryHidden) Then GoTo NextSheet
        End If

        ' tag check (presence and optional value filter)
        If modTags.HasSheetTag(sheet, tagName, tagValue) Then
            If LenB(tagValueFilter) = 0 Then
                sheetDict(sheet.Name) = sheet
            ElseIf StrComp(CStr(tagValue), CStr(tagValueFilter), cmpMode) = 0 Then
                sheetDict(sheet.Name) = sheet
            End If
        End If

NextSheet:
    Next sheet

    Set BuildDictFromSheetsByTag = sheetDict
    Exit Function

ErrHandler:
    ' Fail-safe: still return an empty dictionary
    Dim dicSafe As Scripting.Dictionary
    Set dicSafe = New Scripting.Dictionary
    Set BuildDictFromSheetsByTag = dicSafe
    Err.Clear
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
