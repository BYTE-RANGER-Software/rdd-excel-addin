Attribute VB_Name = "modTags"
' ===================================================================================
' Module    : modTags
' Purpose   : Provide a small tagging system for worksheets and workbook using CustomProperties.
'             Tags are stored as worksheet-level CustomProperties or workbook-level CustomDocumentPropteries whose names are
'             normalized (safe) and prefixed with a constant to avoid collisions.
''
' Notes     :
'   - Only the CustomProperty 'name' is sanitized; the stored 'value' remains
'     the original free-text input (where provided).
'   - Uses module 'modProps'
' ===================================================================================
Option Explicit
Option Private Module

' ===== Constants (Private) =========================================================

Private Const TAG_PREFIX As String = "__TAG__" ' unique prefix for sheet tag

' ===== Private Helpers =============================================================

' -----------------------------------------------------------------------------------
' Function  : SanitizeTag
' Purpose   : Create a safe token (letters, digits, underscore) from free-text tag.
' Parameters:
'   strTagText [String] - Free-text tag input.
' Returns   : String - Lower-cased, sanitized token (non-matching chars mapped to "_").
' Notes     : Used to build a safe CustomProperty name; does not change the stored value.
' -----------------------------------------------------------------------------------
Private Function SanitizeTag(ByVal strTagText As String) As String
    Dim lngIndex As Long, strCh As String, strWork As String
    strWork = Trim$(strTagText)
    If Len(strWork) = 0 Then strWork = "tag"
    For lngIndex = 1 To Len(strWork)
        strCh = Mid$(strWork, lngIndex, 1)
        If strCh Like "[A-Za-z0-9_]" Then
            SanitizeTag = SanitizeTag & strCh
        Else
            SanitizeTag = SanitizeTag & "_"
        End If
    Next
    SanitizeTag = LCase$(SanitizeTag)
End Function

' -----------------------------------------------------------------------------------
' Function  : TagPropName
' Purpose   : Build the CustomProperty name used to store a tag on a sheet/workbook.
' Parameters:
'   strTagText [String] - Original (free-text) tag input.
'
' Returns   : String - Property name built as TAG_PREFIX & SanitizeTag(strTagText).
' Notes     :
' -----------------------------------------------------------------------------------
Private Function TagPropName(ByVal strTagText As String) As String
    TagPropName = TAG_PREFIX & SanitizeTag(strTagText)
End Function

' ===== Public API - Worksheet Tags =================================================

' -----------------------------------------------------------------------------------
' Procedure : TagSheet
' Purpose   : Add a tag to a worksheet, stored as a CustomProperty.
' Parameters:
'   wks        [Worksheet] - Target worksheet.
'   strTagText [String]    - Tag text (free-text). Stored as value; name is normalized.
'   strValue   [String]    - (Optional) Value of Tag
'
' Returns   : (none)
' Notes     : Requires SetCustomProperty().
' -----------------------------------------------------------------------------------
Public Sub TagSheet(ByVal wks As Worksheet, ByVal strTagText As String, Optional strValue As String = vbNullString)
    Dim strProp As String
    strProp = TagPropName(strTagText)
    ' tag name is normalized for safety
    SetCustomProperty wks, strProp, strValue
End Sub

' -----------------------------------------------------------------------------------
' Procedure : UntagSheet
' Purpose   : Remove a specific tag from a worksheet.
' Parameters:
'   wks        [Worksheet] - Target worksheet.
'   strTagText [String]    - Tag text to remove.
' Returns   : (none)
' Notes     : Safe if tag is not present (no error raised).
' -----------------------------------------------------------------------------------
Public Sub UntagSheet(ByVal wks As Worksheet, ByVal strTagText As String)
    Dim objCp As customProperty
    Dim strProp As String
    strProp = TagPropName(strTagText)
    If modProps.CustomPropertyExists(wks, strProp, objCp) Then
        objCp.Delete
    End If
End Sub

' -----------------------------------------------------------------------------------
' Function  : HasSheetTag
' Purpose   : Check whether a worksheet carries the given tag.
' Parameters:
'   wks        [Worksheet] - Target worksheet.
'   strTagText [String]    - Tag text to check.
'   r_strValue [String]    - (Optional, ByRef) returns the Value of Tag
'
' Returns   : Boolean - True if the tag exists on the worksheet; otherwise False.
' Notes     : Uses CustomPropertyExists to probe presence.
' -----------------------------------------------------------------------------------
Public Function HasSheetTag(ByVal wks As Worksheet, ByVal strTagText As String, Optional ByRef r_strValue As String = vbNullString) As Boolean
    Dim objCp As customProperty
    r_strValue = vbNullString
    
    If modProps.CustomPropertyExists(wks, TagPropName(strTagText), objCp) Then
        HasSheetTag = True
        On Error Resume Next
        r_strValue = CStr(objCp.Value)
        If Err.Number <> 0 Then
          
            modErr.ReportError "HasSheetTag", Err.Number, Erl, caption:=modMain.AppProjectName, customMessage:="Read value from " & strTagText & " failed on " & wks.Name
            Err.Clear
            r_strValue = vbNullString
        End If
        On Error GoTo 0
    End If
End Function

' -----------------------------------------------------------------------------------
' Function  : GetAllSheetsByTag
' Purpose   : Collect all worksheets in a workbook that carry the given tag.
' Parameters:
'   wb         [Workbook] - Source workbook.
'   strTagText [String]   - Tag text to search for.
' Returns   : Collection - Collection of Worksheet objects (may be empty).
' Notes     : Uses GetAllSheetsNamesByCustomProperty to find matching sheet names.
' -----------------------------------------------------------------------------------
Public Function GetAllSheetsByTag(ByVal wb As Workbook, ByVal strTagText As String) As Collection
    Dim colSheets As New Collection
    Dim astrNames() As String
    Dim lngIndex As Long
    Dim blnOk As Boolean

    blnOk = GetAllSheetsNamesByCustomProperty(wb, astrNames, TagPropName(strTagText))
    If blnOk Then
        For lngIndex = LBound(astrNames) To UBound(astrNames)
            colSheets.Add wb.Worksheets(astrNames(lngIndex))
        Next
    End If
    Set GetAllSheetsByTag = colSheets
End Function

' -----------------------------------------------------------------------------------
' Function  : GetSheetByTag
' Purpose   : Get first worksheets in a workbook that matches the given tag.
' Parameters:
'   wb         [Workbook] - Source workbook.
'   strTagText [String]   - Tag text to search for.
' Returns   : Worksheet - First matching worksheet (Nothing if none found).
' Notes     :
' -----------------------------------------------------------------------------------
Public Function GetSheetByTag(ByVal wb As Workbook, ByVal strTagText As String) As Worksheet
    Dim wks As Worksheet
    
    For Each wks In wb.Worksheets
        If HasSheetTag(wks, strTagText) Then
            Set GetSheetByTag = wks
            Exit Function
        End If
    Next
End Function

' -----------------------------------------------------------------------------------
' Function  : SheetWithTagExists
' Purpose   : Check if at least one sheet in the workbook has the given tag.
'
' Parameters:
'   wb         [Workbook] - Source workbook.
'   strTagText [String]   - Tag text to search for.
'   r_strValue [String]   - (Optional, ByRef) Returns tag value from the first match.
'
' Returns   : Boolean - True if a matching sheet exists; otherwise False.
'
' Notes     : Stops at first match for performance.
' -----------------------------------------------------------------------------------
Public Function SheetWithTagExists(ByVal wb As Workbook, ByVal strTagText As String, Optional ByRef r_strValue As String = vbNullString) As Boolean
    Dim wks As Worksheet
    For Each wks In wb.Worksheets
        If HasSheetTag(wks, strTagText, r_strValue) Then
            SheetWithTagExists = True
            Exit Function
        End If
    Next
End Function

' -----------------------------------------------------------------------------------
' Procedure : TagSelectedSheets
' Purpose   : Convenience helper to tag all currently selected worksheets.
' Parameters:
'   strTagText [String] - Name of Tag to add to each selected worksheet.
'   strValue   [String] - (Optional) Value of Tag
' Returns   : (none)
' Notes     : Non-worksheet objects in the selection are ignored.
' -----------------------------------------------------------------------------------
Public Sub TagSelectedSheets(ByVal strTagText As String, Optional strValue As String = vbNullString)
    Dim objSel As Object
    If Not ActiveWindow Is Nothing Then
        For Each objSel In ActiveWindow.SelectedSheets
            If TypeOf objSel Is Worksheet Then TagSheet objSel, strTagText, strValue
        Next
    End If
End Sub

' -----------------------------------------------------------------------------------
' Procedure : UntagSelectedSheets
' Purpose   : Convenience helper to remove a tag from all selected worksheets.
' Parameters:
'   strTagText [String] - Tag text to remove from each selected worksheet.
' Returns   : (none)
' Notes     : Non-worksheet objects in the selection are ignored.
' -----------------------------------------------------------------------------------
Public Sub UntagSelectedSheets(ByVal strTagText As String)
    Dim objSel As Object
    If Not ActiveWindow Is Nothing Then
        For Each objSel In ActiveWindow.SelectedSheets
            If TypeOf objSel Is Worksheet Then UntagSheet objSel, strTagText
        Next
    End If
End Sub

' ===== Public API - Workbook Tags ===================================================

' -----------------------------------------------------------------------------------
' Procedure : TagWorkbook
' Purpose   : Add a tag to the workbook (CustomDocumentProperty).
'
' Parameters:
'   wb         [Workbook] - Target workbook.
'   strTagText [String]   - Tag text (free-text). Name is sanitized; value is stored.
'   strValue   [String]   - (Optional) Value to store with the tag.
'
' Returns   : (none)
'
' Notes     : Uses SetDocumentProperty.
' -----------------------------------------------------------------------------------
Public Sub TagWorkbook(ByVal wb As Workbook, ByVal strTagText As String, Optional ByVal strValue As String = vbNullString)
    modProps.SetDocumentProperty wb, TagPropName(strTagText), strValue
End Sub

' -----------------------------------------------------------------------------------
' Procedure : UntagWorkbook
' Purpose   : Remove a tag from the workbook (CustomDocumentProperty).
'
' Parameters:
'   wb         [Workbook] - Target workbook.
'   strTagText [String]   - Tag text to remove.
'
' Returns   : (none)
'
' Notes     : Uses DelDocumentProperty.
' -----------------------------------------------------------------------------------
Public Sub UntagWorkbook(ByVal wb As Workbook, ByVal strTagText As String)
    modProps.DelDocumentProperty wb, TagPropName(strTagText)
End Sub

' -----------------------------------------------------------------------------------
' Function  : HasWorkbookTag
' Purpose   : Check whether the workbook has the given tag and (optionally) return it.
'
' Parameters:
'   wb         [Workbook] - Target workbook.
'   strTagText [String]   - Tag text to check.
'   r_strValue [String]   - (Optional, ByRef) Returns the tag value if present.
'
' Returns   : Boolean - True if the tag exists on the workbook; otherwise False.
'
' Notes     : Uses DocumentPropertyExists.
' -----------------------------------------------------------------------------------
Public Function HasWorkbookTag(ByVal wb As Workbook, ByVal strTagText As String, Optional ByRef r_strValue As String = vbNullString) As Boolean
    Dim objDp As DocumentProperty
    If modProps.DocumentPropertyExists(wb, TagPropName(strTagText), objDp) Then
        HasWorkbookTag = True
        On Error Resume Next
        r_strValue = CStr(objDp.Value)
        If Err.Number <> 0 Then
          
            modErr.ReportError "HasWorkbookTag", Err.Number, Erl, caption:=modMain.AppProjectName, customMessage:="Read value from " & strTagText & " failed on " & wb.Name
            Err.Clear
            r_strValue = vbNullString
        End If
        On Error GoTo 0
    End If
End Function


