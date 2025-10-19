Attribute VB_Name = "modTags"
' ===================================================================================
' Module    : modTags
' Purpose   : Provide a small tagging system for worksheets and workbook using CustomProperties.
'             Tags are stored as worksheet-level CustomProperties or workbook-level CustomDocumentPropteries whose names are
'             normalized (safe) and prefixed with a constant to avoid collisions.
''
' Notes     :
'   - The tag value stored in the CustomProperty remains the original (free-text)
'     input; only the property name is sanitized for safety.
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
' Purpose   : Build the CustomProperty name used to store a tag on a sheet.
' Parameters:
'   strTagText [String] - Original (free-text) tag input.
' Returns   : String - Property name built as TAG_PREFIX & SanitizeTag(strTagText).
' Notes     : Encapsulates the naming; keeps TAG_PREFIX internal to this module.
' -----------------------------------------------------------------------------------
Private Function TagPropName(ByVal strTagText As String) As String
    TagPropName = TAG_PREFIX & SanitizeTag(strTagText)
End Function

' ===== Public API ==================================================================

' -----------------------------------------------------------------------------------
' Procedure : TagSheet
' Purpose   : Add a tag to a worksheet, stored as a CustomProperty.
' Parameters:
'   wks        [Worksheet] - Target worksheet.
'   strTagText [String]    - Tag text (free-text). Stored as value; name is normalized.
' Returns   : (none)
' Notes     : Requires SetCustomProperty(ws, name, value).
' -----------------------------------------------------------------------------------
Public Sub TagSheet(ByVal wks As Worksheet, ByVal strTagText As String)
    Dim strProp As String
    strProp = TagPropName(strTagText)
    ' Store original tag text as value; name is normalized for safety
    SetCustomProperty wks, strProp, strTagText
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
    Dim objCp As CustomProperty
    Dim strProp As String
    strProp = TagPropName(strTagText)
    If modProps.CustomPropertyExists(wks, strProp, objCp) Then
        objCp.Delete
    End If
End Sub

' -----------------------------------------------------------------------------------
' Function  : HasTag
' Purpose   : Check whether a worksheet carries the given tag.
' Parameters:
'   wks        [Worksheet] - Target worksheet.
'   strTagText [String]    - Tag text to check.
' Returns   : Boolean - True if the tag exists on the worksheet; otherwise False.
' Notes     : Uses CustomPropertyExists to probe presence.
' -----------------------------------------------------------------------------------
Public Function HasSheetTag(ByVal wks As Worksheet, ByVal strTagText As String) As Boolean
    Dim objCp As CustomProperty
    HasSheetTag = CustomPropertyExists(wks, TagPropName(strTagText), objCp)
End Function

' -----------------------------------------------------------------------------------
' Function  : GetSheetsByTag
' Purpose   : Collect all worksheets in a workbook that carry the given tag.
' Parameters:
'   wb         [Workbook] - Source workbook.
'   strTagText [String]   - Tag text to search for.
' Returns   : Collection - Collection of Worksheet objects (may be empty).
' Notes     : Uses GetAllSheetsNamesByCustomProperty to find matching sheet names.
' -----------------------------------------------------------------------------------
Public Function GetSheetsByTag(ByVal wb As Workbook, ByVal strTagText As String) As Collection
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
    Set GetSheetsByTag = colSheets
End Function

' Returns True if at least one worksheet in wb has the given tag
Public Function SheetWithTagExists(ByVal wb As Workbook, ByVal strTagText As String) As Boolean
    Dim wks As Worksheet
    For Each wks In wb.Worksheets
        If HasSheetTag(wks, strTagText) Then
            SheetWithTagExists = True
            Exit Function
        End If
    Next
End Function

' -----------------------------------------------------------------------------------
' Procedure : TagSelectedSheets
' Purpose   : Convenience helper to tag all currently selected worksheets.
' Parameters:
'   strTagText [String] - Tag text to add to each selected worksheet.
' Returns   : (none)
' Notes     : Non-worksheet objects in the selection are ignored.
' -----------------------------------------------------------------------------------
Public Sub TagSelectedSheets(ByVal strTagText As String)
    Dim objSel As Object
    For Each objSel In ActiveWindow.SelectedSheets
        If TypeOf objSel Is Worksheet Then TagSheet objSel, strTagText
    Next
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
    For Each objSel In ActiveWindow.SelectedSheets
        If TypeOf objSel Is Worksheet Then UntagSheet objSel, strTagText
    Next
End Sub

Public Sub TagWorkbook(ByVal wb As Workbook, ByVal strTagText As String)
    modProps.SetDocumentProperty wb, TagPropName(strTagText), strTagText
End Sub

Public Sub UntagWorkbook(ByVal wb As Workbook, ByVal strTagText As String)
    modProps.DelDocumentProperty wb, TagPropName(strTagText)
End Sub

Public Function HasWorkbookTag(ByVal wb As Workbook, ByVal strTagText As String) As Boolean
    HasWorkbookTag = modProps.DocumentPropertyExists(wb, TagPropName(strTagText))
End Function

