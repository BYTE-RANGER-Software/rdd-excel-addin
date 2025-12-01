Attribute VB_Name = "modProps"
' -----------------------------------------------------------------------------------
' Module    : modProps
' Purpose   : Provide helpers to read/write workbook and worksheet CustomProperties
'             and DocumentProperties, plus simple search utilities.
'
' Notes     :
'   - This module is a low-level utility; avoid UI and orchestration logic here.
'   - All helpers are designed to fail safely and be reusable across modules.
' -----------------------------------------------------------------------------------
Option Explicit
Option Private Module

' ===== Public API : Worksheet CustomProperties ======================================

' -----------------------------------------------------------------------------------
' Function  : CustomPropertyExists
' Purpose   : Returns True if a worksheet-level CustomProperty with the given name
'             exists. Optionally returns the property object via r_cp.
'
' Parameters:
'   targetSheet           [Worksheet]       - Target worksheet.
'   propertyName          [String]        - Property name to look for.
'   returnProperty        [CustomProperty] - (Optional, ByRef) returns the found property object.
'
' Returns   : Boolean - True if found; otherwise False.
'
' Notes     : Case-insensitive name comparison; safe error handling.
' -----------------------------------------------------------------------------------
Public Function CustomPropertyExists( _
    ByVal targetSheet As Worksheet, _
    ByVal propertyName As String, _
    Optional ByRef returnProperty As customProperty) As Boolean
    
    Dim cp As customProperty
    On Error GoTo ErrHandler
    Err.Clear

    For Each cp In targetSheet.CustomProperties
        If StrComp(cp.name, propertyName, vbTextCompare) = 0 Then
            Set returnProperty = cp
            CustomPropertyExists = True
            Exit Function
        End If
    Next cp

    CustomPropertyExists = False
    Set returnProperty = Nothing
    Exit Function

ErrHandler:
    CustomPropertyExists = False
    Set returnProperty = Nothing
    Err.Clear
End Function
' -----------------------------------------------------------------------------------
' Function  : GetAllSheetsNamesByCustomProperty
' Purpose   : Collect names of all worksheets in a workbook that contain the given
'             worksheet CustomProperty.
'
' Parameters:
'   sourceWorkb                [Workbook] - Source workbook.
'   returnSheetNamesArray()    [String]   - (ByRef) output array of worksheet names.
'   propertyName               [String]   - Worksheet CustomProperty name to search.
'
' Returns   : Boolean - True if at least one sheet matched; otherwise False.
'
' Notes     :
'   - The output array is erased if no matches are found or an error occurs.
' -----------------------------------------------------------------------------------
Public Function GetAllSheetsNamesByCustomProperty( _
    ByVal sourceWorkb As Workbook, _
    ByRef returnSheetNamesArray() As String, _
    ByVal propertyName As String) As Boolean

    Dim matchCount As Long
    Dim sheet As Worksheet
    Dim cp As customProperty
    
    On Error GoTo ErrHandler

    ReDim returnSheetNamesArray(sourceWorkb.Worksheets.Count - 1)
    matchCount = -1

    For Each sheet In sourceWorkb.Worksheets
        If CustomPropertyExists(sheet, propertyName, cp) Then
            matchCount = matchCount + 1
            returnSheetNamesArray(matchCount) = sheet.name
        End If
    Next sheet

    If matchCount >= 0 Then
        ReDim Preserve returnSheetNamesArray(matchCount)
        GetAllSheetsNamesByCustomProperty = True
    Else
        Erase returnSheetNamesArray
    End If

    Exit Function
    
ErrHandler:
    Erase returnSheetNamesArray
    GetAllSheetsNamesByCustomProperty = False
    Err.Clear
End Function

' -----------------------------------------------------------------------------------
' Function  : GetSheetByCustomProperty
' Purpose   : Find the first worksheet that contains the specified worksheet
'             CustomProperty, optionally with a specific value.
'
' Parameters:
'   sourceWorkbook     [Workbook] - Source workbook to scan.
'   propertyName       [String]   - Worksheet CustomProperty name to match.
'   propertyValue      [String]   - (Optional) Expected value; if empty, only presence
'                                   of the property is required.
'
' Returns   : Worksheet - The first worksheet that matches; Nothing if not found.
'
' Notes     :
'   - Value comparison uses case-insensitive StrComp like the name check.
' -----------------------------------------------------------------------------------
Public Function GetSheetByCustomProperty( _
    ByVal sourceWorkbook As Workbook, _
    ByVal propertyName As String, _
    Optional ByVal propertyValue As String = "") As Worksheet

    Dim sourceSheet As Worksheet
    Dim cp As customProperty

    If Not sourceWorkbook Is Nothing Then
        For Each sourceSheet In sourceWorkbook.Worksheets
            For Each cp In sourceSheet.CustomProperties
                If StrComp(cp.name, propertyName, vbTextCompare) <> 0 Then
                    ' continue looping
                ElseIf StrComp(cp.value, propertyValue, vbTextCompare) = 0 Then
                    Set GetSheetByCustomProperty = sourceSheet
                    Exit Function
                ElseIf StrComp(cp.name, propertyName, vbTextCompare) = 0 _
                       And propertyValue = "" Then
                    Set GetSheetByCustomProperty = sourceSheet
                    Exit Function
                End If
            Next cp
        Next sourceSheet
    End If

    Set GetSheetByCustomProperty = Nothing
End Function


' -----------------------------------------------------------------------------------
' Procedure : SetCustomProperty
' Purpose   : Create or update a worksheet-level CustomProperty with a given value.
'
' Parameters:
'   targetsheet         [Worksheet] - Target worksheet.
'   propertyName        [String]    - CustomProperty name.
'   propertyValue       [String]    - Value to set.
'
' Returns   : (none)
'
' Notes     : Adds the property if it does not exist; updates otherwise.
' -----------------------------------------------------------------------------------
Public Sub SetCustomProperty(ByVal targetSheet As Worksheet, ByVal propertyName As String, ByVal propertyValue As String)
    Dim cp As customProperty
    If CustomPropertyExists(targetSheet, propertyName, cp) Then
        cp.value = CStr(propertyValue)
    Else
        targetSheet.CustomProperties.Add name:=propertyName, value:=IIf(LenB(propertyValue) = 0, "-", propertyValue)
    End If
End Sub

' -----------------------------------------------------------------------------------
' Procedure : ClearAllCustomProperties
' Purpose   : Delete all worksheet-level CustomProperties from the given worksheet.
'
' Parameters:
'   targetSheet [Worksheet] - Target worksheet.
'
' Returns   : (none)
'
' Notes     : Uses a simple loop deleting the first item until none remain.
' -----------------------------------------------------------------------------------
Public Sub ClearAllCustomProperties(ByVal targetSheet As Worksheet)

    On Error Resume Next
    Do While targetSheet.CustomProperties.Count > 0
        targetSheet.CustomProperties(1).Delete
    Loop
End Sub

' -----------------------------------------------------------------------------------
' Function  : GetCustomPropertyValue
' Purpose   : Read the value of a worksheet-level CustomProperty by name.
'
' Parameters:
'   sourceSheet         [Worksheet] - source worksheet.
'   propertyName        [String]    - CustomProperty name.
'
' Returns   : String - The property value if found; otherwise empty string.
'
' Notes     : Returns empty string on not found or when worksheet is Nothing.
' -----------------------------------------------------------------------------------
Public Function GetCustomPropertyValue(ByVal sourceSheet As Worksheet, ByVal propertyName As String) As String
    Dim cp As customProperty
    
    On Error GoTo ErrHandler

    If Not sourceSheet Is Nothing Then
        For Each cp In sourceSheet.CustomProperties
            If StrComp(cp.name, propertyName, vbTextCompare) = 0 Then
                GetCustomPropertyValue = cp.value
                Exit Function
            End If
        Next
    End If
    
ErrHandler:
    GetCustomPropertyValue = vbNullString
    Err.Clear
End Function

' ===== Public API : Workbook DocumentProperties ====================================

' -----------------------------------------------------------------------------------
' Function  : DocumentPropertyExists
' Purpose   : Check if a workbook-level CustomDocumentProperty exists.
'
' Parameters:
'   targetWorkbook   [Workbook]        - Target workbook.
'   propertyName     [String]          - Document property name to search.
'   returnProperty   [DocumentProperty]- (Optional, ByRef) returns the found property.
'
' Returns   : Boolean - True if the property exists; otherwise False.
'
' Notes     : Case-insensitive name comparison. Safe on errors.
' -----------------------------------------------------------------------------------
Public Function DocumentPropertyExists(ByVal targetWorkbook As Workbook, ByVal propertyName As String, Optional ByRef returnProperty As DocumentProperty = Nothing) As Boolean
    Dim dp As DocumentProperty
    On Error GoTo ErrHandler
    Err.Clear
  
   
    For Each dp In targetWorkbook.CustomDocumentProperties
        If StrComp(dp.name, propertyName, vbTextCompare) = 0 Then
            Set returnProperty = dp
            DocumentPropertyExists = True
            Exit Function
        End If
    Next dp

    DocumentPropertyExists = False
    Set returnProperty = Nothing
    Exit Function

ErrHandler:
    DocumentPropertyExists = False
    Set returnProperty = Nothing
    Err.Clear
End Function

' -----------------------------------------------------------------------------------
' Function  : GetDocumentPropertyValue
' Purpose   : Read a workbook-level CustomDocumentProperty value with default fallback.
'
' Parameters:
'   sourceWorkbook         [Workbook] - source workbook.
'   propertyName[String]   - Document property name to read.
'   defaultValue [Variant]  - Default value if the property does not exist.
'
' Returns   : Variant - The property value or the provided default.
'
' Notes     : Uses DocumentPropertyExists to probe presence.
' -----------------------------------------------------------------------------------
Public Function GetDocumentPropertyValue(ByVal sourceWorkbook As Workbook, ByVal propertyName As String, ByVal defaultValue As Variant) As Variant
    Dim dp As DocumentProperty

    If DocumentPropertyExists(sourceWorkbook, propertyName, dp) Then
        GetDocumentPropertyValue = dp.value
    Else
        GetDocumentPropertyValue = defaultValue
    End If
            
End Function

' -----------------------------------------------------------------------------------
' Function  : GetDocumentProperty
' Purpose   : Retrieve a workbook-level CustomDocumentProperty object by name with
'             an optional prefix.
'
' Parameters:
'   sourceWorkbook         [Workbook]       - source workbook.
'   propertyName           [String]         - Base property name (without prefix).
'
' Returns   : DocumentProperty - The property object if found; otherwise Nothing.
'
' Notes     : Wrapper over DocumentPropertyExists.
' -----------------------------------------------------------------------------------
Public Function GetDocumentProperty(ByVal sourceWorkbook As Workbook, ByVal propertyName As String) As DocumentProperty
    Dim dp As DocumentProperty

    If DocumentPropertyExists(sourceWorkbook, propertyName, dp) Then
        Set GetDocumentProperty = dp
    End If
            
End Function

' -----------------------------------------------------------------------------------
' Procedure : SetDocumentProperty
' Purpose   : Create or update a workbook-level CustomDocumentProperty with the
'             correct MsoDocProperties type inferred from the Variant value.
'
' Parameters:
'   targetWorkbook          [Workbook]       - Target workbook.
'   propertyName            [String]         - Base property name (without prefix).
'   propertyValue           [Variant]        - Value to set.
'
' Returns   : (none)
'
' Notes     : If the property exists, its value is updated; otherwise the property
'             is created using the inferred type.
' -----------------------------------------------------------------------------------
Public Sub SetDocumentProperty(ByVal targetWorkbook As Workbook, ByVal propertyName As String, ByVal propertyValue As Variant)
    Dim dp As DocumentProperty
    Dim propertyType As MsoDocProperties

    Select Case VarType(propertyValue)
        Case 2, 3, 10, 17, 20
            propertyType = msoPropertyTypeNumber
        Case 4, 5, 6, 14
            propertyType = msoPropertyTypeFloat
        Case 7
            propertyType = msoPropertyTypeDate
        Case 8
            propertyType = msoPropertyTypeString
        Case 11
            propertyType = msoPropertyTypeBoolean
        Case Else
            Exit Sub
    End Select
    
    
    If DocumentPropertyExists(targetWorkbook, propertyName, dp) Then
        dp.value = propertyValue
    Else
                
        If propertyType = msoPropertyTypeString Then
            targetWorkbook.CustomDocumentProperties.Add name:=propertyName, LinkToContent:=False, Type:=propertyType, value:=CStr(propertyValue)
        ElseIf propertyType = msoPropertyTypeBoolean Then
            targetWorkbook.CustomDocumentProperties.Add name:=propertyName, LinkToContent:=False, Type:=propertyType, value:=CBool(propertyValue)
        ElseIf propertyType = msoPropertyTypeDate Then
            targetWorkbook.CustomDocumentProperties.Add name:=propertyName, LinkToContent:=False, Type:=propertyType, value:=CDate(propertyValue)
        ElseIf propertyType = msoPropertyTypeNumber Then
            targetWorkbook.CustomDocumentProperties.Add name:=propertyName, LinkToContent:=False, Type:=propertyType, value:=CLng(propertyValue)
        ElseIf propertyType = msoPropertyTypeFloat Then
            targetWorkbook.CustomDocumentProperties.Add name:=propertyName, LinkToContent:=False, Type:=propertyType, value:=CDbl(propertyValue)
        End If
    End If
End Sub

' -----------------------------------------------------------------------------------
' Procedure : DelDocumentProperty
' Purpose   : Delete a workbook-level CustomDocumentProperty when present.
'
' Parameters:
'   targetWorkbook        [Workbook] - Target workbook.
'   propName              [String]   - Property name to remove.
'
' Returns   : (none)
'
' Notes     : Safe no-op if the property does not exist.
' -----------------------------------------------------------------------------------
Public Sub DelDocumentProperty(ByVal targetWorkbook As Workbook, ByVal propName As String)
    Dim dp As DocumentProperty
    If DocumentPropertyExists(targetWorkbook, propName, dp) Then dp.Delete
End Sub


