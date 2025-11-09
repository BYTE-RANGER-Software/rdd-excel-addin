Attribute VB_Name = "modProps"
' ===================================================================================
' Module    : modProps
' Purpose   : Provide helpers to read/write workbook and worksheet CustomProperties
'             and DocumentProperties, plus simple search utilities.
'
' Notes     :
' ===================================================================================
Option Explicit
Option Private Module

' ===== Public API : Worksheet CustomProperties ======================================

' -----------------------------------------------------------------------------------
' Function  : CustomPropertyExists
' Purpose   : Returns True if a worksheet-level CustomProperty with the given name
'             exists. Optionally returns the property object via r_cp.
'
' Parameters:
'   wks           [Worksheet]       - Target worksheet.
'   strPropName   [String]          - Property name to look for.
'   r_cp          [CustomProperty]  - (Optional, ByRef) returns the found property object.
'
' Returns   : Boolean - True if found; otherwise False.
'
' Notes     : Case-insensitive name comparison; safe error handling.
' -----------------------------------------------------------------------------------
Public Function CustomPropertyExists( _
    ByVal wks As Worksheet, _
    ByVal strPropName As String, _
    Optional ByRef r_cp As CustomProperty) As Boolean
    
    Dim cp As CustomProperty
    On Error GoTo ErrHandler         ' fail-safe error handling
    Err.Clear

    For Each cp In wks.CustomProperties
        If StrComp(cp.Name, strPropName, vbTextCompare) = 0 Then
            Set r_cp = cp
            CustomPropertyExists = True
            Exit Function
        End If
    Next cp

    CustomPropertyExists = False
    Set r_cp = Nothing
    Exit Function

ErrHandler:
    CustomPropertyExists = False
    Set r_cp = Nothing
    Err.Clear
End Function

' -----------------------------------------------------------------------------------
' Function  : GetAllSheetsNamesByCustomProperty
' Purpose   : Collect names of all worksheets in a workbook that contain the given
'             worksheet CustomProperty.
'
' Parameters:
'   wb           [Workbook] - Source workbook.
'   r_astrWkShNames() [String]   - (ByRef) output array of worksheet names.
'   strPropName       [String]   - Worksheet CustomProperty name to search.
'
' Returns   : Boolean - True if at least one sheet matched; otherwise False.
'
' Notes     : The output array is erased if no matches are found.
' -----------------------------------------------------------------------------------
Public Function GetAllSheetsNamesByCustomProperty(ByVal wb As Workbook, ByRef r_astrWkShNames() As String, ByVal strPropName As String) As Boolean
    Dim lngCount As Long
    Dim wks As Worksheet
    Dim cp As CustomProperty
    
    On Error GoTo ErrHandler

    ReDim r_astrWkShNames(wb.Worksheets.Count - 1)
    lngCount = -1
    For Each wks In wb.Worksheets
        If CustomPropertyExists(wks, strPropName, cp) Then
            lngCount = lngCount + 1
            r_astrWkShNames(lngCount) = wks.Name
        End If
    Next

    If lngCount >= 0 Then
        ReDim Preserve r_astrWkShNames(lngCount)
        GetAllSheetsNamesByCustomProperty = True
    Else
        Erase r_astrWkShNames()
    End If

    Exit Function
    
ErrHandler:
    Erase r_astrWkShNames
    GetAllSheetsNamesByCustomProperty = False
    Err.Clear
End Function

' -----------------------------------------------------------------------------------
' Function  : GetSheetByCustomProperty
' Purpose   : Find the first worksheet that contains the specified worksheet
'             CustomProperty, optionally with a specific value.
'
' Parameters:
'   wb           [Workbook] - Source workbook to scan.
'   strPropName  [String]   - Worksheet CustomProperty name to match.
'   strPropValue [String]   - (Optional) Expected value; if empty, only presence
'                             of the property is required.
'
' Returns   : Worksheet - The first worksheet that matches; Nothing if not found.
'
' Notes     : Value comparison uses case-insensitive StrComp like the name check.
' -----------------------------------------------------------------------------------
Public Function GetSheetByCustomProperty(ByVal wb As Workbook, ByVal strPropName As String, Optional ByVal strPropValue As String = "") As Worksheet
    Dim wks As Worksheet
    Dim cp As CustomProperty

    If Not wb Is Nothing Then
        For Each wks In wb.Worksheets
            For Each cp In wks.CustomProperties
                If StrComp(cp.Name, strPropName, vbTextCompare) <> 0 Then
                    'Continue For
                ElseIf StrComp(cp.value, strPropValue, vbTextCompare) = 0 Then
                    Set GetSheetByCustomProperty = wks
                    Exit Function
                ElseIf StrComp(cp.Name, strPropName, vbTextCompare) = 0 And strPropValue = "" Then
                    Set GetSheetByCustomProperty = wks
                    Exit Function
                End If

            Next
        Next
    End If
    Set GetSheetByCustomProperty = Nothing
End Function

' -----------------------------------------------------------------------------------
' Procedure : SetCustomProperty
' Purpose   : Create or update a worksheet-level CustomProperty with a given value.
'
' Parameters:
'   wks         [Worksheet] - Target worksheet.
'   strPropName [String]    - CustomProperty name.
'   strPropValue[String]    - Value to set.
'
' Returns   : (none)
'
' Notes     : Adds the property if it does not exist; updates otherwise.
' -----------------------------------------------------------------------------------
Public Sub SetCustomProperty(ByVal wks As Worksheet, ByVal strPropName As String, ByVal strPropValue As String)
    Dim cp As CustomProperty
    If CustomPropertyExists(wks, strPropName, cp) Then
        cp.value = CStr(strPropValue)
    Else
        wks.CustomProperties.Add Name:=strPropName, value:=IIf(LenB(strPropValue) = 0, "-", strPropValue)
    End If
End Sub

' -----------------------------------------------------------------------------------
' Procedure : ClearAllCustomProperties
' Purpose   : Delete all worksheet-level CustomProperties from the given worksheet.
'
' Parameters:
'   wks [Worksheet] - Target worksheet.
'
' Returns   : (none)
'
' Notes     : Uses a simple loop deleting the first item until none remain.
' -----------------------------------------------------------------------------------
Public Sub ClearAllCustomProperties(ByVal wks As Worksheet)

    On Error Resume Next
    Do While wks.CustomProperties.Count > 0
        wks.CustomProperties(1).Delete
    Loop
End Sub

' -----------------------------------------------------------------------------------
' Function  : GetCustomPropertyValue
' Purpose   : Read the value of a worksheet-level CustomProperty by name.
'
' Parameters:
'   wks         [Worksheet] - Target worksheet.
'   strPropName [String]    - CustomProperty name.
'
' Returns   : String - The property value if found; otherwise empty string.
'
' Notes     : Returns empty string on not found or when worksheet is Nothing.
' -----------------------------------------------------------------------------------
Public Function GetCustomPropertyValue(ByVal wks As Worksheet, ByVal strPropName As String) As String
    Dim cp As CustomProperty
    
    On Error GoTo ErrHandler

    If Not wks Is Nothing Then
        For Each cp In wks.CustomProperties
            If StrComp(cp.Name, strPropName, vbTextCompare) = 0 Then
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
'   wb              [Workbook]        - Target workbook.
'   strPropName     [String]          - Document property name to search.
'   r_objReturnProp [DocumentProperty]- (Optional, ByRef) returns the found property.
'
' Returns   : Boolean - True if the property exists; otherwise False.
'
' Notes     : Case-insensitive name comparison. Safe on errors.
' -----------------------------------------------------------------------------------
Public Function DocumentPropertyExists(ByVal wb As Workbook, ByVal strPropName As String, Optional ByRef r_dp As DocumentProperty = Nothing) As Boolean
    Dim dp As DocumentProperty
    On Error GoTo ErrHandler         ' fail-safe error handling
    Err.Clear
  
   
    For Each dp In wb.CustomDocumentProperties
        If StrComp(dp.Name, strPropName, vbTextCompare) = 0 Then
            Set r_dp = dp
            DocumentPropertyExists = True
            Exit Function
        End If
    Next dp

    DocumentPropertyExists = False
    Set r_dp = Nothing
    Exit Function

ErrHandler:
    DocumentPropertyExists = False
    Set r_dp = Nothing
    Err.Clear
End Function

' -----------------------------------------------------------------------------------
' Function  : GetDocumentPropertyValue
' Purpose   : Read a workbook-level CustomDocumentProperty value with default fallback.
'
' Parameters:
'   wb         [Workbook] - Target workbook.
'   strPropName[String]   - Document property name to read.
'   varDefault [Variant]  - Default value if the property does not exist.
'
' Returns   : Variant - The property value or the provided default.
'
' Notes     : Uses DocumentPropertyExists to probe presence.
' -----------------------------------------------------------------------------------
Public Function GetDocumentPropertyValue(ByVal wb As Workbook, ByVal strPropName As String, ByVal varDefault As Variant) As Variant
    Dim dp As DocumentProperty

    If DocumentPropertyExists(wb, strPropName, dp) Then
        GetDocumentPropertyValue = dp.value
    Else
        GetDocumentPropertyValue = varDefault
    End If
            
End Function

' -----------------------------------------------------------------------------------
' Function  : GetDocumentProperty
' Purpose   : Retrieve a workbook-level CustomDocumentProperty object by name with
'             an optional prefix.
'
' Parameters:
'   wb         [Workbook]       - Target workbook.
'   strPropName[String]         - Base property name (without prefix).
'
' Returns   : DocumentProperty - The property object if found; otherwise Nothing.
'
' Notes     : Wrapper over DocumentPropertyExists.
' -----------------------------------------------------------------------------------
Public Function GetDocumentProperty(ByVal wb As Workbook, ByVal strPropName As String) As DocumentProperty
    Dim dp As DocumentProperty

    If DocumentPropertyExists(wb, strPropName, dp) Then
        Set GetDocumentProperty = dp
    End If
            
End Function

' -----------------------------------------------------------------------------------
' Procedure : SetDocumentProperty
' Purpose   : Create or update a workbook-level CustomDocumentProperty with the
'             correct MsoDocProperties type inferred from the Variant value.
'
' Parameters:
'   wb          [Workbook]       - Target workbook.
'   strPropName [String]         - Base property name (without prefix).
'   varPropValue[Variant]        - Value to set.
'
' Returns   : (none)
'
' Notes     : If the property exists, its value is updated; otherwise the property
'             is created using the inferred type.
' -----------------------------------------------------------------------------------
Public Sub SetDocumentProperty(ByVal wb As Workbook, ByVal strPropName As String, ByVal varPropValue As Variant)
    Dim dp As DocumentProperty
    Dim intPropType As MsoDocProperties

    Select Case VarType(varPropValue)
        Case 2, 3, 10, 17, 20
            intPropType = msoPropertyTypeNumber
        Case 4, 5, 6, 14
            intPropType = msoPropertyTypeFloat
        Case 7
            intPropType = msoPropertyTypeDate
        Case 8
            intPropType = msoPropertyTypeString
        Case 11
            intPropType = msoPropertyTypeBoolean
        Case Else
            Exit Sub
    End Select
    
    
    If DocumentPropertyExists(wb, strPropName, dp) Then
        dp.value = varPropValue
    Else
                
        If intPropType = msoPropertyTypeString Then
            wb.CustomDocumentProperties.Add Name:=strPropName, LinkToContent:=False, Type:=intPropType, value:=CStr(varPropValue)
        ElseIf intPropType = msoPropertyTypeBoolean Then
            wb.CustomDocumentProperties.Add Name:=strPropName, LinkToContent:=False, Type:=intPropType, value:=CBool(varPropValue)
        ElseIf intPropType = msoPropertyTypeDate Then
            wb.CustomDocumentProperties.Add Name:=strPropName, LinkToContent:=False, Type:=intPropType, value:=CDate(varPropValue)
        ElseIf intPropType = msoPropertyTypeNumber Then
            wb.CustomDocumentProperties.Add Name:=strPropName, LinkToContent:=False, Type:=intPropType, value:=CLng(varPropValue)
        ElseIf intPropType = msoPropertyTypeFloat Then
            wb.CustomDocumentProperties.Add Name:=strPropName, LinkToContent:=False, Type:=intPropType, value:=CDbl(varPropValue)
        End If
    End If
End Sub

' -----------------------------------------------------------------------------------
' Procedure : DelDocumentProperty
' Purpose   : Delete a workbook-level CustomDocumentProperty when present.
'
' Parameters:
'   wb        [Workbook] - Target workbook.
'   propName  [String]   - Property name to remove.
'
' Returns   : (none)
'
' Notes     : Safe no-op if the property does not exist.
' -----------------------------------------------------------------------------------
Public Sub DelDocumentProperty(ByVal wb As Workbook, ByVal propName As String)
    Dim p As DocumentProperty
    If DocumentPropertyExists(wb, propName, p) Then p.Delete
End Sub


