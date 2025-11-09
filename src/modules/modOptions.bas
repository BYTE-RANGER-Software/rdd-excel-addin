Attribute VB_Name = "modOptions"
' ===================================================================================
' Module    : modOptions
' Purpose   : Manages all general add-in settings and RDD workbook settings.
'             Provides a small DTO for options, change tracking, and persistence
'             to registry (general) and document properties (workbook).
'
' Notes     :
'   - Keep only orchestration-free getters/setters and persistence here.
'   - UI and validation belong in forms or calling modules.
'   - needs module modOptionsTypes
' ===================================================================================

Option Explicit
Option Private Module

Private m_udtOptions As tOptions

Private m_lngOptionsChanged As Long   ' bit flags: &H1 = General, &H2 = Workbook

' ===== Constants ===================================================================
Private Const REG_APP_NAME           As String = "RDD-AddIn"
Private Const REG_SECTION_GENERAL    As String = "General"
Private Const REG_MANUAL_PATH        As String = "ManualPath"

Private Const OPTIONS_GENERAL_CHANGED  As Long = &H1   ' Bit 0 -> 0001
Private Const OPTIONS_WORKBOOK_CHANGED As Long = &H2   ' Bit 1 -> 0010

' ===== Change Flags API =============================================================

' -----------------------------------------------------------------------------------
' Property  : GeneralOptionsChanged (Get)
' Purpose   : Indicates whether general options have been changed.
' Returns   : Boolean
' -----------------------------------------------------------------------------------
Public Property Get GeneralOptionsChanged() As Boolean
    GeneralOptionsChanged = ((m_lngOptionsChanged And OPTIONS_GENERAL_CHANGED) <> 0)
End Property

' -----------------------------------------------------------------------------------
' Property  : GeneralOptionsChanged (Let)
' Purpose   : Sets/clears the "general options changed" flag.
' Parameters:
'   blnVal [Boolean] - True to set, False to clear.
' -----------------------------------------------------------------------------------
Public Property Let GeneralOptionsChanged(ByVal blnVal As Boolean)
    If blnVal Then
        m_lngOptionsChanged = m_lngOptionsChanged Or OPTIONS_GENERAL_CHANGED
    Else
        m_lngOptionsChanged = m_lngOptionsChanged And Not OPTIONS_GENERAL_CHANGED
    End If
End Property

' -----------------------------------------------------------------------------------
' Property  : WorkBookOptionsChanged (Get)
' Purpose   : Indicates whether workbook options have been changed.
' Returns   : Boolean
' -----------------------------------------------------------------------------------
Public Property Get WorkbookOptionsChanged() As Boolean
    WorkbookOptionsChanged = ((m_lngOptionsChanged And OPTIONS_WORKBOOK_CHANGED) <> 0)
End Property

' -----------------------------------------------------------------------------------
' Property  : WorkBookOptionsChanged (Let)
' Purpose   : Sets/clears the "workbook options changed" flag.
' Parameters:
'   blnVal [Boolean] - True to set, False to clear.
' -----------------------------------------------------------------------------------
Public Property Let WorkbookOptionsChanged(ByVal blnVal As Boolean)
    If blnVal Then
        m_lngOptionsChanged = m_lngOptionsChanged Or OPTIONS_WORKBOOK_CHANGED
    Else
        m_lngOptionsChanged = m_lngOptionsChanged And Not OPTIONS_WORKBOOK_CHANGED
    End If
End Property

' -----------------------------------------------------------------------------------
' Property  : OptionsChanged (Get)
' Purpose   : Indicates whether any option set (general or workbook) changed.
' Returns   : Boolean
' -----------------------------------------------------------------------------------
Public Property Get OptionsChanged() As Boolean
    OptionsChanged = (m_lngOptionsChanged <> 0)
End Property

' ===== General Options API ==========================================================

' -----------------------------------------------------------------------------------
' Property  : Opt_ManualPath (Get)
' Purpose   : Returns the current manual path option.
' Returns   : String
' -----------------------------------------------------------------------------------
Public Property Get Opt_ManualPath() As String
    Opt_ManualPath = m_udtOptions.strManualPath
End Property

' -----------------------------------------------------------------------------------
' Property  : Opt_ManualPath (Let)
' Purpose   : Sets the manual path option and marks general options as changed.
' Parameters:
'   value [String] - New path value.
' -----------------------------------------------------------------------------------
Public Property Let Opt_ManualPath(ByVal value As String)
    If value <> m_udtOptions.strManualPath Then
        m_udtOptions.strManualPath = value
        GeneralOptionsChanged = True
    End If
End Property

' ===== DTO Bulk Access ===============================================================

' -----------------------------------------------------------------------------------
' Function  : GetAllOptions
' Purpose   : Returns a snapshot DTO of all options.
' Returns   : tOptions
' -----------------------------------------------------------------------------------
Public Function GetAllOptions() As tOptions
    Dim udt As tOptions
    udt = m_udtOptions
    GetAllOptions = udt
End Function

' -----------------------------------------------------------------------------------
' Procedure : SetAllOptions
' Purpose   : Replaces all options from a provided DTO and marks both change flags.
' Parameters:
'   udt [tOptions] - New options DTO (ByRef to avoid copying large UDTs).
' -----------------------------------------------------------------------------------
Public Sub SetAllOptions(ByRef udt As tOptions)
    m_udtOptions = udt
    GeneralOptionsChanged = True
    WorkbookOptionsChanged = True
End Sub

' ===== Persistence: Read =============================================================

' -----------------------------------------------------------------------------------
' Procedure : ReadGeneralOptions
' Purpose   : Reads general settings from registry into the in-memory DTO.
' Notes     : Resets the general-changed flag after loading.
' -----------------------------------------------------------------------------------
Public Sub ReadGeneralOptions()
    m_udtOptions.strManualPath = GetSetting(REG_APP_NAME, REG_SECTION_GENERAL, REG_MANUAL_PATH, GetDefaultManualPath())
    GeneralOptionsChanged = False
End Sub

' -----------------------------------------------------------------------------------
' Procedure : ReadWorkbookOptions
' Purpose   : Reads workbook-level settings (custom document properties) into DTO.
' Parameters:
'   objWkBk [Workbook] - Source workbook.
' Notes     : Extend with concrete property names as needed; clears changed flag.
' -----------------------------------------------------------------------------------
Public Sub ReadWorkbookOptions(ByVal objWkBk As Workbook)
    ' Example:
    ' Opt_... = modProps.GetDocumentPropertyValue(objWkBk, PROP_..., defaultValue)
    WorkbookOptionsChanged = False
End Sub

' ===== Persistence: Save =============================================================

' -----------------------------------------------------------------------------------
' Procedure : SaveGeneralOptions
' Purpose   : Writes general settings to the registry when changed (or forced).
' Parameters:
'   blnIgnoreChangeFlag [Boolean] - True to persist regardless of change flag.
' -----------------------------------------------------------------------------------
Public Sub SaveGeneralOptions(Optional ByVal blnIgnoreChangeFlag As Boolean = False)
    If GeneralOptionsChanged Or blnIgnoreChangeFlag Then
        SaveSetting REG_APP_NAME, REG_SECTION_GENERAL, REG_MANUAL_PATH, m_udtOptions.strManualPath
        GeneralOptionsChanged = False
    End If
End Sub

' -----------------------------------------------------------------------------------
' Procedure : SaveWorkbookOptions
' Purpose   : Writes workbook-level settings to custom document properties when
'             changed (or forced).
' Parameters:
'   objWkBk             [Workbook] - Target workbook.
'   blnIgnoreChangeFlag [Boolean]  - True to persist regardless of change flag.
' -----------------------------------------------------------------------------------
Public Sub SaveWorkbookOptions(ByVal objWkBk As Workbook, Optional ByVal blnIgnoreChangeFlag As Boolean = False)
    If WorkbookOptionsChanged Or blnIgnoreChangeFlag Then
        ' Example:
        ' modProps.SetDocumentProperty objWkBk, PROP_..., Opt_...
        If objWkBk.path <> "" Then objWkBk.Save
        WorkbookOptionsChanged = False
    End If
End Sub

' ===== Validation ==================================================================

' -----------------------------------------------------------------------------------
' Function  : ValidateOptions
' Purpose   : Validates the provided options DTO. Returns an empty string when valid,
'             otherwise a descriptive error message.
'
' Parameters:
'   udtOptions [tOptions] - Options DTO to validate (ByRef; not modified).
'
' Returns   : String - "" if OK; otherwise an error message.
'
' Notes     :
'   - Manual path may be empty; if it is set, it must point to an existing folder.
' -----------------------------------------------------------------------------------
Public Function ValidateOptions(ByRef udt As tOptions) As String
    Dim strErr As String
    ' Example: Manual path may be empty, but if set, it must exist
    If LenB(udt.strManualPath) > 0 Then
        If Dir$(udt.strManualPath, vbDirectory) = vbNullString Then
            strErr = "Manual path does not exist."
        End If
    End If
    ValidateOptions = strErr
End Function

' ===== Private Helpers ===============================================================

' -----------------------------------------------------------------------------------
' Function  : GetDefaultManualPath
' Purpose   : Builds the default manual path string.
' Returns   : String - e.g., "<MyDocuments>\<Project>\Doku"
' Notes     : Uses modConst.WILDCARD_MY_DOCUMENTS and AppProjectName.
' -----------------------------------------------------------------------------------
Private Function GetDefaultManualPath() As String
    GetDefaultManualPath = modConst.WILDCARD_MY_DOCUMENTS & "\" & AppProjectName & "\Doku"
End Function
