Attribute VB_Name = "modOptions"
' -----------------------------------------------------------------------------------
' Module    : modOptions
' Purpose   : Manages all general add-in settings and RDD workbook settings.
'             Provides a small DTO for options, change tracking, and persistence
'             to registry (general) and document properties (workbook).
'
' Notes     :
'   - Keep only orchestration-free getters/setters and persistence here.
'   - UI and validation belong in forms or calling modules.
'   - needs module modOptionsTypes
' -----------------------------------------------------------------------------------

Option Explicit
Option Private Module

Private Enum OptionsChangedFlags
    OCF_None = 0
    OCF_General = &H1 ' Bit 0 -> 0001
    OCF_Workbook = &H2 ' Bit 1 -> 001
End Enum

Private m_options As tOptions

Private m_optionsChangedFlags As OptionsChangedFlags   ' bit flags: &H1 = General, &H2 = Workbook

' ===== Constants ===================================================================
Private Const REG_APP_NAME           As String = "RDD-AddIn"
Private Const REG_SECTION_GENERAL    As String = "General"
Private Const REG_MANUAL_PATH        As String = "ManualPath"


' ===== Change Flags API =============================================================

' -----------------------------------------------------------------------------------
' Property  : HasGeneralOptionsChanged (Get)
' Purpose   : Indicates whether general options have been changed.
' Returns   : Boolean
' -----------------------------------------------------------------------------------
Public Property Get HasGeneralOptionsChanged() As Boolean
    HasGeneralOptionsChanged = ((m_optionsChangedFlags And OCF_General) <> 0)
End Property

' -----------------------------------------------------------------------------------
' Property  : HasGeneralOptionsChanged (Let)
' Purpose   : Sets/clears the "general options changed" flag.
' Parameters:
'   Value [Boolean] - True to set, False to clear.
' -----------------------------------------------------------------------------------
Public Property Let HasGeneralOptionsChanged(ByVal value As Boolean)
    If value Then
        m_optionsChangedFlags = m_optionsChangedFlags Or OCF_General
    Else
        m_optionsChangedFlags = m_optionsChangedFlags And Not OCF_General
    End If
End Property

' -----------------------------------------------------------------------------------
' Property  : HasWorkbookOptionsChanged (Get)
' Purpose   : Indicates whether workbook options have been changed.
' Returns   : Boolean
' -----------------------------------------------------------------------------------
Public Property Get HasWorkbookOptionsChanged() As Boolean
    HasWorkbookOptionsChanged = ((m_optionsChangedFlags And OCF_Workbook) <> 0)
End Property

' -----------------------------------------------------------------------------------
' Property  : HasWorkbookOptionsChanged (Let)
' Purpose   : Sets/clears the "workbook options changed" flag.
' Parameters:
'   Value [Boolean] - True to set, False to clear.
' -----------------------------------------------------------------------------------
Public Property Let HasWorkbookOptionsChanged(ByVal value As Boolean)
    If value Then
        m_optionsChangedFlags = m_optionsChangedFlags Or OCF_Workbook
    Else
        m_optionsChangedFlags = m_optionsChangedFlags And Not OCF_Workbook
    End If
End Property

' -----------------------------------------------------------------------------------
' Property  : OptionsChanged (Get)
' Purpose   : Indicates whether any option set (general or workbook) changed.
' Returns   : Boolean
' -----------------------------------------------------------------------------------
Public Property Get HasOptionsChanged() As Boolean
    HasOptionsChanged = (m_optionsChangedFlags <> OCF_None)
End Property

' ===== General Options API ==========================================================

' -----------------------------------------------------------------------------------
' Property  : Opt_ManualPath (Get)
' Purpose   : Returns the current manual path option.
' Returns   : String
' -----------------------------------------------------------------------------------
Public Property Get Opt_ManualPath() As String
    Opt_ManualPath = m_options.manualPath
End Property

' -----------------------------------------------------------------------------------
' Property  : Opt_ManualPath (Let)
' Purpose   : Sets the manual path option and marks general options as changed.
' Parameters:
'   value [String] - New path value.
' -----------------------------------------------------------------------------------
Public Property Let Opt_ManualPath(ByVal value As String)
    If value <> m_options.manualPath Then
        m_options.manualPath = value
        HasGeneralOptionsChanged = True
    End If
End Property

' ===== DTO Bulk Access ===============================================================

' -----------------------------------------------------------------------------------
' Function  : GetAllOptions
' Purpose   : Returns a snapshot DTO of all options.
' Returns   : tOptions
' -----------------------------------------------------------------------------------
Public Function GetAllOptions() As tOptions
    Dim optionsDto As tOptions
    optionsDto = m_options
    GetAllOptions = optionsDto
End Function

' -----------------------------------------------------------------------------------
' Procedure : SetAllOptions
' Purpose   : Replaces all options from a provided DTO and marks both change flags.
' Parameters:
'   optionsDto [tOptions] - New options DTO (ByRef to avoid copying large UDTs).
' -----------------------------------------------------------------------------------
Public Sub SetAllOptions(ByRef optionsDto As tOptions)
    m_options = optionsDto
    HasGeneralOptionsChanged = True
    HasWorkbookOptionsChanged = True
End Sub

' ===== Persistence: Read =============================================================

' -----------------------------------------------------------------------------------
' Procedure : ReadGeneralOptions
' Purpose   : Reads general settings from registry into the in-memory DTO.
' Notes     : Resets the general-changed flag after loading.
' -----------------------------------------------------------------------------------
Public Sub ReadGeneralOptions()
    m_options.manualPath = GetSetting(REG_APP_NAME, REG_SECTION_GENERAL, REG_MANUAL_PATH, GetDefaultManualPath())
    HasGeneralOptionsChanged = False
End Sub

' -----------------------------------------------------------------------------------
' Procedure : ReadWorkbookOptions
' Purpose   : Reads workbook-level settings (custom document properties) into DTO.
' Parameters:
'   sourceWorkbook [Workbook] - Source workbook.
' Notes     : Extend with concrete property names as needed; clears changed flag.
' -----------------------------------------------------------------------------------
Public Sub ReadWorkbookOptions(ByVal sourceWorkbook As Workbook)
    ' Example:
    ' Opt_... = modProps.GetDocumentPropertyValue(sourceWorkbook, PROP_..., defaultValue)
    HasWorkbookOptionsChanged = False
End Sub

' ===== Persistence: Save =============================================================

' -----------------------------------------------------------------------------------
' Procedure : SaveGeneralOptions
' Purpose   : Writes general settings to the registry when changed (or forced).
' Parameters:
'   ignoreChangeFlag [Boolean] - True to persist regardless of change flag.
' -----------------------------------------------------------------------------------
Public Sub SaveGeneralOptions(Optional ByVal ignoreChangeFlag As Boolean = False)
    If HasGeneralOptionsChanged Or ignoreChangeFlag Then
        SaveSetting REG_APP_NAME, REG_SECTION_GENERAL, REG_MANUAL_PATH, m_options.manualPath
        HasGeneralOptionsChanged = False
    End If
End Sub

' -----------------------------------------------------------------------------------
' Procedure : SaveWorkbookOptions
' Purpose   : Writes workbook-level settings to custom document properties when
'             changed (or forced).
' Parameters:
'   targetWorkbook             [Workbook] - Target workbook.
'   ignoreChangeFlag [Boolean]  - True to persist regardless of change flag.
' -----------------------------------------------------------------------------------
Public Sub SaveWorkbookOptions(ByVal targetWorkbook As Workbook, Optional ByVal ignoreChangeFlag As Boolean = False)
    If HasWorkbookOptionsChanged Or ignoreChangeFlag Then
        ' Example:
        ' modProps.SetDocumentProperty targetWorkbook, PROP_..., Opt_...
        If targetWorkbook.path <> "" Then targetWorkbook.Save
        HasWorkbookOptionsChanged = False
    End If
End Sub

' ===== Validation ==================================================================

' -----------------------------------------------------------------------------------
' Function  : ValidateOptions
' Purpose   : Validates the provided options DTO. Returns an empty string when valid,
'             otherwise a descriptive error message.
'
' Parameters:
'   optionsDto [tOptions] - Options DTO to validate (ByRef; not modified).
'
' Returns   : String - "" if OK; otherwise an error message.
'
' Notes     :
'   - Manual path may be empty; if it is set, it must point to an existing folder.
' -----------------------------------------------------------------------------------
Public Function ValidateOptions(ByRef optionsDto As tOptions) As String
    Dim errMsg As String
    ' Example: Manual path may be empty, but if set, it must exist
    If LenB(optionsDto.manualPath) > 0 Then
        If Dir$(optionsDto.manualPath, vbDirectory) = vbNullString Then
            errMsg = "Manual path does not exist."
        End If
    End If
    ValidateOptions = errMsg
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
