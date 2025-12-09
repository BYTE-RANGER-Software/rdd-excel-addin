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
Private Const REG_SECTION_LOGGING       As String = "Logging"

' General Section Keys
Private Const REG_MANUAL_PATH           As String = "ManualPath"

' Logging Section Keys
Private Const REG_LOG_RETENTION_DAYS    As String = "LogRetentionDays"

' ===== Document Property Constants ==================================================
Private Const PROP_PREFIX               As String = "RDD_"

' Room Defaults
Private Const PROP_DEFAULT_GAME_WIDTH   As String = PROP_PREFIX & "DefaultGameWidth"
Private Const PROP_DEFAULT_GAME_HEIGHT  As String = PROP_PREFIX & "DefaultGameHeight"
Private Const PROP_DEFAULT_BG_WIDTH     As String = PROP_PREFIX & "DefaultBGWidth"
Private Const PROP_DEFAULT_BG_HEIGHT    As String = PROP_PREFIX & "DefaultBGHeight"
Private Const PROP_DEFAULT_UI_HEIGHT    As String = PROP_PREFIX & "DefaultUIHeight"
Private Const PROP_DEFAULT_PERSPECTIVE  As String = PROP_PREFIX & "DefaultPerspective"
Private Const PROP_DEFAULT_PARALLAX     As String = PROP_PREFIX & "DefaultParallax"
Private Const PROP_DEFAULT_SCENE_MODE   As String = PROP_PREFIX & "DefaultSceneMode"

' Behavior
Private Const PROP_AUTO_SYNC_LISTS      As String = PROP_PREFIX & "AutoSyncLists"
Private Const PROP_SHOW_VALIDATION_WARN As String = PROP_PREFIX & "ShowValidationWarnings"
Private Const PROP_PROTECT_ROOM_SHEETS  As String = PROP_PREFIX & "ProtectRoomSheets"

' ===== Default Values ===============================================================
Private Const DEF_GAME_WIDTH            As Long = 320
Private Const DEF_GAME_HEIGHT           As Long = 200
Private Const DEF_BG_WIDTH              As Long = 320
Private Const DEF_BG_HEIGHT             As Long = 200
Private Const DEF_UI_HEIGHT             As Long = 40
Private Const DEF_PERSPECTIVE           As String = ""
Private Const DEF_PARALLAX              As String = "None"
Private Const DEF_SCENE_MODE            As String = ""
Private Const DEF_LOG_RETENTION_DAYS    As Long = 30
Private Const DEF_AUTO_SYNC_LISTS       As Boolean = True
Private Const DEF_SHOW_VALIDATION_WARN  As Boolean = True
Private Const DEF_PROTECT_ROOM_SHEETS   As Boolean = True

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

' ===== General Options Properties ==========================================================

' -----------------------------------------------------------------------------------
' Property  : Opt_ManualPath (Get)
' Purpose   : Returns the current manual path option.
' Returns   : String
' -----------------------------------------------------------------------------------
Public Property Get Opt_ManualPath() As String
    Opt_ManualPath = m_options.General.manualPath
End Property

' -----------------------------------------------------------------------------------
' Property  : Opt_ManualPath (Let)
' Purpose   : Sets the manual path option and marks general options as changed.
' Parameters:
'   value [String] - New path value.
' -----------------------------------------------------------------------------------
Public Property Let Opt_ManualPath(ByVal value As String)
    If value <> m_options.General.manualPath Then
        m_options.General.manualPath = value
        HasGeneralOptionsChanged = True
    End If
End Property

' --- LogRetentionDays ---
Public Property Get Opt_LogRetentionDays() As Long
    Opt_LogRetentionDays = m_options.General.logRetentionDays
End Property

Public Property Let Opt_LogRetentionDays(ByVal value As Long)
    If value <> m_options.General.logRetentionDays Then
        m_options.General.logRetentionDays = value
        HasGeneralOptionsChanged = True
    End If
End Property

' ===== Workbook Options Properties ==================================================

' --- DefaultGameWidth ---
Public Property Get Opt_DefaultGameWidth() As Long
    Opt_DefaultGameWidth = m_options.Workbook.defaultGameWidth
End Property

Public Property Let Opt_DefaultGameWidth(ByVal value As Long)
    If value <> m_options.Workbook.defaultGameWidth Then
        m_options.Workbook.defaultGameWidth = value
        HasWorkbookOptionsChanged = True
    End If
End Property

' --- DefaultGameHeight ---
Public Property Get Opt_DefaultGameHeight() As Long
    Opt_DefaultGameHeight = m_options.Workbook.defaultGameHeight
End Property

Public Property Let Opt_DefaultGameHeight(ByVal value As Long)
    If value <> m_options.Workbook.defaultGameHeight Then
        m_options.Workbook.defaultGameHeight = value
        HasWorkbookOptionsChanged = True
    End If
End Property

' --- DefaultBGWidth ---
Public Property Get Opt_DefaultBGWidth() As Long
    Opt_DefaultBGWidth = m_options.Workbook.defaultBGWidth
End Property

Public Property Let Opt_DefaultBGWidth(ByVal value As Long)
    If value <> m_options.Workbook.defaultBGWidth Then
        m_options.Workbook.defaultBGWidth = value
        HasWorkbookOptionsChanged = True
    End If
End Property

' --- DefaultBGHeight ---
Public Property Get Opt_DefaultBGHeight() As Long
    Opt_DefaultBGHeight = m_options.Workbook.defaultBGHeight
End Property

Public Property Let Opt_DefaultBGHeight(ByVal value As Long)
    If value <> m_options.Workbook.defaultBGHeight Then
        m_options.Workbook.defaultBGHeight = value
        HasWorkbookOptionsChanged = True
    End If
End Property

' --- DefaultUIHeight ---
Public Property Get Opt_DefaultUIHeight() As Long
    Opt_DefaultUIHeight = m_options.Workbook.defaultUIHeight
End Property

Public Property Let Opt_DefaultUIHeight(ByVal value As Long)
    If value <> m_options.Workbook.defaultUIHeight Then
        m_options.Workbook.defaultUIHeight = value
        HasWorkbookOptionsChanged = True
    End If
End Property

' --- DefaultPerspective ---
Public Property Get Opt_DefaultPerspective() As String
    Opt_DefaultPerspective = m_options.Workbook.defaultPerspective
End Property

Public Property Let Opt_DefaultPerspective(ByVal value As String)
    If value <> m_options.Workbook.defaultPerspective Then
        m_options.Workbook.defaultPerspective = value
        HasWorkbookOptionsChanged = True
    End If
End Property

' --- DefaultParallax ---
Public Property Get Opt_DefaultParallax() As String
    Opt_DefaultParallax = m_options.Workbook.defaultParallax
End Property

Public Property Let Opt_DefaultParallax(ByVal value As String)
    If value <> m_options.Workbook.defaultParallax Then
        m_options.Workbook.defaultParallax = value
        HasWorkbookOptionsChanged = True
    End If
End Property

' --- DefaultSceneMode ---
Public Property Get Opt_DefaultSceneMode() As String
    Opt_DefaultSceneMode = m_options.Workbook.defaultSceneMode
End Property

Public Property Let Opt_DefaultSceneMode(ByVal value As String)
    If value <> m_options.Workbook.defaultSceneMode Then
        m_options.Workbook.defaultSceneMode = value
        HasWorkbookOptionsChanged = True
    End If
End Property

' --- AutoSyncLists ---
Public Property Get Opt_AutoSyncLists() As Boolean
    Opt_AutoSyncLists = m_options.Workbook.autoSyncLists
End Property

Public Property Let Opt_AutoSyncLists(ByVal value As Boolean)
    If value <> m_options.Workbook.autoSyncLists Then
        m_options.Workbook.autoSyncLists = value
        HasWorkbookOptionsChanged = True
    End If
End Property

' --- ShowValidationWarnings ---
Public Property Get Opt_ShowValidationWarnings() As Boolean
    Opt_ShowValidationWarnings = m_options.Workbook.showValidationWarnings
End Property

Public Property Let Opt_ShowValidationWarnings(ByVal value As Boolean)
    If value <> m_options.Workbook.showValidationWarnings Then
        m_options.Workbook.showValidationWarnings = value
        HasWorkbookOptionsChanged = True
    End If
End Property

' --- ProtectRoomSheets ---
Public Property Get Opt_ProtectRoomSheets() As Boolean
    Opt_ProtectRoomSheets = m_options.Workbook.protectRoomSheets
End Property

Public Property Let Opt_ProtectRoomSheets(ByVal value As Boolean)
    If value <> m_options.Workbook.protectRoomSheets Then
        m_options.Workbook.protectRoomSheets = value
        HasWorkbookOptionsChanged = True
    End If
End Property

' ===== DTO Bulk Access ===============================================================

' -----------------------------------------------------------------------------------
' Function  : GetAllOptions
' Purpose   : Returns a snapshot DTO of all options.
' Returns   : tOptions
' -----------------------------------------------------------------------------------
Public Function GetAllOptions() As tOptions
    GetAllOptions = m_options
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

' -----------------------------------------------------------------------------------
' Function: GetDefaultOptions
' Purpose: Returns a DTO with all default values.
' Returns   : tOptions
' -----------------------------------------------------------------------------------
Public Function GetDefaultOptions() As tOptions
    Dim opts As tOptions
    
    ' General Defaults
    opts.General.manualPath = GetDefaultManualPath()
    opts.General.logRetentionDays = DEF_LOG_RETENTION_DAYS
    
    ' Workbook Defaults
    opts.Workbook.defaultGameWidth = DEF_GAME_WIDTH
    opts.Workbook.defaultGameHeight = DEF_GAME_HEIGHT
    opts.Workbook.defaultBGWidth = DEF_BG_WIDTH
    opts.Workbook.defaultBGHeight = DEF_BG_HEIGHT
    opts.Workbook.defaultUIHeight = DEF_UI_HEIGHT
    opts.Workbook.defaultPerspective = DEF_PERSPECTIVE
    opts.Workbook.defaultParallax = DEF_PARALLAX
    opts.Workbook.defaultSceneMode = DEF_SCENE_MODE
    opts.Workbook.autoSyncLists = DEF_AUTO_SYNC_LISTS
    opts.Workbook.showValidationWarnings = DEF_SHOW_VALIDATION_WARN
    opts.Workbook.protectRoomSheets = DEF_PROTECT_ROOM_SHEETS
    
    GetDefaultOptions = opts
End Function

' ===== Persistence: Read =============================================================

' -----------------------------------------------------------------------------------
' Procedure : ReadGeneralOptions
' Purpose   : Reads general settings from registry into the in-memory DTO.
' Notes     : Resets the general-changed flag after loading.
' -----------------------------------------------------------------------------------
Public Sub ReadGeneralOptions()
On Error GoTo ErrHandler

    Dim defaults As tOptions
    defaults = GetDefaultOptions()
    
    ' General Section
    m_options.General.manualPath = GetSetting(REG_APP_NAME, REG_SECTION_GENERAL, _
        REG_MANUAL_PATH, defaults.General.manualPath)
    
    ' Logging Section
    m_options.General.logRetentionDays = CLng(GetSetting(REG_APP_NAME, REG_SECTION_LOGGING, _
        REG_LOG_RETENTION_DAYS, CStr(defaults.General.logRetentionDays)))
    
    HasGeneralOptionsChanged = False
    
Exit Sub
    
ErrHandler:
    modErr.ReportError "ReadGeneralOptions", Err.Number, Erl, caption:=modMain.AppProjectName
End Sub

' -----------------------------------------------------------------------------------
' Procedure : ReadWorkbookOptions
' Purpose   : Reads workbook-level settings (custom document properties) into DTO.
' Parameters:
'   sourceWorkbook [Workbook] - Source workbook.
' Notes     : Extend with concrete property names as needed; clears changed flag.
' -----------------------------------------------------------------------------------
Public Sub ReadWorkbookOptions(ByVal sourceWorkbook As Workbook)
On Error GoTo ErrHandler

    Dim defaults As tOptions
    defaults = GetDefaultOptions()
    
    ' Room Defaults
    m_options.Workbook.defaultGameWidth = CLng(modProps.GetDocumentPropertyValue(sourceWorkbook, _
        PROP_DEFAULT_GAME_WIDTH, CStr(defaults.Workbook.defaultGameWidth)))
    m_options.Workbook.defaultGameHeight = CLng(modProps.GetDocumentPropertyValue(sourceWorkbook, _
        PROP_DEFAULT_GAME_HEIGHT, CStr(defaults.Workbook.defaultGameHeight)))
    m_options.Workbook.defaultBGWidth = CLng(modProps.GetDocumentPropertyValue(sourceWorkbook, _
        PROP_DEFAULT_BG_WIDTH, CStr(defaults.Workbook.defaultBGWidth)))
    m_options.Workbook.defaultBGHeight = CLng(modProps.GetDocumentPropertyValue(sourceWorkbook, _
        PROP_DEFAULT_BG_HEIGHT, CStr(defaults.Workbook.defaultBGHeight)))
    m_options.Workbook.defaultUIHeight = CLng(modProps.GetDocumentPropertyValue(sourceWorkbook, _
        PROP_DEFAULT_UI_HEIGHT, CStr(defaults.Workbook.defaultUIHeight)))
    m_options.Workbook.defaultPerspective = modProps.GetDocumentPropertyValue(sourceWorkbook, _
        PROP_DEFAULT_PERSPECTIVE, defaults.Workbook.defaultPerspective)
    m_options.Workbook.defaultParallax = modProps.GetDocumentPropertyValue(sourceWorkbook, _
        PROP_DEFAULT_PARALLAX, defaults.Workbook.defaultParallax)
    m_options.Workbook.defaultSceneMode = modProps.GetDocumentPropertyValue(sourceWorkbook, _
        PROP_DEFAULT_SCENE_MODE, defaults.Workbook.defaultSceneMode)
    
    ' Behavior
    m_options.Workbook.autoSyncLists = CBool(modProps.GetDocumentPropertyValue(sourceWorkbook, _
        PROP_AUTO_SYNC_LISTS, CStr(defaults.Workbook.autoSyncLists)))
    m_options.Workbook.showValidationWarnings = CBool(modProps.GetDocumentPropertyValue(sourceWorkbook, _
        PROP_SHOW_VALIDATION_WARN, CStr(defaults.Workbook.showValidationWarnings)))
    m_options.Workbook.protectRoomSheets = CBool(modProps.GetDocumentPropertyValue(sourceWorkbook, _
        PROP_PROTECT_ROOM_SHEETS, CStr(defaults.Workbook.protectRoomSheets)))
    
    HasWorkbookOptionsChanged = False
    
Exit Sub
    
ErrHandler:
    modErr.ReportError "ReadWorkbookOptions", Err.Number, Erl, caption:=modMain.AppProjectName
End Sub

' ===== Persistence: Save =============================================================

' -----------------------------------------------------------------------------------
' Procedure : SaveGeneralOptions
' Purpose   : Writes general settings to the registry when changed (or forced).
' Parameters:
'   ignoreChangeFlag [Boolean] - True to persist regardless of change flag.
' -----------------------------------------------------------------------------------
Public Sub SaveGeneralOptions(Optional ByVal ignoreChangeFlag As Boolean = False)
On Error GoTo ErrHandler

    If HasGeneralOptionsChanged Or ignoreChangeFlag Then
        ' General Section
        SaveSetting REG_APP_NAME, REG_SECTION_GENERAL, REG_MANUAL_PATH, _
            m_options.General.manualPath
        
        ' Logging Section
        SaveSetting REG_APP_NAME, REG_SECTION_LOGGING, REG_LOG_RETENTION_DAYS, _
            CStr(m_options.General.logRetentionDays)
        
        HasGeneralOptionsChanged = False
    End If
    
Exit Sub
    
ErrHandler:
    modErr.ReportError "SaveGeneralOptions", Err.Number, Erl, caption:=modMain.AppProjectName
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
On Error GoTo ErrHandler

    If HasWorkbookOptionsChanged Or ignoreChangeFlag Then
        ' Room Defaults
        modProps.SetDocumentPropertyValue targetWorkbook, PROP_DEFAULT_GAME_WIDTH, _
            CStr(m_options.Workbook.defaultGameWidth)
        modProps.SetDocumentPropertyValue targetWorkbook, PROP_DEFAULT_GAME_HEIGHT, _
            CStr(m_options.Workbook.defaultGameHeight)
        modProps.SetDocumentPropertyValue targetWorkbook, PROP_DEFAULT_BG_WIDTH, _
            CStr(m_options.Workbook.defaultBGWidth)
        modProps.SetDocumentPropertyValue targetWorkbook, PROP_DEFAULT_BG_HEIGHT, _
            CStr(m_options.Workbook.defaultBGHeight)
        modProps.SetDocumentPropertyValue targetWorkbook, PROP_DEFAULT_UI_HEIGHT, _
            CStr(m_options.Workbook.defaultUIHeight)
        modProps.SetDocumentPropertyValue targetWorkbook, PROP_DEFAULT_PERSPECTIVE, _
            m_options.Workbook.defaultPerspective
        modProps.SetDocumentPropertyValue targetWorkbook, PROP_DEFAULT_PARALLAX, _
            m_options.Workbook.defaultParallax
        modProps.SetDocumentPropertyValue targetWorkbook, PROP_DEFAULT_SCENE_MODE, _
            m_options.Workbook.defaultSceneMode
        
        ' Behavior
        modProps.SetDocumentPropertyValue targetWorkbook, PROP_AUTO_SYNC_LISTS, _
            CStr(m_options.Workbook.autoSyncLists)
        modProps.SetDocumentPropertyValue targetWorkbook, PROP_SHOW_VALIDATION_WARN, _
            CStr(m_options.Workbook.showValidationWarnings)
        modProps.SetDocumentPropertyValue targetWorkbook, PROP_PROTECT_ROOM_SHEETS, _
            CStr(m_options.Workbook.protectRoomSheets)
        
        HasWorkbookOptionsChanged = False
    End If
    
Exit Sub
    
ErrHandler:
    modErr.ReportError "SaveWorkbookOptions", Err.Number, Erl, caption:=modMain.AppProjectName
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
    Dim resolvedPath As String
    
    ' Validate Manual Path (wenn gesetzt)
    If LenB(optionsDto.General.manualPath) > 0 Then
        resolvedPath = ReplaceWildcards(optionsDto.General.manualPath)
        If Dir$(resolvedPath, vbDirectory) = vbNullString Then
            errMsg = "Handbuch-Pfad existiert nicht: " & resolvedPath
            GoTo ValidationEnd
        End If
    End If
    
    ' Validate Log Retention
    If optionsDto.General.logRetentionDays < 0 Then
        errMsg = "Log-Aufbewahrungsdauer darf nicht negativ sein."
        GoTo ValidationEnd
    End If
    
    ' Validate Dimensions
    If optionsDto.Workbook.defaultGameWidth <= 0 Then
        errMsg = "Game-Breite muss größer als 0 sein."
        GoTo ValidationEnd
    End If
    If optionsDto.Workbook.defaultGameHeight <= 0 Then
        errMsg = "Game-Höhe muss größer als 0 sein."
        GoTo ValidationEnd
    End If
    If optionsDto.Workbook.defaultBGWidth <= 0 Then
        errMsg = "Hintergrund-Breite muss größer als 0 sein."
        GoTo ValidationEnd
    End If
    If optionsDto.Workbook.defaultBGHeight <= 0 Then
        errMsg = "Hintergrund-Höhe muss größer als 0 sein."
        GoTo ValidationEnd
    End If
    If optionsDto.Workbook.defaultUIHeight < 0 Then
        errMsg = "UI-Höhe darf nicht negativ sein."
        GoTo ValidationEnd
    End If

ValidationEnd:
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
