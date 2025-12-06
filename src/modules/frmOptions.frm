VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOptions 
   Caption         =   "Settings - %1"
   ClientHeight    =   10860
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10200
   OleObjectBlob   =   "frmOptions.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_blnInit As Boolean
Private m_blnConfirmed As Boolean
Private m_options As tOptions
Private m_hasRDDWorkbook As Boolean     ' True if an RDD workbook is active

' ===== Constants ====================================================================
Private Const PAGE_GENERAL       As Long = 0
Private Const PAGE_ROOM_DEFAULTS As Long = 1
Private Const PAGE_BEHAVIOR      As Long = 2
' -----------------------------------------------------------------------------------
' Procedure : Init (Friend)
' Purpose   : Initializes the Options dialog with project name and current settings.
'             Loads options DTO into form controls.
'
' Parameters:
'   strProjectName [String]   - Project name for caption templating
'   optionsDto     [tOptions] - Current options data transfer object (ByRef)
'   hasRDDWorkbook [Boolean]  - True if an RDD workbook is active
'
' Returns   : (none)
'
' Notes     :
'   - Friend scope; called by modMain.ShowOptionsDialog
'   - Guards against multiple initialization with m_blnInit flag
'   - Centers form to Excel window
'   - Disables workbook-specific tabs when no RDD workbook is active
'   - Configures MultiPage control starting at page index 0
'   - Sets m_blnConfirmed = False (assumes cancel by default)
' -----------------------------------------------------------------------------------
Friend Sub Init(ByVal strProjectName As String, ByRef optionsDto As tOptions, _
    Optional ByVal hasRDDWorkbook As Boolean = False)
    
    If m_blnInit Then Exit Sub
    
    ' Store state
    m_options = optionsDto
    m_hasRDDWorkbook = hasRDDWorkbook
    
    'adjust Form
    CenterToExcelWindow
    Me.caption = Replace$(Me.caption, "%1", strProjectName)
        
    ' Load DTO into controls
    LoadGeneralOptions
    LoadWorkbookOptions
          
    ' Enable/Disable workbook-specific tabs
    EnableWorkbookTabs m_hasRDDWorkbook
    
    ' Configure MultiPage
    SetupMultiPage

    m_blnConfirmed = False
    m_blnInit = True
    
End Sub

' ===== Public API ===================================================================

Public Property Get Confirmed() As Boolean
    Confirmed = m_blnConfirmed
End Property

Public Property Get ResultOptions() As tOptions
    ResultOptions = m_options
End Property

' ===== Private Methods: Load ========================================================

Private Sub LoadGeneralOptions()
    ' --- Paths ---
    Me.txtManualPath.text = m_options.General.manualPath
        
    'Me.txtLogRetentionDays.text = CStr(m_options.General.logRetentionDays)
End Sub

Private Sub LoadWorkbookOptions()
    ' --- Room Defaults ---
    'Me.txtGameWidth.text = CStr(m_options.Workbook.defaultGameWidth)
    'Me.txtGameHeight.text = CStr(m_options.Workbook.defaultGameHeight)
    'Me.txtBGWidth.text = CStr(m_options.Workbook.defaultBGWidth)
    'Me.txtBGHeight.text = CStr(m_options.Workbook.defaultBGHeight)
    'Me.txtUIHeight.text = CStr(m_options.Workbook.defaultUIHeight)
    
    'Me.txtPerspective.text = m_options.Workbook.defaultPerspective
    'Me.txtParallax.text = m_options.Workbook.defaultParallax
    'Me.txtSceneMode.text = m_options.Workbook.defaultSceneMode
    
    ' --- Behavior ---
    'Me.chkAutoSyncLists.value = m_options.Workbook.autoSyncLists
    'Me.chkShowValidationWarnings.value = m_options.Workbook.showValidationWarnings
    'Me.chkProtectRoomSheets.value = m_options.Workbook.protectRoomSheets
End Sub

' ===== Private Methods: Save ========================================================

Private Function SaveSettings() As Boolean
    ' Save General Options
    m_options.General.manualPath = Trim$(Me.txtManualPath.text)
    
    'm_options.General.logRetentionDays = CLng(Val(Me.txtLogRetentionDays.text))
    
    ' Save Workbook Options (nur wenn RDD-Workbook aktiv)
    If m_hasRDDWorkbook Then
        'm_options.Workbook.defaultGameWidth = CLng(Val(Me.txtGameWidth.text))
        'm_options.Workbook.defaultGameHeight = CLng(Val(Me.txtGameHeight.text))
        'm_options.Workbook.defaultBGWidth = CLng(Val(Me.txtBGWidth.text))
        'm_options.Workbook.defaultBGHeight = CLng(Val(Me.txtBGHeight.text))
        'm_options.Workbook.defaultUIHeight = CLng(Val(Me.txtUIHeight.text))
        'm_options.Workbook.defaultPerspective = Trim$(Me.txtPerspective.text)
        'm_options.Workbook.defaultParallax = Trim$(Me.txtParallax.text)
        'm_options.Workbook.defaultSceneMode = Trim$(Me.txtSceneMode.text)
        
        'm_options.Workbook.autoSyncLists = Me.chkAutoSyncLists.value
        'm_options.Workbook.showValidationWarnings = Me.chkShowValidationWarnings.value
        'm_options.Workbook.protectRoomSheets = Me.chkProtectRoomSheets.value
    End If
    
    ' Validate
    Dim errMsg As String
    errMsg = modOptions.ValidateOptions(m_options)
    
    If LenB(errMsg) > 0 Then
        MsgBox errMsg, vbExclamation, "Validierungsfehler"
        SaveSettings = False
    Else
        SaveSettings = True
    End If
End Function

' ===== Private Methods: UI Helpers ==================================================

Private Sub SetupMultiPage()
    Const lngStartIndx As Long = PAGE_GENERAL
    With Me.mpgOptions
        .Pages(lngStartIndx).caption = ChkMark & .Pages(lngStartIndx).caption
        .Tag = lngStartIndx
        .value = lngStartIndx
    End With
End Sub

Private Sub EnableWorkbookTabs(ByVal enable As Boolean)
    ' Aktiviere/Deaktiviere Workbook-spezifische Tabs
    Me.mpgOptions.Pages(PAGE_ROOM_DEFAULTS).enabled = enable
    Me.mpgOptions.Pages(PAGE_BEHAVIOR).enabled = enable
    
    ' Visuelles Feedback
    Dim clr As Long
    clr = IIf(enable, &H80000005, &H8000000F)   ' WindowBackground oder GrayText
    
    ' Room Defaults Tab Controls
    'Me.txtGameWidth.enabled = enable
    'Me.txtGameHeight.enabled = enable
    'Me.txtBGWidth.enabled = enable
    'Me.txtBGHeight.enabled = enable
    'Me.txtUIHeight.enabled = enable
    'Me.txtPerspective.enabled = enable
    'Me.txtParallax.enabled = enable
    'Me.txtSceneMode.enabled = enable
    
    ' Behavior Tab Controls
    'Me.chkAutoSyncLists.enabled = enable
    'Me.chkShowValidationWarnings.enabled = enable
    'Me.chkProtectRoomSheets.enabled = enable
End Sub

Function ChkMark() As String
    'Purpose: return ballot box with check + blank space
    ChkMark = ChrW(&H26AB) & ChrW(&HA0)  ' ballot box with check + blank
End Function

Private Sub CenterToExcelWindow()
    ' Centers the form relative to the Excel application window.
    ' Works even if StartUpPosition is 0 (manual).
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (Application.Width - Me.Width) / 2
    Me.Top = Application.Top + (Application.Height - Me.Height) / 2
End Sub

Function oldPage(mp As MSForms.MultiPage) As MSForms.Page
    'Purpose: return currently marked page in given multipage
    With mp
        Set oldPage = .Pages(Val(.Tag))
    End With
End Function

' ===== Event Handlers ===============================================================

Private Sub cmdCancel_Click()
    m_blnConfirmed = False
    Unload Me
End Sub

Private Sub cmdConfirm_Click()
    If SaveSettings() Then
    m_blnConfirmed = True
    Unload Me
    End If
End Sub

Private Sub cmdResetDefaults_Click()
    Dim result As VbMsgBoxResult
    result = MsgBox("Alle Einstellungen auf Standardwerte zurücksetzen?", _
        vbQuestion + vbYesNo, "Standardwerte")
    
    If result = vbYes Then
        m_options = modOptions.GetDefaultOptions()
        LoadGeneralOptions
        If m_hasRDDWorkbook Then LoadWorkbookOptions
    End If
End Sub

Private Sub cmdSelectManualPath_Click()

       With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = ReplaceWildcards(Me.txtManualPath.text) & "\"
        .AllowMultiSelect = False
        .title = "Please select a folder"
        .InitialView = msoFileDialogViewList
        .ButtonName = "Select"
       
        If .Show = -1 Then
           
            Me.txtManualPath.text = AddWildcards(.SelectedItems(1))

        End If
      
   
    End With
End Sub

Private Sub mpgOptions_Change()
    'Purpose: mark current page caption by a checkmark
    With Me.mpgOptions
        Dim pg As MSForms.Page
        
        'a) de-mark old caption
        Set pg = oldPage(Me.mpgOptions)
        pg.caption = Replace(pg.caption, ChkMark, vbNullString)
        
        'b) mark new caption & remember latest multipage value
        Set pg = .Pages(.value)
        pg.caption = ChkMark & pg.caption
        .Tag = .value                         ' << remember latest page index
    End With
End Sub

Private Sub UserForm_Layout()
    Me.Move Application.Left + Application.Width / 2 - Me.Width / 2, Application.Top + Application.Height / 2 - Me.Height / 2
End Sub

' ===== Validation Helpers (TextBox Change Events) ===================================

Private Sub txtLogRetentionDays_Change()
    ' Nur Zahlen erlauben
'    Dim txt As String
'    txt = Me.txtLogRetentionDays.text
'    If Len(txt) > 0 Then
'        If Not IsNumeric(txt) Or InStr(txt, ".") > 0 Or InStr(txt, ",") > 0 Then
'            Me.txtLogRetentionDays.text = CStr(m_options.General.logRetentionDays)
'        End If
'    End If
End Sub

Private Sub txtGameWidth_Change()
    'ValidateNumericTextBox Me.txtGameWidth, m_options.Workbook.defaultGameWidth
End Sub

Private Sub txtGameHeight_Change()
    'ValidateNumericTextBox Me.txtGameHeight, m_options.Workbook.defaultGameHeight
End Sub

Private Sub txtBGWidth_Change()
    'ValidateNumericTextBox Me.txtBGWidth, m_options.Workbook.defaultBGWidth
End Sub

Private Sub txtBGHeight_Change()
    'ValidateNumericTextBox Me.txtBGHeight, m_options.Workbook.defaultBGHeight
End Sub

Private Sub txtUIHeight_Change()
    'ValidateNumericTextBox Me.txtUIHeight, m_options.Workbook.defaultUIHeight
End Sub

Private Sub ValidateNumericTextBox(ByRef tb As MSForms.TextBox, ByVal defaultVal As Long)
    Dim txt As String
    txt = tb.text
    If Len(txt) > 0 Then
        If Not IsNumeric(txt) Or InStr(txt, ".") > 0 Or InStr(txt, ",") > 0 Then
            tb.text = CStr(defaultVal)
        End If
    End If
End Sub




