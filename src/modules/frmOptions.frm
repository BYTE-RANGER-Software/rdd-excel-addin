VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOptions 
   Caption         =   "Settings %1"
   ClientHeight    =   7414
   ClientLeft      =   121
   ClientTop       =   462
   ClientWidth     =   6963
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

' -----------------------------------------------------------------------------------
' Procedure : Init (Friend)
' Purpose   : Initializes the Options dialog with project name and current settings.
'             Loads options DTO into form controls.
'
' Parameters:
'   strProjectName [String]   - Project name for caption templating
'   optionsDto     [tOptions] - Current options data transfer object (ByRef)
'
' Returns   : (none)
'
' Notes     :
'   - Friend scope; called by modMain.ShowOptionsDialog
'   - Guards against multiple initialization with m_blnInit flag
'   - Centers form to Excel window
'   - Populates txtManualPath from optionsDto.manualPath
'   - Configures MultiPage control starting at page index 0
'   - Sets m_blnConfirmed = False (assumes cancel by default)
' -----------------------------------------------------------------------------------
Friend Sub Init(ByVal strProjectName As String, ByRef optionsDto As tOptions)
    If m_blnInit Then Exit Sub
    
    'adjust Form
    CenterToExcelWindow

    Me.caption = Replace$(Me.caption, "%1", strProjectName)
        
    ' Load DTO into controls
    m_options = optionsDto
    Me.txtManualPath.text = m_options.manualPath
        
    'Purpose: mark start page & remember page index
    Const lngStartIndx As Long = 0
    With Me.mpgOptions
        .Pages(lngStartIndx).caption = ChkMark & .Pages(lngStartIndx).caption
        .Tag = lngStartIndx
        .value = lngStartIndx
    End With

    m_blnInit = True
    
End Sub

' --- Public API for the dialog ---
Public Property Get Confirmed() As Boolean
    Confirmed = m_blnConfirmed
End Property

Public Property Get ResultOptions() As tOptions
    ResultOptions = m_options
End Property

Function ChkMark() As String
    'Purpose: return ballot box with check + blank space
    ChkMark = ChrW(&H26AB) & ChrW(&HA0)  ' ballot box with check + blank
End Function

Private Sub cmdCancel_Click()
    m_blnConfirmed = False
    Unload Me
End Sub

Private Sub cmdConfirm_Click()
    Call SaveSettings
    m_blnConfirmed = True
    Unload Me
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

Function oldPage(mp As MSForms.MultiPage) As MSForms.Page
    'Purpose: return currently marked page in given multipage
    With mp
        Set oldPage = .Pages(val(.Tag))
    End With
End Function

Private Sub SaveSettings()

    m_options.manualPath = Me.txtManualPath.text
End Sub

Private Sub UserForm_Layout()
    Me.Move Application.Left + Application.Width / 2 - Me.Width / 2, Application.Top + Application.Height / 2 - Me.Height / 2
End Sub

Private Sub CenterToExcelWindow()
    ' Centers the form relative to the Excel application window.
    ' Works even if StartUpPosition is 0 (manual).
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (Application.Width - Me.Width) / 2
    Me.Top = Application.Top + (Application.Height - Me.Height) / 2
End Sub
