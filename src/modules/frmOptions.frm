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
Private m_udtOptions As tOptions

' Initialize with project name (for caption templating) and current options DTO
Friend Sub Init(ByVal strProjectName As String, ByRef udt As tOptions)
    If m_blnInit Then Exit Sub
    
        'adjust Form
        CenterToExcelWindow

        Me.Caption = Replace$(Me.Caption, "%1", strProjectName)
        
        ' Load DTO into controls
        m_udtOptions = udt
        Me.txtManualPath.Text = m_udtOptions.strManualPath
        
        'Purpose: mark start page & remember page index
        Const lngStartIndx As Long = 0
        With Me.mpgOptions
            .Pages(lngStartIndx).Caption = ChkMark & .Pages(lngStartIndx).Caption
            .Tag = lngStartIndx
            .Value = lngStartIndx
        End With

        m_blnInit = True
    
End Sub

' --- Public API for the dialog ---
Public Property Get Confirmed() As Boolean
    Confirmed = m_blnConfirmed
End Property

Public Property Get ResultOptions() As tOptions
    ResultOptions = m_udtOptions
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
        .InitialFileName = ReplaceWildcards(Me.txtManualPath.Text) & "\"
        .AllowMultiSelect = False
        .title = "Please select a folder"
        .InitialView = msoFileDialogViewList
        .ButtonName = "Select"
       
        If .Show = -1 Then
           
            Me.txtManualPath.Text = AddWildcards(.SelectedItems(1))

        End If
      
   
    End With
End Sub

Private Sub mpgOptions_Change()
    'Purpose: mark current page caption by a checkmark
    With Me.mpgOptions
        Dim pg As MSForms.Page
        'a) de-mark old caption
        Set pg = oldPage(Me.mpgOptions)
        pg.Caption = Replace(pg.Caption, ChkMark, vbNullString)
        'b) mark new caption & remember latest multipage value
        Set pg = .Pages(.Value)
        pg.Caption = ChkMark & pg.Caption
        .Tag = .Value                         ' << remember latest page index
    End With
End Sub

Function oldPage(mp As MSForms.MultiPage) As MSForms.Page
    'Purpose: return currently marked page in given multipage
    With mp
        Set oldPage = .Pages(val(.Tag))
    End With
End Function

Private Sub SaveSettings()

    m_udtOptions.strManualPath = Me.txtManualPath.Text
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
