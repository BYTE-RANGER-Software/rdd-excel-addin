VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOptions 
   Caption         =   "Settings %1"
   ClientHeight    =   7410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6960
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
Private m_objActWkBk As Workbook

Friend Sub Init(ByVal objWkBk As Workbook)
    If Not m_blnInit Then
    
        Dim sngTop As Single, sngLeft As Single
        Dim aobjTempSheets() As Worksheet
        Dim i As Integer
   
        Set m_objActWkBk = objWkBk
        'adjust Form
        Me.StartUpPosition = 0
        sngLeft = Application.Left + Application.Width / 2 - Me.Width / 2
        sngTop = Application.Top + Application.Height / 2 - Me.Height / 2

        Me.Left = sngLeft
        Me.Top = sngTop
        Me.Caption = Replace$(Me.Caption, "%1", modMain.AppProjectName)
        
        'Load Settings
        Me.txtManualPath.Text = Opt_ManualPath
        
        'Purpose: mark start page & remember page index
        Const startIndx As Long = 0
        With Me.mpgOptions
            .Pages(startIndx).Caption = ChkMark & .Pages(startIndx).Caption
            .Tag = startIndx
        End With
        Me.mpgOptions.Value = 0

        m_blnInit = True
    End If
End Sub

Function ChkMark() As String
    'Purpose: return ballot box with check + blank space
    ChkMark = ChrW(&H26AB) & ChrW(&HA0)  ' ballot box with check + blank
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdConfirm_Click()
    Call SaveSettings
    Unload Me
End Sub

Private Sub cmdSelectManualPath_Click()

       With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = ReplaceWildCards(Me.txtManualPath.Text) & "\"
        .AllowMultiSelect = False
        .Title = "Please select a folder"
        .InitialView = msoFileDialogViewList
        .ButtonName = "Select"
       
        If .Show = -1 Then
           
            Me.txtManualPath.Text = AddWildCards(.SelectedItems(1))

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

    Call SaveGeneralSettings
    OptionsChanged = True
    If m_objActWkBk.Path <> "" Then m_objActWkBk.Save
End Sub

Private Sub SaveGeneralSettings()
    Opt_ManualPath = Me.txtManualPath.Text
End Sub

Private Sub UserForm_Layout()
    Me.Move Application.Left + Application.Width / 2 - Me.Width / 2, Application.Top + Application.Height / 2 - Me.Height / 2
End Sub

