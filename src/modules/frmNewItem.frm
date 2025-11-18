VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmNewItem 
   Caption         =   "New Item"
   ClientHeight    =   2655
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7815
   OleObjectBlob   =   "frmNewItem.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmNewItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_blnCancelled As Boolean

Private Sub UserForm_Initialize()
        
        CenterToExcelWindow

        Me.caption = "New Item"
        
        Me.lblName.caption = "Name"
        
        Me.lblID.caption = "ID"
        Me.lblID.Visible = False
        Me.txtID.Visible = False
        Me.txtID.Locked = True
        Me.txtID.TabStop = False
        
        cmdAdd.enabled = False

        m_blnCancelled = True 'default to cancelled until OK is pressed
End Sub

Private Sub cmdAdd_Click()

    m_blnCancelled = False
    Me.Hide

End Sub

Private Sub cmdCancel_Click()
    m_blnCancelled = True
    Me.Hide
End Sub

Private Sub txtName_Change()
    cmdAdd.enabled = Len(txtName.Text) >= 1
End Sub

Private Sub CenterToExcelWindow()
    ' Centers the form relative to the Excel application window.
    ' Works even if StartUpPosition is 0 (manual).
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (Application.Width - Me.Width) / 2
    Me.Top = Application.Top + (Application.Height - Me.Height) / 2
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    'verhindert das UserForm richtig geschlossen wird bzw. UserForm_Terminate ausgeführt wird
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Me.Hide
    End If
End Sub

' Caption of the form
Public Property Let FormCaption(ByVal sValue As String)
    Me.caption = sValue
End Property

' Label text for "Name"
Public Property Let NameLabel(ByVal sValue As String)
    Me.lblName.caption = sValue
End Property

' Label text for "ID"
Public Property Let IDLabel(ByVal sValue As String)
    Me.lblID.caption = sValue
End Property


' Show or hide the whole ID row (label + textbox)
Public Property Let IDVisible(ByVal bValue As Boolean)
    Me.lblID.Visible = bValue
    Me.txtID.Visible = bValue
End Property

Public Property Get nameText() As String

    nameText = Me.txtName.Text

End Property

Public Property Let nameText(ByVal NewValue As String)

    Me.txtName.Text = NewValue
    
End Property

' (Optional convenience) set or read ID display text
Public Property Let IDText(ByVal sValue As String)
    Me.txtID.Text = sValue
End Property

' Indicates whether the dialog was cancelled
Public Property Get Cancelled() As Boolean
    Cancelled = m_blnCancelled
End Property
