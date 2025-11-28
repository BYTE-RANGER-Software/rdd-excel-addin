VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmObjectEdit 
   Caption         =   "New Item"
   ClientHeight    =   3267
   ClientLeft      =   99
   ClientTop       =   429
   ClientWidth     =   7799
   OleObjectBlob   =   "frmObjectEdit.frx":0000
End
Attribute VB_Name = "frmObjectEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_hasCancelled As Boolean
Private m_txt1RequiresValue As Boolean
Private m_txt2RequiresValue As Boolean
Private m_txt3RequiresValue As Boolean
Private m_txt4RequiresValue As Boolean

' -----------------------------------------------------------------------------------
' Procedure : UserForm_Initialize
' Purpose   : Initializes the New Item dialog, configures controls visibility and
'             default states.
'
' Parameters: (none)
'
' Returns   : (none)
'
' Notes     :
'   - Centers form to Excel window
'   - Hides ID field (m_lbl2, txtID) by default
'   - Locks txtID to prevent editing
'   - Disables m_cmdOK button until validation passes
'   - Sets m_hasCancelled = True (assumes cancellation by default)
' -----------------------------------------------------------------------------------
Private Sub UserForm_Initialize()
        
        ' Default Settings
        Field1Visible = True
        Field2Visible = False
        Field3Visible = False
        Field4Visible = False
                
        CenterToExcelWindow
        
        Text1Locked = False
        Text2Locked = True
        
        m_cmdOK.enabled = False

        m_hasCancelled = True 'default to cancelled until OK is pressed
End Sub

' -----------------------------------------------------------------------------------
' Procedure : m_cmdOK_Click
' Purpose   : Handles the Add button click event, validates input and closes dialog
'             with confirmed state.
'
' Parameters: (none)
'
' Returns   : (none)
'
' Notes     :
'   - Sets m_hasCancelled = False to indicate successful completion
'   - Should validate m_txt1 and txtID before closing
'   - Calling code should check Cancelled property before using input
' -----------------------------------------------------------------------------------
Private Sub m_cmdOK_Click()

    m_hasCancelled = False
    Me.Hide

End Sub

Private Sub m_cmdCancel_Click()
    m_hasCancelled = True
    Me.Hide
End Sub

Private Sub m_txt1_Change()
    Call UpdateOKButton
End Sub
Private Sub m_txt2_Change()
    Call UpdateOKButton
End Sub
Private Sub m_txt3_Change()
    Call UpdateOKButton
End Sub
Private Sub m_txt4_Change()
    Call UpdateOKButton
End Sub

Private Sub CenterToExcelWindow()
    ' Centers the form relative to the Excel application window.
    ' Works even if StartUpPosition is 0 (manual).
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (Application.Width - Me.Width) / 2
    Me.Top = Application.Top + (Application.Height - Me.Height) / 2
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    'verhindert das UserForm richtig geschlossen wird bzw. UserForm_Terminate ausgef hrt wird
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Me.Hide
    End If
End Sub

' Caption of the form
Public Property Let FormCaption(ByVal sValue As String)
    Me.caption = sValue
End Property

' Label text for m_lbl1
Public Property Let Label1Text(ByVal sValue As String)
    Me.m_lbl1.caption = sValue
End Property

' Label text for m_lbl2
Public Property Let Label2Text(ByVal sValue As String)
    Me.m_lbl2.caption = sValue
End Property

' Label text for m_lbl2
Public Property Let Label3Text(ByVal sValue As String)
    Me.m_lbl3.caption = sValue
End Property

' Label text for m_lbl2
Public Property Let Label4Text(ByVal sValue As String)
    Me.m_lbl4.caption = sValue
End Property

Public Property Get Field1Visible() As Boolean
    Field1Visible = Me.m_txt1.Visible
End Property

Public Property Let Field1Visible(value As Boolean)

    Dim delta As Single
    delta = Me.m_txt1.Height + 6

    If value = False And Me.m_txt1.Visible Then
        Me.Height = Me.Height - delta
        Me.m_cmdOK.Top = Me.m_cmdOK.Top - delta
        Me.m_cmdCancel.Top = Me.m_cmdCancel.Top - delta
     ElseIf value = True And Not Me.m_txt1.Visible Then
        Me.Height = Me.Height + delta
        Me.m_cmdOK.Top = Me.m_cmdOK.Top + delta
        Me.m_cmdCancel.Top = Me.m_cmdCancel.Top + delta
    End If

    Me.m_txt1.Visible = value
    Me.m_lbl1.Visible = value
    
End Property

Public Property Get Field2Visible() As Boolean
    Field2Visible = Me.m_txt2.Visible
End Property

Public Property Let Field2Visible(value As Boolean)

    Dim delta As Single
    delta = Me.m_txt2.Height + 6

    If value = False And Me.m_txt2.Visible Then
        Me.Height = Me.Height - delta
        Me.m_cmdOK.Top = Me.m_cmdOK.Top - delta
        Me.m_cmdCancel.Top = Me.m_cmdCancel.Top - delta
     ElseIf value = True And Not Me.m_txt2.Visible Then
        Me.Height = Me.Height + delta
        Me.m_cmdOK.Top = Me.m_cmdOK.Top + delta
        Me.m_cmdCancel.Top = Me.m_cmdCancel.Top + delta
    End If

    Me.m_txt2.Visible = value
    Me.m_lbl2.Visible = value
    
End Property

Public Property Get Field3Visible() As Boolean
    Field3Visible = Me.m_txt3.Visible
End Property

Public Property Let Field3Visible(value As Boolean)

    Dim delta As Single
    delta = Me.m_txt3.Height + 6

    If value = False And Me.m_txt3.Visible Then
        Me.Height = Me.Height - delta
        Me.m_cmdOK.Top = Me.m_cmdOK.Top - delta
        Me.m_cmdCancel.Top = Me.m_cmdCancel.Top - delta
    ElseIf value = True And Not Me.m_txt3.Visible Then
        Me.Height = Me.Height + delta
        Me.m_cmdOK.Top = Me.m_cmdOK.Top + delta
        Me.m_cmdCancel.Top = Me.m_cmdCancel.Top + delta
    End If
        
    Me.m_txt3.Visible = value
    Me.m_lbl3.Visible = value
End Property

Public Property Get Field4Visible() As Boolean
    Field4Visible = Me.m_txt4.Visible
End Property

Public Property Let Field4Visible(value As Boolean)

    Dim delta As Single
    delta = Me.m_txt4.Height + 6

    If value = False And Me.m_txt4.Visible Then
        Me.Height = Me.Height - delta
        Me.m_cmdOK.Top = Me.m_cmdOK.Top - delta
        Me.m_cmdCancel.Top = Me.m_cmdCancel.Top - delta
    ElseIf value = True And Not Me.m_txt4.Visible Then
        Me.Height = Me.Height + delta
        Me.m_cmdOK.Top = Me.m_cmdOK.Top + delta
        Me.m_cmdCancel.Top = Me.m_cmdCancel.Top + delta
    End If
    
    Me.m_txt4.Visible = value
    Me.m_lbl4.Visible = value
End Property

Public Property Get Text1Value() As String

    Text1Value = Me.m_txt1.text

End Property

Public Property Let Text1Value(ByVal value As String)

    Me.m_txt1.text = value
    
End Property

Public Property Get Text2Value() As String

    Text2Value = Me.m_txt2.text

End Property

Public Property Let Text2Value(ByVal value As String)

    Me.m_txt2.text = value
    
End Property

Public Property Get Text3Value() As String

    Text3Value = Me.m_txt3.text

End Property

Public Property Let Text3Value(ByVal value As String)

    Me.m_txt3.text = value
    
End Property

Public Property Get Text4Value() As String

    Text4Value = Me.m_txt4.text

End Property

Public Property Let Text4Value(ByVal value As String)

    Me.m_txt4.text = value
    
End Property

Public Property Get Text1Locked() As Boolean

    Text1Locked = Me.m_txt1.Locked
    
End Property

Public Property Let Text1Locked(ByVal value As Boolean)

    Me.m_txt1.Locked = value

    If value Then
        Me.m_txt1.BackColor = &H8000000F
    Else
        Me.m_txt1.BackColor = &H80000005
    End If
    
End Property

Public Property Get Text2Locked() As Boolean

    Text2Locked = Me.m_txt2.Locked
    
End Property

Public Property Let Text2Locked(ByVal value As Boolean)

    Me.m_txt2.Locked = value

    If value Then
        Me.m_txt2.BackColor = &H8000000F
    Else
        Me.m_txt2.BackColor = &H80000005
    End If
    
End Property

Public Property Get Text3Locked() As Boolean

    Text3Locked = Me.m_txt3.Locked
    
End Property

Public Property Let Text3Locked(ByVal value As Boolean)

    Me.m_txt3.Locked = value

    If value Then
        Me.m_txt3.BackColor = &H8000000F
    Else
        Me.m_txt3.BackColor = &H80000005
    End If
    
End Property

Public Property Get Text4Locked() As Boolean

    Text4Locked = Me.m_txt4.Locked
    
End Property

Public Property Let Text4Locked(ByVal value As Boolean)

    Me.m_txt4.Locked = value

    If value Then
        Me.m_txt4.BackColor = &H8000000F
    Else
        Me.m_txt4.BackColor = &H80000005
    End If
    
End Property

Public Property Get Text1RequiresValue() As Boolean
    Text1RequiresValue = m_txt1RequiresValue
End Property

Public Property Let Text1RequiresValue(ByVal value As Boolean)
     m_txt1RequiresValue = value
End Property

Public Property Get Text2RequiresValue() As Boolean
    Text2RequiresValue = m_txt1RequiresValue
End Property

Public Property Let Text2RequiresValue(ByVal value As Boolean)
     m_txt2RequiresValue = value
End Property

Public Property Get Text3RequiresValue() As Boolean
    Text3RequiresValue = m_txt1RequiresValue
End Property

Public Property Let Text3RequiresValue(ByVal value As Boolean)
     m_txt3RequiresValue = value
End Property

Public Property Get Text4RequiresValue() As Boolean
    Text4RequiresValue = m_txt1RequiresValue
End Property

Public Property Let Text4RequiresValue(ByVal value As Boolean)
     m_txt4RequiresValue = value
End Property

' -----------------------------------------------------------------------------------
' Property  : Cancelled (Get)
' Purpose   : Returns whether the dialog was cancelled by the user
'
' Returns   : Boolean - True if cancelled, False if confirmed
'
' Notes     :
'   - Defaults to True in UserForm_Initialize
'   - Set to False in m_cmdOK_Click
' -----------------------------------------------------------------------------------
Public Property Get Cancelled() As Boolean
    Cancelled = m_hasCancelled
End Property

Private Sub UpdateOKButton()
    Dim canEnable As Boolean
    canEnable = True ' Start assuming OK can be enabled

    ' Check every required field
    If m_txt1RequiresValue And Len(m_txt1.text) = 0 Then canEnable = False
    If m_txt2RequiresValue And Len(m_txt2.text) = 0 Then canEnable = False
    If m_txt3RequiresValue And Len(m_txt3.text) = 0 Then canEnable = False
    If m_txt4RequiresValue And Len(m_txt4.text) = 0 Then canEnable = False

    m_cmdOK.enabled = canEnable
End Sub


