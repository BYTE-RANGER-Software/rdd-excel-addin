VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmObjectEdit 
   Caption         =   "New Item"
   ClientHeight    =   4334
   ClientLeft      =   110
   ClientTop       =   440
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
Private m_txt5RequiresValue As Boolean
Private m_txt6RequiresValue As Boolean

Private m_txt1NumericOnly As Boolean
Private m_txt2NumericOnly As Boolean
Private m_txt3NumericOnly As Boolean
Private m_txt4NumericOnly As Boolean
Private m_txt5NumericOnly As Boolean
Private m_txt6NumericOnly As Boolean

Private m_txt1Prefix As String
Private m_txt2Prefix As String
Private m_txt3Prefix As String
Private m_txt4Prefix As String
Private m_txt5Prefix As String
Private m_txt6Prefix As String

Private m_isUpdatingText As Boolean  ' Verhindert Endlosschleifen

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
'   - Disables cmdOK button until validation passes
'   - Sets m_hasCancelled = True (assumes cancellation by default)
' -----------------------------------------------------------------------------------
Private Sub UserForm_Initialize()
        
    ' Default Settings
                
    CenterToExcelWindow
                
    cmdOK.enabled = False

    m_hasCancelled = True 'default to cancelled until OK is pressed
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    'Prevents the UserForm from closing correctly or UserForm_Terminate from being executed.
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Call cmdCancel_Click
    End If
End Sub

' Caption of the form
Public Property Let FormCaption(ByVal sValue As String)
    Me.caption = sValue
End Property

' Label text for m_lbl1
Public Property Let Label1Text(ByVal sValue As String)
    Me.lbl1.caption = sValue
End Property

' Label text for m_lbl2
Public Property Let Label2Text(ByVal sValue As String)
    Me.lbl2.caption = sValue
End Property

' Label text for m_lbl2
Public Property Let Label3Text(ByVal sValue As String)
    Me.lbl3.caption = sValue
End Property

' Label text for m_lbl2
Public Property Let Label4Text(ByVal sValue As String)
    Me.lbl4.caption = sValue
End Property

' Label text for m_lbl2
Public Property Let Label5Text(ByVal sValue As String)
    Me.lbl5.caption = sValue
End Property

' Label text for m_lbl2
Public Property Let Label6Text(ByVal sValue As String)
    Me.lbl6.caption = sValue
End Property

Public Property Get Field1Visible() As Boolean
    Field1Visible = Me.txt1.Visible
End Property

Public Property Let Field1Visible(value As Boolean)

    Dim delta As Single
    delta = Me.txt1.Height + 6
    
    ' If Field1 is hidden, hide all following fields
    If value = False Then
        Field2Visible = False
        Field3Visible = False
        Field4Visible = False
        Field5Visible = False
        Field6Visible = False
    End If

    If value = False And Me.txt1.Visible Then
        Me.Height = Me.Height - delta
        Me.cmdOK.Top = Me.cmdOK.Top - delta
        Me.cmdCancel.Top = Me.cmdCancel.Top - delta
    ElseIf value = True And Not Me.txt1.Visible Then
        Me.Height = Me.Height + delta
        Me.cmdOK.Top = Me.cmdOK.Top + delta
        Me.cmdCancel.Top = Me.cmdCancel.Top + delta
    End If

    Me.txt1.Visible = value
    Me.lbl1.Visible = value
    
End Property

Public Property Get Field2Visible() As Boolean
    Field2Visible = Me.txt2.Visible
End Property

Public Property Let Field2Visible(value As Boolean)

    Dim delta As Single
    delta = Me.txt2.Height + 6

    ' Field2 can only be visible if Field1 is visible
    If value = True And Not Me.txt1.Visible Then
        Field1Visible = True
    ElseIf value = False Then ' If Field2 is hidden, hide all following fields
        Field3Visible = False
        Field4Visible = False
        Field5Visible = False
        Field6Visible = False
    End If
    
    If value = False And Me.txt2.Visible Then
        Me.Height = Me.Height - delta
        Me.cmdOK.Top = Me.cmdOK.Top - delta
        Me.cmdCancel.Top = Me.cmdCancel.Top - delta
    ElseIf value = True And Not Me.txt2.Visible Then
        Me.Height = Me.Height + delta
        Me.cmdOK.Top = Me.cmdOK.Top + delta
        Me.cmdCancel.Top = Me.cmdCancel.Top + delta
    End If

    Me.txt2.Visible = value
    Me.lbl2.Visible = value
    
End Property

Public Property Get Field3Visible() As Boolean
    Field3Visible = Me.txt3.Visible
End Property

Public Property Let Field3Visible(value As Boolean)

    Dim delta As Single
    delta = Me.txt3.Height + 6
    
    ' Field3 can only be visible if Field1 and Field2 are visible
    If value = True And Not Me.txt1.Visible Then
        Field1Visible = True
    End If
    If value = True And Not Me.txt2.Visible Then
        Field2Visible = True
    End If
    
    If value = False Then ' If Field3 is hidden, hide all following fields
        Field4Visible = False
        Field5Visible = False
        Field6Visible = False
    End If

    If value = False And Me.txt3.Visible Then
        Me.Height = Me.Height - delta
        Me.cmdOK.Top = Me.cmdOK.Top - delta
        Me.cmdCancel.Top = Me.cmdCancel.Top - delta
    ElseIf value = True And Not Me.txt3.Visible Then
        Me.Height = Me.Height + delta
        Me.cmdOK.Top = Me.cmdOK.Top + delta
        Me.cmdCancel.Top = Me.cmdCancel.Top + delta
    End If
        
    Me.txt3.Visible = value
    Me.lbl3.Visible = value
End Property

Public Property Get Field4Visible() As Boolean
    Field4Visible = Me.txt4.Visible
End Property

Public Property Let Field4Visible(value As Boolean)

    Dim delta As Single
    delta = Me.txt4.Height + 6
    
    ' Field4 can only be visible if Field1, Field2 and Field3 are visible
    If value = True And Not Me.txt1.Visible Then
        Field1Visible = True
    End If
    If value = True And Not Me.txt2.Visible Then
        Field2Visible = True
    End If
    If value = True And Not Me.txt3.Visible Then
        Field3Visible = True
    End If
    
    If value = False Then ' If Field4 is hidden, hide all following fields
        Field5Visible = False
        Field6Visible = False
    End If

    If value = False And Me.txt4.Visible Then
        Me.Height = Me.Height - delta
        Me.cmdOK.Top = Me.cmdOK.Top - delta
        Me.cmdCancel.Top = Me.cmdCancel.Top - delta
    ElseIf value = True And Not Me.txt4.Visible Then
        Me.Height = Me.Height + delta
        Me.cmdOK.Top = Me.cmdOK.Top + delta
        Me.cmdCancel.Top = Me.cmdCancel.Top + delta
    End If
    
    Me.txt4.Visible = value
    Me.lbl4.Visible = value
End Property

Public Property Get Field5Visible() As Boolean
    Field5Visible = Me.txt5.Visible
End Property

Public Property Let Field5Visible(value As Boolean)

    Dim delta As Single
    delta = Me.txt5.Height + 6
    
    ' Field5 can only be visible if Field1 to Field4 are visible
    If value = True And Not Me.txt1.Visible Then
        Field1Visible = True
    End If
    If value = True And Not Me.txt2.Visible Then
        Field2Visible = True
    End If
    If value = True And Not Me.txt3.Visible Then
        Field3Visible = True
    End If
    If value = True And Not Me.txt4.Visible Then
        Field4Visible = True
    End If
    
    If value = False Then ' If Field5 is hidden, hide all following fields
        Field6Visible = False
    End If

    If value = False And Me.txt5.Visible Then
        Me.Height = Me.Height - delta
        Me.cmdOK.Top = Me.cmdOK.Top - delta
        Me.cmdCancel.Top = Me.cmdCancel.Top - delta
    ElseIf value = True And Not Me.txt5.Visible Then
        Me.Height = Me.Height + delta
        Me.cmdOK.Top = Me.cmdOK.Top + delta
        Me.cmdCancel.Top = Me.cmdCancel.Top + delta
    End If
    
    Me.txt5.Visible = value
    Me.lbl5.Visible = value
End Property

Public Property Get Field6Visible() As Boolean
    Field6Visible = Me.txt6.Visible
End Property

Public Property Let Field6Visible(value As Boolean)

    Dim delta As Single
    delta = Me.txt6.Height + 6
    
    ' Field6 can only be visible if Field1 to 5 are visible
    If value = True And Not Me.txt1.Visible Then
        Field1Visible = True
    End If
    If value = True And Not Me.txt2.Visible Then
        Field2Visible = True
    End If
    If value = True And Not Me.txt3.Visible Then
        Field3Visible = True
    End If
    If value = True And Not Me.txt4.Visible Then
        Field4Visible = True
    End If
    If value = True And Not Me.txt5.Visible Then
        Field5Visible = True
    End If

    If value = False And Me.txt6.Visible Then
        Me.Height = Me.Height - delta
        Me.cmdOK.Top = Me.cmdOK.Top - delta
        Me.cmdCancel.Top = Me.cmdCancel.Top - delta
    ElseIf value = True And Not Me.txt6.Visible Then
        Me.Height = Me.Height + delta
        Me.cmdOK.Top = Me.cmdOK.Top + delta
        Me.cmdCancel.Top = Me.cmdCancel.Top + delta
    End If
    
    Me.txt6.Visible = value
    Me.lbl6.Visible = value
End Property
Public Property Get Text1Value() As String

    Text1Value = Me.txt1.text

End Property

Public Property Let Text1Value(ByVal value As String)

    Me.txt1.text = value
    
End Property

Public Property Get Text2Value() As String

    Text2Value = Me.txt2.text

End Property

Public Property Let Text2Value(ByVal value As String)

    Me.txt2.text = value
    
End Property

Public Property Get Text3Value() As String

    Text3Value = Me.txt3.text

End Property

Public Property Let Text3Value(ByVal value As String)

    Me.txt3.text = value
    
End Property

Public Property Get Text4Value() As String

    Text4Value = Me.txt4.text

End Property

Public Property Let Text4Value(ByVal value As String)

    Me.txt4.text = value
    
End Property

Public Property Get Text5Value() As String

    Text5Value = Me.txt5.text

End Property

Public Property Let Text5Value(ByVal value As String)

    Me.txt5.text = value
    
End Property

Public Property Get Text6Value() As String

    Text6Value = Me.txt6.text

End Property

Public Property Let Text6Value(ByVal value As String)

    Me.txt6.text = value
    
End Property

Public Property Let Text1Tip(ByVal value As String)

    Me.txt1.ControlTipText = value
    
End Property

Public Property Let Text2Tip(ByVal value As String)

    Me.txt2.ControlTipText = value
    
End Property

Public Property Let Text3Tip(ByVal value As String)

    Me.txt3.ControlTipText = value
    
End Property

Public Property Let Text4Tip(ByVal value As String)

    Me.txt4.ControlTipText = value
    
End Property

Public Property Let Text5Tip(ByVal value As String)

    Me.txt5.ControlTipText = value
    
End Property

Public Property Let Text6Tip(ByVal value As String)

    Me.txt6.ControlTipText = value
    
End Property

Public Property Get Text1Locked() As Boolean

    Text1Locked = Me.txt1.Locked
    
End Property

Public Property Let Text1Locked(ByVal value As Boolean)

    Me.txt1.Locked = value

    If value Then
        Me.txt1.BackColor = &H8000000F
    Else
        Me.txt1.BackColor = &H80000005
    End If
    
End Property

Public Property Get Text2Locked() As Boolean

    Text2Locked = Me.txt2.Locked
    
End Property

Public Property Let Text2Locked(ByVal value As Boolean)

    Me.txt2.Locked = value

    If value Then
        Me.txt2.BackColor = &H8000000F
    Else
        Me.txt2.BackColor = &H80000005
    End If
    
End Property

Public Property Get Text3Locked() As Boolean

    Text3Locked = Me.txt3.Locked
    
End Property

Public Property Let Text3Locked(ByVal value As Boolean)

    Me.txt3.Locked = value

    If value Then
        Me.txt3.BackColor = &H8000000F
    Else
        Me.txt3.BackColor = &H80000005
    End If
    
End Property

Public Property Get Text4Locked() As Boolean

    Text4Locked = Me.txt4.Locked
    
End Property

Public Property Let Text4Locked(ByVal value As Boolean)

    Me.txt4.Locked = value

    If value Then
        Me.txt4.BackColor = &H8000000F
    Else
        Me.txt4.BackColor = &H80000005
    End If
    
End Property

Public Property Get Text5Locked() As Boolean

    Text5Locked = Me.txt5.Locked
    
End Property

Public Property Let Text5Locked(ByVal value As Boolean)

    Me.txt5.Locked = value

    If value Then
        Me.txt5.BackColor = &H8000000F
    Else
        Me.txt5.BackColor = &H80000005
    End If
    
End Property

Public Property Get Text6Locked() As Boolean

    Text6Locked = Me.txt6.Locked
    
End Property

Public Property Let Text6Locked(ByVal value As Boolean)

    Me.txt6.Locked = value

    If value Then
        Me.txt6.BackColor = &H8000000F
    Else
        Me.txt6.BackColor = &H80000005
    End If
    
End Property

Public Property Get Text1RequiresValue() As Boolean
    Text1RequiresValue = m_txt1RequiresValue
End Property

Public Property Let Text1RequiresValue(ByVal value As Boolean)
    m_txt1RequiresValue = value
End Property

Public Property Get Text2RequiresValue() As Boolean
    Text2RequiresValue = m_txt2RequiresValue
End Property

Public Property Let Text2RequiresValue(ByVal value As Boolean)
    m_txt2RequiresValue = value
End Property

Public Property Get Text3RequiresValue() As Boolean
    Text3RequiresValue = m_txt3RequiresValue
End Property

Public Property Let Text3RequiresValue(ByVal value As Boolean)
    m_txt3RequiresValue = value
End Property

Public Property Get Text4RequiresValue() As Boolean
    Text4RequiresValue = m_txt4RequiresValue
End Property

Public Property Let Text4RequiresValue(ByVal value As Boolean)
    m_txt4RequiresValue = value
End Property

Public Property Get Text5RequiresValue() As Boolean
    Text5RequiresValue = m_txt5RequiresValue
End Property

Public Property Let Text5RequiresValue(ByVal value As Boolean)
    m_txt5RequiresValue = value
End Property

Public Property Get Text6RequiresValue() As Boolean
    Text6RequiresValue = m_txt6RequiresValue
End Property

Public Property Let Text6RequiresValue(ByVal value As Boolean)
    m_txt6RequiresValue = value
End Property

' -----------------------------------------------------------------------------------
' Property  : Text1NumericOnly (Get/Let)
' Purpose   : Controls whether m_txt1 accepts only numeric input
' -----------------------------------------------------------------------------------
Public Property Get Text1NumericOnly() As Boolean
    Text1NumericOnly = m_txt1NumericOnly
End Property

Public Property Let Text1NumericOnly(ByVal value As Boolean)
    m_txt1NumericOnly = value
End Property

' -----------------------------------------------------------------------------------
' Property  : Text2NumericOnly (Get/Let)
' Purpose   : Controls whether m_txt2 accepts only numeric input
' -----------------------------------------------------------------------------------
Public Property Get Text2NumericOnly() As Boolean
    Text2NumericOnly = m_txt2NumericOnly
End Property

Public Property Let Text2NumericOnly(ByVal value As Boolean)
    m_txt2NumericOnly = value
End Property

' -----------------------------------------------------------------------------------
' Property  : Text3NumericOnly (Get/Let)
' Purpose   : Controls whether m_txt3 accepts only numeric input
' -----------------------------------------------------------------------------------
Public Property Get Text3NumericOnly() As Boolean
    Text3NumericOnly = m_txt3NumericOnly
End Property

Public Property Let Text3NumericOnly(ByVal value As Boolean)
    m_txt3NumericOnly = value
End Property

' -----------------------------------------------------------------------------------
' Property  : Text4NumericOnly (Get/Let)
' Purpose   : Controls whether m_txt4 accepts only numeric input
' -----------------------------------------------------------------------------------
Public Property Get Text4NumericOnly() As Boolean
    Text4NumericOnly = m_txt4NumericOnly
End Property

Public Property Let Text4NumericOnly(ByVal value As Boolean)
    m_txt4NumericOnly = value
End Property

' -----------------------------------------------------------------------------------
' Property  : Text5NumericOnly (Get/Let)
' Purpose   : Controls whether m_txt5 accepts only numeric input
' -----------------------------------------------------------------------------------
Public Property Get Text5NumericOnly() As Boolean
    Text5NumericOnly = m_txt5NumericOnly
End Property

Public Property Let Text5NumericOnly(ByVal value As Boolean)
    m_txt5NumericOnly = value
End Property

' -----------------------------------------------------------------------------------
' Property  : Text6NumericOnly (Get/Let)
' Purpose   : Controls whether m_txt6 accepts only numeric input
' -----------------------------------------------------------------------------------
Public Property Get Text6NumericOnly() As Boolean
    Text6NumericOnly = m_txt6NumericOnly
End Property

Public Property Let Text6NumericOnly(ByVal value As Boolean)
    m_txt6NumericOnly = value
End Property

' -----------------------------------------------------------------------------------
' Property  : Text1Prefix (Get/Let)
' Purpose   : Sets a protected prefix for m_txt1 that cannot be deleted
' -----------------------------------------------------------------------------------
Public Property Get Text1Prefix() As String
    Text1Prefix = m_txt1Prefix
End Property

Public Property Let Text1Prefix(ByVal value As String)
    m_txt1Prefix = value
    If Len(value) > 0 And Not txt1.text Like value & "*" Then
        txt1.text = value
    End If
End Property

' -----------------------------------------------------------------------------------
' Property  : Text2Prefix (Get/Let)
' Purpose   : Sets a protected prefix for m_txt2 that cannot be deleted
' -----------------------------------------------------------------------------------
Public Property Get Text2Prefix() As String
    Text2Prefix = m_txt2Prefix
End Property

Public Property Let Text2Prefix(ByVal value As String)
    m_txt2Prefix = value
    If Len(value) > 0 And Not txt2.text Like value & "*" Then
        txt2.text = value
    End If
End Property

' -----------------------------------------------------------------------------------
' Property  : Text3Prefix (Get/Let)
' Purpose   : Sets a protected prefix for m_txt3 that cannot be deleted
' -----------------------------------------------------------------------------------
Public Property Get Text3Prefix() As String
    Text3Prefix = m_txt3Prefix
End Property

Public Property Let Text3Prefix(ByVal value As String)
    m_txt3Prefix = value
    If Len(value) > 0 And Not txt3.text Like value & "*" Then
        txt3.text = value
    End If
End Property

' -----------------------------------------------------------------------------------
' Property  : Text4Prefix (Get/Let)
' Purpose   : Sets a protected prefix for m_txt4 that cannot be deleted
' -----------------------------------------------------------------------------------
Public Property Get Text4Prefix() As String
    Text4Prefix = m_txt4Prefix
End Property

Public Property Let Text4Prefix(ByVal value As String)
    m_txt4Prefix = value
    If Len(value) > 0 And Not txt4.text Like value & "*" Then
        txt4.text = value
    End If
End Property

' -----------------------------------------------------------------------------------
' Property  : Text5Prefix (Get/Let)
' Purpose   : Sets a protected prefix for m_txt5 that cannot be deleted
' -----------------------------------------------------------------------------------
Public Property Get Text5Prefix() As String
    Text5Prefix = m_txt5Prefix
End Property

Public Property Let Text5Prefix(ByVal value As String)
    m_txt5Prefix = value
    If Len(value) > 0 And Not txt5.text Like value & "*" Then
        txt5.text = value
    End If
End Property

' -----------------------------------------------------------------------------------
' Property  : Text6Prefix (Get/Let)
' Purpose   : Sets a protected prefix for m_txt6 that cannot be deleted
' -----------------------------------------------------------------------------------
Public Property Get Text6Prefix() As String
    Text6Prefix = m_txt6Prefix
End Property

Public Property Let Text6Prefix(ByVal value As String)
    m_txt6Prefix = value
    If Len(value) > 0 And Not txt6.text Like value & "*" Then
        txt6.text = value
    End If
End Property

' -----------------------------------------------------------------------------------
' Property  : Cancelled (Get)
' Purpose   : Returns whether the dialog was cancelled by the user
'
' Returns   : Boolean - True if cancelled, False if confirmed
'
' Notes     :
'   - Defaults to True in UserForm_Initialize
'   - Set to False in cmdOK_Click
' -----------------------------------------------------------------------------------
Public Property Get Cancelled() As Boolean
    Cancelled = m_hasCancelled
End Property

Private Sub UpdateOKButton()
    Dim canEnable As Boolean
    canEnable = True ' Start assuming OK can be enabled

    ' Check every required field
    If m_txt1RequiresValue And Len(txt1.text) = 0 Then canEnable = False
    If m_txt2RequiresValue And Len(txt2.text) = 0 Then canEnable = False
    If m_txt3RequiresValue And Len(txt3.text) = 0 Then canEnable = False
    If m_txt4RequiresValue And Len(txt4.text) = 0 Then canEnable = False
    If m_txt5RequiresValue And Len(txt5.text) = 0 Then canEnable = False
    If m_txt6RequiresValue And Len(txt6.text) = 0 Then canEnable = False

    cmdOK.enabled = canEnable
End Sub

Private Sub txt1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If m_txt1NumericOnly Then KeyAscii = ValidateNumericInput(KeyAscii)
End Sub

Private Sub txt2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If m_txt2NumericOnly Then KeyAscii = ValidateNumericInput(KeyAscii)
End Sub

Private Sub txt3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If m_txt3NumericOnly Then KeyAscii = ValidateNumericInput(KeyAscii)
End Sub

Private Sub txt4_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If m_txt4NumericOnly Then KeyAscii = ValidateNumericInput(KeyAscii)
End Sub

Private Sub txt5_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If m_txt5NumericOnly Then KeyAscii = ValidateNumericInput(KeyAscii)
End Sub

Private Sub txt6_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If m_txt6NumericOnly Then KeyAscii = ValidateNumericInput(KeyAscii)
End Sub

Private Sub txt1_Change()
    EnforcePrefixForTextBox txt1, m_txt1Prefix
    Call UpdateOKButton
End Sub

Private Sub txt2_Change()
    EnforcePrefixForTextBox txt2, m_txt2Prefix
    Call UpdateOKButton
End Sub

Private Sub txt3_Change()
    EnforcePrefixForTextBox txt3, m_txt3Prefix
    Call UpdateOKButton
End Sub

Private Sub txt4_Change()
    EnforcePrefixForTextBox txt4, m_txt4Prefix
    Call UpdateOKButton
End Sub

Private Sub txt5_Change()
    EnforcePrefixForTextBox txt5, m_txt5Prefix
    Call UpdateOKButton
End Sub

Private Sub txt6_Change()
    EnforcePrefixForTextBox txt6, m_txt6Prefix
    Call UpdateOKButton
End Sub

' -----------------------------------------------------------------------------------
' Procedure : cmdOK_Click
' Purpose   : Handles the Add button click event, validates input and closes dialog
'             with confirmed state.
'
' Parameters: (none)
'
' Returns   : (none)
'
' Notes     :
'   - Sets m_hasCancelled = False to indicate successful completion
'   - Should validate txt1 and txtID before closing
'   - Calling code should check Cancelled property before using input
' -----------------------------------------------------------------------------------
Private Sub cmdOK_Click()

    m_hasCancelled = False
    Me.Hide

End Sub

Private Sub cmdCancel_Click()
    m_hasCancelled = True
    Me.Hide
End Sub


Private Sub CenterToExcelWindow()
    ' Centers the form relative to the Excel application window.
    ' Works even if StartUpPosition is 0 (manual).
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (Application.Width - Me.Width) / 2
    Me.Top = Application.Top + (Application.Height - Me.Height) / 2
End Sub

' -----------------------------------------------------------------------------------
' Function  : ValidateNumericInput
' Purpose   : Validates and filters numeric input for textboxes
'
' Parameters: KeyAscii - The key code of the pressed key
'
' Returns   : Integer - 0 to reject the key, original value to accept
'
' Notes     : Allows digits (0-9), backspace (8), and delete (127)
' -----------------------------------------------------------------------------------
Private Function ValidateNumericInput(ByVal KeyAscii As Integer) As Integer
    Select Case KeyAscii
        Case 48 To 57  ' Digits 0-9
            ValidateNumericInput = KeyAscii
        Case 8, 127    ' Backspace and Delete
            ValidateNumericInput = KeyAscii
        Case Else
            ValidateNumericInput = 0  ' Reject all other keys
    End Select
End Function

' -----------------------------------------------------------------------------------
' Procedure : EnforcePrefixForTextBox
' Purpose   : Ensures a textbox maintains its required prefix
'
' Parameters: txt - The textbox control to enforce prefix on
'             prefix - The prefix string to protect
'
' Returns   : (none)
'
' Notes     : Uses m_isUpdatingText flag to prevent infinite loops
' -----------------------------------------------------------------------------------
Private Sub EnforcePrefixForTextBox(ByRef txt As MSForms.TextBox, ByVal prefix As String)
    If m_isUpdatingText Then Exit Sub
    If Len(prefix) = 0 Then Exit Sub
    
    If Not txt.text Like prefix & "*" Then
        m_isUpdatingText = True
        
        Dim userText As String
        Dim cursorPos As Long
        
        cursorPos = txt.SelStart
        userText = Replace(txt.text, prefix, vbNullString)
        
        txt.text = prefix & userText
        
        ' Cursor nach Präfix positionieren
        If cursorPos < Len(prefix) Then cursorPos = Len(prefix)
        txt.SelStart = cursorPos
        
        m_isUpdatingText = False
    End If
End Sub
