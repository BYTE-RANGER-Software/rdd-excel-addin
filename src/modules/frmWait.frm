VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmWait 
   Caption         =   "Initializing..."
   ClientHeight    =   990
   ClientLeft      =   121
   ClientTop       =   462
   ClientWidth     =   4554
   OleObjectBlob   =   "frmWait.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const GWL_STYLE = -16
Private Const WS_CAPTION = &HC00000

Private Declare PtrSafe Function GetWindowLong _
    Lib "User32.dll" Alias "GetWindowLongA" ( _
    ByVal HWnd As LongPtr, _
    ByVal nIndex As LongPtr) As Long
Private Declare PtrSafe Function SetWindowLong _
    Lib "User32.dll" Alias "SetWindowLongA" ( _
    ByVal HWnd As LongPtr, _
    ByVal nIndex As LongPtr, _
    ByVal dwNewLongPtr As LongPtr) As LongPtr
Private Declare PtrSafe Function DrawMenuBar _
    Lib "User32.dll" ( _
    ByVal HWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function FindWindow _
    Lib "user32" Alias "FindWindowA" ( _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long
    
Private m_blnInit As Boolean


' Initialize the Userform frmWait for the first use
Public Sub Init()
    If m_blnInit Then Exit Sub
    
    Debug.Print "Init frmWait"
    Application.ScreenUpdating = False
    Me.Label1.caption = "Initializing..."
    Me.Show
    Call HideTitleBar
            
    Call CenterToExcelWindow
    Me.Hide
    Application.ScreenUpdating = True
    
    m_blnInit = True
End Sub

Public Property Let Message(ByVal value As String)
    Me.Label1.caption = value
End Property

Public Sub ShowDialog(Optional ByVal msg As String = "Please Wait ...")
    Me.Label1.caption = msg
    Me.Show
    DoEvents
End Sub
Private Sub CenterToExcelWindow()
    ' Centers the form relative to the Excel application window.
    ' Works even if StartUpPosition is 0 (manual).
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (Application.Width - Me.Width) / 2
    Me.Top = Application.Top + (Application.Height - Me.Height) / 2
End Sub

Private Sub HideTitleBar()
    Dim lngWindow As Long
    Dim lFrmHdl As Long
    lFrmHdl = FindWindow(vbNullString, Me.caption)
    lngWindow = GetWindowLong(lFrmHdl, GWL_STYLE)
    lngWindow = lngWindow And (Not WS_CAPTION)
    Call SetWindowLong(lFrmHdl, GWL_STYLE, lngWindow)
    Call DrawMenuBar(lFrmHdl)
End Sub
