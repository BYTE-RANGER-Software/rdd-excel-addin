VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAbout 
   Caption         =   "About %1"
   ClientHeight    =   3872
   ClientLeft      =   44
   ClientTop       =   385
   ClientWidth     =   5863
   OleObjectBlob   =   "frmAbout.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub UserForm_Initialize()
        ' Align form
        CenterToExcelWindow

        Me.Caption = Replace$(Me.Caption, "%1", modMain.AppProjectName)
    
        Me.lblAppName.Caption = modMain.AppProjectName
        Me.lblVersion.Caption = modMain.AppVersion
        Me.lblCopyright.Caption = "© " & Year(Date) & " " & "Byte Ranger Software"
        Me.lblCompanyName.Caption = "Byte Ranger Software"
        Me.lblWebsiteLink.Caption = "Website"
        Me.lblWebsiteLink.Tag = "https://byte-ranger-software.github.io/"
        Me.lblLicenseLink.Caption = "MIT License"
        Me.lblLicenseLink.Tag = "https://opensource.org/licenses/MIT"
        
        Me.lblDescription.Caption = "Room design document add-in, including puzzle dependency diagram."
End Sub

Private Sub lblLicenseLink_Click()
    On Error GoTo errHandler
    Dim iErr As Long
    
    OpenLinkSafe Me.lblLicenseLink.Tag
        
    On Error GoTo 0
    Exit Sub

errHandler:
    iErr = Err.Number
    MsgBox "Error:" & iErr & " (" & Err.Description & ") ", vbOKOnly Or vbCritical, modMain.AppProjectName
    LogError "lblLicenseLink_Click", iErr
End Sub

Private Sub lblWebsiteLink_Click()
    On Error GoTo errHandler
    Dim iErr As Long
    
    OpenLinkSafe Me.lblWebsiteLink.Tag
        
    On Error GoTo 0
    Exit Sub

errHandler:
    iErr = Err.Number
    MsgBox "Error:" & iErr & " (" & Err.Description & ") ", vbOKOnly Or vbCritical, AppProjectName
    LogError "lblWebsiteLink_Click", iErr
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub CenterToExcelWindow()
    ' Centers the form relative to the Excel application window.
    ' Works even if StartUpPosition is 0 (manual).
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (Application.Width - Me.Width) / 2
    Me.Top = Application.Top + (Application.Height - Me.Height) / 2
End Sub

Private Sub OpenLinkSafe(ByVal sUrl As String)
    On Error Resume Next
    If Len(sUrl) > 0 Then ThisWorkbook.FollowHyperlink sUrl
End Sub
