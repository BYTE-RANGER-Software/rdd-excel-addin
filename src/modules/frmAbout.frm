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
'   - Attribute VB_PredeclaredId = True
Option Explicit

' -----------------------------------------------------------------------------------
' Procedure : UserForm_Initialize
' Purpose   : Initializes the About dialog with application information including
'             name, version, copyright, and links to website and license.
'
' Parameters: (none)
'
' Returns   : (none)
'
' Notes     :
'   - Centers form to Excel window
'   - Populates labels with values from modMain (AppProjectName, AppVersion)
'   - Sets dynamic copyright year
'   - Configures clickable links for website and license
'   - Description: "Room design document add-in, including puzzle dependency diagram"
'   - Error handling with CleanExit label
' -----------------------------------------------------------------------------------
Private Sub UserForm_Initialize()
        On Error GoTo ErrHandler
        ' Align form
        CenterToExcelWindow

        Me.caption = Replace$(Me.caption, "%1", modMain.AppProjectName)
    
        Me.lblAppName.caption = modMain.AppProjectName
        Me.lblVersion.caption = modMain.AppVersion
        Me.lblCopyright.caption = "  " & Year(Date) & " " & "Byte Ranger Software"
        Me.lblCompanyName.caption = "Byte Ranger Software"
        Me.lblWebsiteLink.caption = "Website"
        Me.lblWebsiteLink.Tag = "https://byte-ranger-software.github.io/"
        Me.lblLicenseLink.caption = "MIT License"
        Me.lblLicenseLink.Tag = "https://opensource.org/licenses/MIT"
        
        Me.lblDescription.caption = "Room design document add-in, including puzzle dependency diagram."
        
CleanExit:
    Exit Sub
ErrHandler:
    modErr.ReportError "UserForm_Initialize", Err.Number, Erl, True, modMain.AppProjectName, , _
                       "Form=" & Me.name
    Resume CleanExit
End Sub

Private Sub lblLicenseLink_Click()
    On Error GoTo ErrHandler
    Dim iErr As Long
    
    OpenLinkSafe Me.lblLicenseLink.Tag
        
CleanExit:
    Exit Sub
ErrHandler:
    modErr.ReportError "lblLicenseLink_Click", Err.Number, Erl, True, modMain.AppProjectName, , _
                       "Form=" & Me.name & ", Control=lblLicenseLink"
    Resume CleanExit
End Sub

Private Sub lblWebsiteLink_Click()
    On Error GoTo ErrHandler
    Dim iErr As Long
    
    OpenLinkSafe Me.lblWebsiteLink.Tag
        
CleanExit:
    Exit Sub
ErrHandler:
    modErr.ReportError "lblWebsiteLink_Click", Err.Number, Erl, True, modMain.AppProjectName, , _
                       "Form=" & Me.name & ", Control=lblWebsiteLink"
    Resume CleanExit
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

' -----------------------------------------------------------------------------------
' Procedure : CenterToExcelWindow
' Purpose   : Centers the UserForm to the Excel application window
'
' Parameters: (none)
'
' Returns   : (none)
'
' Notes     :
'   - Calculates Excel window center position
'   - Adjusts form's StartUpPosition or Left/Top properties
'   - Common utility used across all UserForms
' -----------------------------------------------------------------------------------
Private Sub CenterToExcelWindow()
    ' Centers the form relative to the Excel application window.
    ' Works even if StartUpPosition is 0 (manual).
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (Application.Width - Me.Width) / 2
    Me.Top = Application.Top + (Application.Height - Me.Height) / 2
End Sub

Private Sub OpenLinkSafe(ByVal sUrl As String)
    On Error GoTo ErrHandler
    If Len(sUrl) > 0 Then ThisWorkbook.FollowHyperlink sUrl
CleanExit:
    Exit Sub
ErrHandler:
    ' Log but avoid an extra MsgBox here; the caller already shows UI on error.
    modErr.ReportError "OpenLinkSafe", Err.Number, Erl, False, modMain.AppProjectName
    Resume CleanExit
End Sub
