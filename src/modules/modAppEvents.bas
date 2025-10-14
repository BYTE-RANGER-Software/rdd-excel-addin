Attribute VB_Name = "modAppEvents"
Option Explicit
Option Private Module
Dim AppEvents As New clsAppEvents

Public Sub ConnectEventHandler()

    On Error Resume Next
    Set AppEvents.App = Application
    If Err <> 0 Then
        'Unable to enable application events.
        MsgBox "Unable to enable application events.", vbOKOnly Or vbCritical, "RDD Add-In"
    End If
End Sub

Public Sub DisconnectEventHandler()
    Set AppEvents.App = Nothing
End Sub

