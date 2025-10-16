Attribute VB_Name = "modMain"
Option Explicit
Option Private Module

Private m_objLog As clsLog          ' zentraler Logger
Dim m_objAppEvents As clsAppEvents

' -----------------------------------------------------------------------------------
' Function  : AppStart
' Purpose   : Application startup: init logging, wire App events, init state, refresh UI.
' Parameters: (none)
' Returns   :
' Notes     : Requires clsAppEvents and (falls genutzt) eine State-Instanz.
' -----------------------------------------------------------------------------------
Public Sub AppStart()
    ' Logger
    Set m_objLog = New clsLog
    m_objLog.PathLogFile = AppTempPath
    m_objLog.LogFilename = AppProjectName & "_Error"
    m_objLog.OpenLog AppProjectName & " " & AppVersion

    ' App-Events verbinden
    ConnectEventHandler

    ' State initialisieren (so wie du es nutzt)
    clsState.Init
    clsState.InvalidateRibbon
End Sub

' -----------------------------------------------------------------------------------
' Function  : AppStop
' Purpose   : Application shutdown: unhook events, cleanup state, close log.
' Parameters: (none)
' Returns   :
' Notes     : Safe to call multiple times.
' -----------------------------------------------------------------------------------
Public Sub AppStop()
    ' Events lösen
    DisconnectEventHandler

    ' Ribbon/UI aktualisieren und State räumen
    clsState.InvalidateRibbon
    clsState.Cleanup

    ' Log schließen
    If Not m_objLog Is Nothing Then
        m_objLog.CloseLog
    End If
    Set m_objLog = Nothing
End Sub

' -----------------------------------------------------------------------------------
' Function  : ConnectEventHandler
' Purpose   : Enables application-level event handling by assigning the Excel
'             Application object to the clsAppEvents instance.
'
' Parameters:
'   (none)
'
' Returns   :
'
' Notes     :
'   - Requires a class module `clsAppEvents` exposing an `App` property (WithEvents).
'   - Uses `On Error Resume Next`; shows a critical message box on failure
'     ("Unable to enable application events.", title "RDD Add-In").
'   - Keeps a private instance alive in this module.
' -----------------------------------------------------------------------------------
Public Sub ConnectEventHandler()

    On Error Resume Next
    If m_objAppEvents Is Nothing Then Set m_objAppEvents = New clsAppEvents
    Set m_objAppEvents.App = Application
    If Err.Number <> 0 Then
        MsgBox "Unable to enable application events.", vbOKOnly Or vbCritical, "RDD Add-In"
    End If
    On Error GoTo 0
End Sub

' -----------------------------------------------------------------------------------
' Function  : DisconnectEventHandler
' Purpose   : Disables application-level event handling by releasing the reference
'             to the Excel Application object.
'
' Parameters:
'   (none)
'
' Returns   :
'
' Notes     :
'   - Safe to call multiple times; sets `AppEvents.App = Nothing`.
'   - Clears instance reference so events stop firing.
' -----------------------------------------------------------------------------------
Public Sub DisconnectEventHandler()
    On Error Resume Next
    If Not m_objAppEvents Is Nothing Then Set m_objAppEvents.App = Nothing
    Set m_objAppEvents = Nothing
    On Error GoTo 0
End Sub

