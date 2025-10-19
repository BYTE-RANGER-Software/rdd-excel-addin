Attribute VB_Name = "modMain"
Option Explicit
Option Private Module

Private m_objLog As clsLog          ' central Logger
Private m_strAppTempPath As String

Private m_objAppEvents As clsAppEvents

' -----------------------------------------------------------------------------------
' Function  : AppProjectName
' Purpose   : Returns the VBA project name (best-effort).
' Parameters: -
' Returns   : String - Project name
' Notes     : Uses Err.Source trick; consider a constant or document property instead.
' -----------------------------------------------------------------------------------
Public Property Get AppProjectName() As String
    On Error Resume Next
    Err.Raise 999
    AppProjectName = Err.Source
    On Error GoTo 0
End Property

' -----------------------------------------------------------------------------------
' Function  : AppTempPath (Get/Let)
' Purpose   : Gets/Sets the path used for temp/log files.
' Parameters: -
' Returns   : String (Get)
' Notes     : Ensure trailing "\" when setting.
' -----------------------------------------------------------------------------------
Public Property Get AppTempPath() As String

    AppTempPath = m_strAppTempPath

End Property

Public Property Let AppTempPath(ByVal strNewValue As String)

   ' Ensure trailing backslash
    If Len(strNewValue) > 0 Then
        If Right$(strNewValue, 1) <> "\" Then
            strNewValue = strNewValue & "\"
        End If
    End If
    m_strAppTempPath = strNewValue

End Property

' -----------------------------------------------------------------------------------
' Function  : AppVersion
' Purpose   : Returns version string from a custom document property.
' Parameters: -
' Returns   : String - e.g., "1.2.3"
' Notes     : Uses GetDocumentPropertyValue("RDD_AddInVersion").
' -----------------------------------------------------------------------------------
Public Property Get AppVersion() As String
    AppVersion = GetDocumentPropertyValue(ThisWorkbook, "RDD_AddInVersion", "0.0.0")
End Property

' -----------------------------------------------------------------------------------
' Function  : AppStart
' Purpose   : Application startup: init logging, wire App events, init state, refresh UI.
' Parameters: (none)
' Returns   :
' Notes     : Requires clsAppEvents and (falls genutzt) eine State-Instanz.
' -----------------------------------------------------------------------------------
Public Sub AppStart()

    ' Ensure temp path exists before logging
    m_strAppTempPath = modUtil.GetTempFolder & "\BYTE RANGER"
    If Dir(m_strAppTempPath, vbDirectory) = "" Then MkDir m_strAppTempPath
    m_strAppTempPath = m_strAppTempPath & "\" & AppProjectName & "\"
    If Dir(m_strAppTempPath, vbDirectory) = "" Then MkDir m_strAppTempPath
        
    ' Logger
    Call OpenLog

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
    ' Detach events
    DisconnectEventHandler

    '  Update ribbon/UI and clear state
    clsState.InvalidateRibbon
    clsState.Cleanup

    ' close Log
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

    On Error GoTo errHandler
    
    If m_objAppEvents Is Nothing Then Set m_objAppEvents = New clsAppEvents
    Set m_objAppEvents.App = Application
    
    Exit Sub
    
errHandler:
    On Error Resume Next
    LogError "ConnectEventHandler", Err.Number, Erl
    
    MsgBox "Unable to enable application events." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbOKOnly Or vbCritical, AppProjectName
           
    ' Ensure instance is detached/cleared on failure
    If Not m_objAppEvents Is Nothing Then Set m_objAppEvents.App = Nothing
    Set m_objAppEvents = Nothing
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

' -----------------------------------------------------------------------------------
' Procedure : LogError
' Purpose   : Writes an error entry to the central log. Ensures the logger instance
'             is initialized (opens the log on first use).
'
' Parameters:
'   sNameFunc [String]   - Name of the procedure where the error occurred
'   iErrNum   [Long]     - (Optional) VBA Err.Number
'   iErrLine  [Integer]  - (Optional) Line number (Erl)
'
' Returns   : (none)
'
' Notes     :
'   - Safe to call in any error Handler block.
'   - Calls OpenLog on demand; forwards to clsLog.WriteError.
' -----------------------------------------------------------------------------------
Public Sub LogError(ByVal sNameFunc As String, Optional iErrNum As Long = -9999, Optional iErrLine As Integer = -1)
    On Error Resume Next
    If m_objLog Is Nothing Then
        Call OpenLog
    End If
    
    m_objLog.WriteError sNameFunc, iErrNum, iErrLine
    On Error GoTo 0
End Sub

' -----------------------------------------------------------------------------------
' Procedure : OpenLog
' Purpose   : Creates the central logger instance and opens the log file with a
'             standard header (project name + version).
'
' Parameters:
'   (none)
'
' Returns   : (none)
'
' Notes     :
'   - Uses AppTempPath as base directory (must exist).
'   - Sets file name to "<Project>_Log".
'   - Private helper intended to be called by startup or ShowLog.
' -----------------------------------------------------------------------------------
Private Sub OpenLog()
    Set m_objLog = New clsLog
    m_objLog.PathLogFile = AppTempPath
    m_objLog.LogFilename = AppProjectName & "_Log"
    m_objLog.OpenLog AppProjectName & " " & AppVersion
End Sub

' Mark workbook as compatible with this add-in if not already tagged (ID/value defined by APP_DOC_TAG_KEY/VAL)
Public Sub EnsureWorkbookIsTagged(ByVal wb As Workbook)
    If Not modProps.DocumentPropertyExists(wb, APP_DOC_TAG_KEY) Then
        modProps.SetDocumentProperty wb, APP_DOC_TAG_KEY, APP_DOC_TAG_VAL
    End If
End Sub

' Check in Ribbon getVisible, or before enabling features
Public Function IsAddinWorkbook(ByVal wb As Workbook) As Boolean
    On Error GoTo SafeExit
    If wb Is Nothing Then GoTo SafeExit
    
    Dim val As String
    val = modProps.GetDocumentPropertyValue(wb, APP_DOC_TAG_KEY, "")
    IsAddinWorkbook = (StrComp(val, APP_DOC_TAG_VAL, vbTextCompare) = 0)
    Exit Function
SafeExit:
    IsAddinWorkbook = False
    IsAddinWorkbook = (modProps.GetDocumentProperty(wb, APP_DOC_TAG_KEY) = APP_DOC_TAG_VAL)
End Function

' ================================================
' --- Ribbon Callback Targets ---
' ================================================

' -----------------------------------------------------------------------------------
' Procedure : ShowLog
' Purpose   : Displays the current log to the user. Ensures a logger instance exists.
'
' Parameters:
'   (none)
'
' Returns   : (none)
'
' Notes     :
'   - Creates and opens the log lazily if needed.
'   - Delegates UI/display to clsLog.ShowLog.
' -----------------------------------------------------------------------------------
Public Sub ShowLog()
    If m_objLog Is Nothing Then
        Call OpenLog
    End If
    m_objLog.ShowLog
End Sub

' -----------------------------------------------------------------------------------
' Procedure : ShowManual
' Purpose   : Opens the manual (PDF/HTML) from the configured path. Shows a message
'             if the file cannot be found and logs any runtime errors.
'
' Parameters:
'   (none)
'
' Returns   : (none)
'
' Notes     :
'   - Uses ReplaceWildCards(Opt_ManualPath) and MANUAL_FILENAME.
'   - On failure, displays a message box and calls LogError with Err/Erl.
' -----------------------------------------------------------------------------------
Public Sub ShowManual()

    On Error GoTo errHandler
    Dim intErr As Long
    
    Dim strPath As String
    'strPath = ReplaceWildCards(Opt_ManualPath) & "\"  ' TODO: Add options Formular


    If Dir(strPath & FILENAME_MANUAL) <> "" Then
        ThisWorkbook.FollowHyperlink strPath & FILENAME_MANUAL
    Else
        ' Datei xxxx nicht gefunden
        MsgBox "File " & strPath & FILENAME_MANUAL & _
            " not found!", vbExclamation, AppProjectName
    End If
    

    On Error GoTo 0
    Exit Sub

errHandler:
    intErr = Err.Number
    MsgBox "Error " & intErr & " (" & Err.Description & ") in procedure ShowManual, line " & Erl & ".", vbCritical, AppProjectName
    LogError "ShowManual", intErr, Erl

End Sub



