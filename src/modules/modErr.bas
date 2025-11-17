Attribute VB_Name = "modErr"
Option Explicit
Option Private Module

' -----------------------------------------------------------------------------------
' Module    : modErr
' Purpose   : Centralized error reporting and logging facade. Holds the logger
'             instance, provides initialization, reporting, show and close methods.
'
' Public API:
'   - InitLogger(basePath, projectName, versionHeader)
'   - ReportError(procedureName, errNo, errLine, showMsg, caption, buttons, customMessage)
'   - ShowLog()
'   - CloseLogger()
'
' Dependencies:
'   - Requires clsLog implementing ILog.
'
' Notes     :
'   - Keeps UI policy (MsgBox) out of clsLog and out of feature modules.
'   - Safe to call ReportError from any ErrHandler.
' -----------------------------------------------------------------------------------

Private m_log As ILog                 ' central logger instance
Private m_cfgPath As String           ' last configured base path
Private m_cfgFileStub As String       ' "<Project>_Log"
Private m_cfgHeader As String         ' e.g., "<Project> <Version>"

' -----------------------------------------------------------------------------------
' Procedure : InitLogger
' Purpose   : Creates/configures the central logger and opens the log file.
'
' Parameters:
'   basePath       [String] - Base folder for log files (must exist).
'   projectName    [String] - Project name used to build the log file name.
'   versionHeader  [String] - Header line written on open (e.g., "App 1.2.3").
'
' Returns   : (none)
'
' Notes     :
'   - Idempotent: Reconfigures if already initialized.
' -----------------------------------------------------------------------------------
Public Sub InitLogger(ByVal basePath As String, ByVal projectName As String, ByVal versionHeader As String)
    m_cfgPath = basePath
    m_cfgFileStub = projectName & "_Log"
    m_cfgHeader = versionHeader

    Set m_log = New clsLog
    m_log.LogFilePath = m_cfgPath
    m_log.LogFileName = m_cfgFileStub
    m_log.OpenLog m_cfgHeader
End Sub

' -----------------------------------------------------------------------------------
' Procedure : ReportError
' Purpose   : Log an error and optionally show a MsgBox in one call.
'
' Parameters:
'   procedureName [String]              - Procedure where error occurred.
'   errNo         [Long]       (Opt)    - Err.Number; if 0, uses current Err.Number.
'   errLine       [Integer]    (Opt)    - Erl; pass 0 if unknown.
'   showMsg       [Boolean]    (Opt)    - True to show MsgBox; default True.
'   caption       [String]     (Opt)    - MsgBox caption; default procedureName.
'   buttons       [VbMsgBoxStyle] (Opt) - MsgBox style; default Critical + OK.
'   customMessage [String]     (Opt)    - Extra information appended to MsgBox.
'
' Returns   : (none)
'
' Notes     :
'   - Lazily (re)creates the logger if not initialized and config is known.
' -----------------------------------------------------------------------------------
Public Sub ReportError( _
        ByVal procedureName As String, _
        Optional ByVal errNo As Long = 0, _
        Optional ByVal errLine As Integer = 0, _
        Optional ByVal showMsg As Boolean = True, _
        Optional ByVal caption As String = vbNullString, _
        Optional ByVal buttons As VbMsgBoxStyle = vbOKOnly Or vbCritical, _
        Optional ByVal customMessage As String = vbNullString)

    Dim desc As String, boxCaption As String, body As String
    EnsureLoggerIfPossible

    If errNo = 0 And Err.Number <> 0 Then errNo = Err.Number
    If LenB(Err.Description) > 0 Then desc = Err.Description

    On Error Resume Next
    If Not m_log Is Nothing Then m_log.WriteError procedureName, errNo, errLine
    On Error GoTo 0

    If showMsg Then
        If LenB(caption) = 0 Then boxCaption = procedureName Else boxCaption = caption
        body = "Error (" & CStr(errNo) & "): " & desc
        If LenB(customMessage) > 0 Then body = body & vbCrLf & customMessage
        MsgBox body, buttons, boxCaption
    End If
End Sub

' -----------------------------------------------------------------------------------
' Procedure : ShowLog
' Purpose   : Displays the current log file (delegates to clsLog).
'
' Parameters: (none)
' Returns   : (none)
' Notes     : Initializes logger on demand if config is present.
' -----------------------------------------------------------------------------------
Public Sub ShowLog()
    EnsureLoggerIfPossible
    If Not m_log Is Nothing Then m_log.ShowLog
End Sub

' -----------------------------------------------------------------------------------
' Procedure : CloseLogger
' Purpose   : Closes the log session and releases the logger instance.
'
' Parameters: (none)
' Returns   : (none)
' Notes     : Safe to call multiple times.
' -----------------------------------------------------------------------------------
Public Sub CloseLogger()
    On Error Resume Next
    If Not m_log Is Nothing Then m_log.CloseLog
    Set m_log = Nothing
    On Error GoTo 0
End Sub

' ===== Private helpers ===============================================================

' -----------------------------------------------------------------------------------
' Procedure : EnsureLoggerIfPossible
' Purpose   : Lazily re-initialize the logger when config exists but instance is gone.
'
' Parameters: (none)
' Returns   : (none)
' Notes     : Does nothing if configuration is not known yet.
' -----------------------------------------------------------------------------------
Private Sub EnsureLoggerIfPossible()
    If m_log Is Nothing Then
        If LenB(m_cfgPath) > 0 And LenB(m_cfgFileStub) > 0 Then
            InitLogger m_cfgPath, Replace(m_cfgFileStub, "_Log", vbNullString), m_cfgHeader
        End If
    End If
End Sub


