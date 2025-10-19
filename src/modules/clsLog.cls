VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' -----------------------------------------------------------------------------------
' Class     : clsLog
' Purpose   : Buffered text-file logger for VBA projects. Collects log entries in
'             memory and flushes them to disk on demand or when a size threshold
'             is exceeded. Supports separators, error entries, and quick viewing.
'
' Usage     :
'   Dim oLog As New clsLog
'   oLog.PathLogFile = ThisWorkbook.Path & "\logs\"
'   oLog.LogFilename = "session"
'   oLog.OpenLog "Session start", True
'   oLog.AddRecord "Initialized", "OK", 0, LOG_SEPARATOR_BEFORE
'   ' ... work ...
'   oLog.SaveLog
'   oLog.CloseLog
'
'
' Notes     :
'   - Default folder: ThisWorkbook.Path
'   - Default file  : "common.log" (extension enforced to ".log")
'   - Timestamp format: "YYYY-MM-DD HH:NN:SS"
'   - Line breaks use vbNewLine; separator text configurable via LogSeparator
'   - I/O helpers use On Error Resume Next; check Boolean returns for success.
'
' -----------------------------------------------------------------------------------

Option Explicit

' Adds visual separators around log records (where to place them relative to the entry)
Public Enum LOG_SEPARATOR_TYPE
    LOG_SEPARATOR_NONE = 0                 ' No separator line
    LOG_SEPARATOR_BEFORE = 1               ' Insert separator before the entry
    LOG_SEPARATOR_AFTER = 2                ' Insert separator after the entry
    LOG_SEPARATOR_BEFORE_AND_AFTER = 3     ' Insert separator before and after the entry
End Enum

Private Const DEFAULT_LOG_EXT As String = ".log"
Private Const DEFAULT_LOG_FILENAME As String = "common" & DEFAULT_LOG_EXT

Private m_strPathLogFile As String          ' base folder for log files
Private m_strLogFileFullName As String

Private m_strBuffer As String
Private m_lngMaxLogSizeBeforeSave As Long
Private m_strLogSeparator As String

' -----------------------------------------------------------------------------------
' Function  : Class_Initialize
' Purpose   : Class setup; initializes default path, separator text, max buffer
'             size, and default log file full path.
'
' Parameters:
'   (none)
'
' Returns   :
'
' Notes     :
'   - Default path = ThisWorkbook.Path & "\".
'   - Default separator = 100× "=" + CRLF.
'   - Default maxLogSizeBeforeSave = 5000 (if zero).
' -----------------------------------------------------------------------------------
Private Sub Class_Initialize()
    m_strPathLogFile = ThisWorkbook.Path & "\"
    m_strLogSeparator = String(100, "=") & vbNewLine
    If m_lngMaxLogSizeBeforeSave = 0 Then MaxLogSizeBeforeSave = 5000
    m_strLogFileFullName = m_strPathLogFile & DEFAULT_LOG_FILENAME
End Sub

' -----------------------------------------------------------------------------------
' Function  : Class_Terminate
' Purpose   : Class teardown; clears the in-memory buffer.
'
' Parameters:
'   (none)
'
' Returns   :
'
' Notes     :
' -----------------------------------------------------------------------------------
Private Sub Class_Terminate()
    m_strBuffer = vbNullString
End Sub

' -----------------------------------------------------------------------------------
' Function  : LogFilename (Property Let)
' Purpose   : Sets the target log file name (without extension); extension enforced
'             to DEFAULT_LOG_EXT and combined with the current path.
'
' Parameters:
'   strLogFilename [String] - File name without extension; empty = DEFAULT_LOG_FILENAME.
'
' Returns   :
'
' Notes     :
'   - Uses the class' current m_strPathLogFile as base folder.
' -----------------------------------------------------------------------------------
Public Property Let LogFilename(ByVal strLogFilename As String)
    If strLogFilename = vbNullString Then
        m_strLogFileFullName = m_strPathLogFile & DEFAULT_LOG_FILENAME
    Else
        m_strLogFileFullName = m_strPathLogFile & strLogFilename & DEFAULT_LOG_EXT
    End If
End Property

' -----------------------------------------------------------------------------------
' Function  : PathLogFile (Property Let)
' Purpose   : Sets the base folder for log files and creates it if it does not exist.
'
' Parameters:
'   strNewValue [String] - Folder path (trailing backslash optional).
'
' Returns   :
'
' Notes     :
'   - Uses Dir to detect; calls MkDir to create.
' -----------------------------------------------------------------------------------
Public Property Let PathLogFile(ByVal strNewValue As String)

    m_strPathLogFile = strNewValue
    If Dir(m_strPathLogFile, vbDirectory) = "" Then MkDir (m_strPathLogFile)

End Property

' -----------------------------------------------------------------------------------
' Function  : Buffer (Property Get)
' Purpose   : Returns the in-memory log buffer.
'
' Parameters:
'   (none)
'
' Returns   : String - Current buffer content.
'
' Notes     :
' -----------------------------------------------------------------------------------
Public Property Get Buffer() As String

    Buffer = m_strBuffer

End Property

' -----------------------------------------------------------------------------------
' Function  : Buffer (Property Let)
' Purpose   : Replaces the in-memory log buffer content.
'
' Parameters:
'   strNewValue [String] - New buffer text.
'
' Returns   :
'
' Notes     :
' -----------------------------------------------------------------------------------
Public Property Let Buffer(ByVal NewValue As String)

    m_strBuffer = NewValue

End Property

' -----------------------------------------------------------------------------------
' Function  : maxLogSizeBeforeSave (Property Get)
' Purpose   : Returns the maximum buffer length before an automatic save is triggered.
'
' Parameters:
'   (none)
'
' Returns   : Long - Current threshold in characters.
'
' Notes     :
' -----------------------------------------------------------------------------------
Public Property Get MaxLogSizeBeforeSave() As Long

    MaxLogSizeBeforeSave = m_lngMaxLogSizeBeforeSave

End Property

' -----------------------------------------------------------------------------------
' Function  : maxLogSizeBeforeSave (Property Let)
' Purpose   : Sets the maximum buffer length before an automatic save is triggered.
'
' Parameters:
'   lngNewValue [Long] - Threshold in characters.
'
' Returns   :
'
' Notes     :
' -----------------------------------------------------------------------------------
Public Property Let MaxLogSizeBeforeSave(ByVal NewValue As Long)

    m_lngMaxLogSizeBeforeSave = NewValue

End Property

' -----------------------------------------------------------------------------------
' Function  : LogSeparator (Property Get)
' Purpose   : Returns the separator string used around log entries.
'
' Parameters:
'   (none)
'
' Returns   : String - The current separator text.
'
' Notes     :
' -----------------------------------------------------------------------------------
Public Property Get LogSeparator() As String

    LogSeparator = m_strLogSeparator

End Property

' -----------------------------------------------------------------------------------
' Function  : LogSeparator (Property Let)
' Purpose   : Sets the separator string used around log entries.
'
' Parameters:
'   strNewValue [String] - New separator text (e.g., "====" & vbCrLf).
'
' Returns   :
'
' Notes     :
' -----------------------------------------------------------------------------------
Public Property Let LogSeparator(ByVal NewValue As String)

    m_strLogSeparator = NewValue

End Property
' -----------------------------------------------------------------------------------
' Function  : WriteError
' Purpose   : Writes a structured error record into the log buffer, including the
'             procedure name, error number, line number, and description.
'
' Parameters:
'   strNameFunction [String]  - Procedure/function name (optional).
'   lngErrNum       [Long]    - Error number (optional; -9999 means "none").
'   lngErrLine      [Integer] - Line number (optional; -1 means "none").
'
' Returns   :
'
' Notes     :
'   - Raises Err internally to retrieve its Description text (if given).
'   - Surrounds the entry with separators (before and after).
' -----------------------------------------------------------------------------------
Friend Sub WriteError(Optional ByVal strNameFunction As String = vbNullString, Optional ByVal lngErrNum As Long = -9999, Optional ByVal lngErrLine As Integer = -1)
    Dim strMsg As String
    Dim sErrDescription As String
    If strNameFunction = vbNullString Then strNameFunction = "The procedure name is not set"
    strMsg = String(2, vbNewLine) & "Function/procedure name:" & strNameFunction & vbNewLine
    If lngErrNum <> -9999 Then strMsg = strMsg & "Error number:" & CStr(lngErrNum) & vbNewLine Else strMsg = strMsg & "Error number:-" & vbNewLine
    If lngErrLine <> -1 Then strMsg = strMsg & "Line number:" & CStr(lngErrLine) Else strMsg = strMsg & "Line number:-"
    If lngErrNum <> -9999 Then
        On Error Resume Next
        Err.Clear
        Err.Raise lngErrNum
        If Err.Number <> 0 Then sErrDescription = Err.Description
        On Error GoTo 0
    End If
    Call AddRecord(strMsg, vbTab & vbTab & "Description:" & sErrDescription, 0, LOG_SEPARATOR_BEFORE_AND_AFTER, True)
End Sub

' -----------------------------------------------------------------------------------
' Function  : SaveLog
' Purpose   : Appends the current buffer content to the log file and clears the buffer
'             if writing succeeded.
'
' Parameters:
'   (none)
'
' Returns   :
'
' Notes     :
'   - Uses AddIntoTXTfile; prints to Immediate Window on error.
' -----------------------------------------------------------------------------------
Friend Sub SaveLog()
    If AddIntoTXTfile(m_strLogFileFullName, m_strBuffer) Then
        m_strBuffer = Empty
    Else
        Debug.Print "Error writing the log to the file" & m_strLogFileFullName
    End If
End Sub

' -----------------------------------------------------------------------------------
' Function  : ClearLog
' Purpose   : Clears the in-memory log buffer and overwrites the log file with empty
'             content.
'
' Parameters:
'   (none)
'
' Returns   :
'
' Notes     :
'   - Calls SaveTXTfile with an empty string.
' -----------------------------------------------------------------------------------
Friend Sub ClearLog()
    m_strBuffer = vbNullString                                        ' Clearing the buffer
    SaveTXTfile m_strLogFileFullName, vbNullString                    ' Saving an empty file
End Sub

' -----------------------------------------------------------------------------------
' Function  : OpenLog
' Purpose   : Starts a logging session: optionally truncates the log file, clears
'             the buffer, and writes a "Starting Log Record..." entry.
'
' Parameters:
'   strAddInfo  [String]  - Optional extra info appended to the start entry.
'   blnNewLog   [Boolean] - If True, creates a new empty log file first.
'
' Returns   :
'
' Notes     :
' -----------------------------------------------------------------------------------
Friend Sub OpenLog(Optional strAddInfo As String = "", Optional ByVal blnNewLog As Boolean = False)
    m_strBuffer = vbNullString                                             ' Clearing the buffer
    If blnNewLog Then SaveTXTfile m_strLogFileFullName, vbNullString  ' Saving a new empty file
    AddRecord "Starting Log Record...", strAddInfo, , LOG_SEPARATOR_BEFORE, True ' start recording
End Sub

' -----------------------------------------------------------------------------------
' Function  : CloseLog
' Purpose   : Writes a closing entry ("Finishing Log Record...") and clears
'             the in-memory buffer.
'
' Parameters:
'   (none)
'
' Returns   :
'
' Notes     :
'   - Does not persist the buffer automatically; call SaveLog if required.
' -----------------------------------------------------------------------------------
Friend Sub CloseLog()
    AddRecord "Finishing Log Record...", , , LOG_SEPARATOR_AFTER, True    ' last entry
    m_strBuffer = vbNullString                                             ' clear the buffer
End Sub

' -----------------------------------------------------------------------------------
' Function  : ShowLog
' Purpose   : Displays the current log file either in the Immediate Window (debug)
'             or opens it with the associated application (Shell).
'
' Parameters:
'   blnShowAsDebugWindow [Boolean] - True to print to Immediate Window; otherwise shell open.
'
' Returns   :
'
' Notes     :
' -----------------------------------------------------------------------------------
Friend Sub ShowLog(Optional blnShowAsDebugWindow As Boolean = False)
    On Error Resume Next
    Dim strPath As String
    If blnShowAsDebugWindow Then
        Debug.Print ReadTXTfile(m_strLogFileFullName)
    Else
        strPath = Chr(34) & m_strLogFileFullName & Chr(34)
        CreateObject("wscript.shell").Run strPath
    End If
End Sub

' -----------------------------------------------------------------------------------
' Function  : AddRecord
' Purpose   : Adds a formatted log entry to the in-memory buffer, optionally placing
'             separators before/after and forcing an immediate save when thresholds
'             are met or requested.
'
' Parameters:
'   strEventX        [String]            - Main event text (left column).
'   strInfo          [String]            - Optional, Additional info (right column).
'   intLevel         [Integer]           - Optional, Indentation level (tabs before timestamp).
'   LogSeparatorType [LOG_SEPARATOR_TYPE]- Where to place separator lines.
'   blnForceSavingLog[Boolean]           - Optional, If True, calls SaveLog immediately.
'
' Returns   :
'
' Notes     :
'   - Pads and aligns text; timestamp format "YYYY-MM-DD HH:NN:SS".
'   - Auto-saves when buffer exceeds maxLogSizeBeforeSave.
' -----------------------------------------------------------------------------------
Friend Sub AddRecord(ByVal strEventX As String, Optional ByVal strInfo As String, _
    Optional ByVal intLevel As Integer = 0, _
    Optional ByVal LogSeparatorType As LOG_SEPARATOR_TYPE = LOG_SEPARATOR_NONE, _
    Optional ByVal blnForceSavingLog As Boolean = False)
    Dim strMsg    As String
    strMsg = String(160, " "): Mid(strMsg, 1) = Left(strEventX, 200): Mid(strMsg, 100) = strInfo
    strMsg = String(intLevel, vbTab) & Format(Now, "YYYY-MM-DD HH:NN:SS") & String(2, vbTab) & Trim(strMsg) & vbNewLine

    strMsg = IIf(LogSeparatorType = LOG_SEPARATOR_BEFORE Or LogSeparatorType = LOG_SEPARATOR_BEFORE_AND_AFTER, LogSeparator, "") & _
        strMsg & IIf(LogSeparatorType = LOG_SEPARATOR_AFTER Or LogSeparatorType = LOG_SEPARATOR_BEFORE_AND_AFTER, LogSeparator, "")

    m_strBuffer = m_strBuffer & strMsg
    If blnForceSavingLog Or (Len(m_strBuffer) > MaxLogSizeBeforeSave) Then SaveLog
End Sub

' -----------------------------------------------------------------------------------
' Function  : SaveLogAs
' Purpose   : Saves the current log (file content) as another file name (with default
'             extension), either overwriting or appending.
'
' Parameters:
'   strFileName  [String]           - Target file name (without path/extension).
'   blnOverwrite [Boolean]          - Optional, True = Overwrite; False = append.
'
' Returns   :
'
' Notes     :
'   - Reads from the current log file (m_strLogFileFullName).
'   - Ensures DEFAULT_LOG_EXT on the target file.
' -----------------------------------------------------------------------------------
Friend Sub SaveLogAs(ByVal strFileName As String, Optional ByVal blnOverwrite As Boolean = False)
    On Error Resume Next
    Dim strMsg    As String
    strFileName = m_strPathLogFile & strFileName
    If Not strFileName Like "*" & DEFAULT_LOG_EXT Then strFileName = strFileName & DEFAULT_LOG_EXT
    strMsg = ReadTXTfile(m_strLogFileFullName)
    If blnOverwrite Then
        SaveTXTfile strFileName, strMsg
    Else
        AddIntoTXTfile strFileName, strMsg
    End If
End Sub

' ================================================================================================================================
' Private Member
' ================================================================================================================================

' -----------------------------------------------------------------------------------
' Function  : ReadTXTfile
' Purpose   : Reads the entire content of a text file and returns it as a string.
'
' Parameters:
'   strFileName [String] - Full path to the file to read.
'
' Returns   : String - File content (may be empty on failure).
'
' Notes     :
'   - Creates a late-bound Scripting.FileSystemObject and TextStream.
' -----------------------------------------------------------------------------------
Private Function ReadTXTfile(ByVal strFileName As String) As String
    Dim objFSO    As Object
    Dim objTS     As Object
    Set objFSO = CreateObject("scripting.filesystemobject")
    Set objTS = objFSO.OpenTextFile(strFileName, 1, True): ReadTXTfile = objTS.ReadAll: objTS.Close
    Set objTS = Nothing: Set objFSO = Nothing
End Function

' -----------------------------------------------------------------------------------
' Function  : SaveTXTfile
' Purpose   : Overwrites a text file with the provided content.
'
' Parameters:
'   strFileName [String] - Full path to the file to write.
'   strMsg      [String] - Content to write.
'
' Returns   : Boolean - True if no error occurred; otherwise False.
'
' Notes     :
'   - Creates or overwrites the file.
' -----------------------------------------------------------------------------------
Private Function SaveTXTfile(ByVal strFileName As String, ByVal strMsg As String) As Boolean
    Dim objFSO    As Object
    Dim objTS     As Object
    On Error Resume Next: Err.Clear
    Set objFSO = CreateObject("scripting.filesystemobject")
    Set objTS = objFSO.CreateTextFile(strFileName, True)
    objTS.Write strMsg: objTS.Close
    SaveTXTfile = Err = 0
    Set objTS = Nothing: Set objFSO = Nothing
End Function

' -----------------------------------------------------------------------------------
' Function  : AddIntoTXTfile
' Purpose   : Appends content to a text file; creates the file if it does not exist.
'
' Parameters:
'   strFileName [String] - Full path to the file to append to.
'   strMsg      [String] - Content to append.
'
' Returns   : Boolean - True if no error occurred; otherwise False.
'
' Notes     :
'   - Opens the file with ForAppending mode (8).
' -----------------------------------------------------------------------------------
Private Function AddIntoTXTfile(ByVal strFileName As String, ByVal strMsg As String) As Boolean
    Dim objFSO    As Object
    Dim objTS     As Object
    On Error Resume Next: Err.Clear
    Set objFSO = CreateObject("scripting.filesystemobject")
    Set objTS = objFSO.OpenTextFile(strFileName, 8, True)
    objTS.Write strMsg: objTS.Close
    Set objTS = Nothing: Set objFSO = Nothing
    AddIntoTXTfile = Err = 0
End Function

