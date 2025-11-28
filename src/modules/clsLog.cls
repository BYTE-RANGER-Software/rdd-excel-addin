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
'   oLog.AddRecord "Initialized", "OK", 0, LST_BEFORE
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
Implements ILog

Private Const DEFAULT_LOG_EXT As String = ".log"
Private Const DEFAULT_LOG_FILENAME As String = "common" & DEFAULT_LOG_EXT

Private m_logFilePath As String          ' base folder for log files
Private m_logFileFullName As String

Private m_buffer As String
Private m_maxLogSizeBeforeSave As Long
Private m_logSeparator As String

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
    m_logFilePath = ThisWorkbook.path & "\"
    m_logSeparator = String(100, "=") & vbNewLine
    If m_maxLogSizeBeforeSave = 0 Then m_maxLogSizeBeforeSave = 5000
    m_logFileFullName = m_logFilePath & DEFAULT_LOG_FILENAME
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
    m_buffer = vbNullString
End Sub

' -----------------------------------------------------------------------------------
' Function  : LogFilename (Property Let)
' Purpose   : Sets the target log file name (without extension); extension enforced
'             to DEFAULT_LOG_EXT and combined with the current path.
'
' Parameters:
'   value [String] - File name without extension; empty = DEFAULT_LOG_FILENAME.
'
' Returns   :
'
' Notes     :
'   - Uses the class' current m_LogFilePath as base folder.
' -----------------------------------------------------------------------------------
Private Property Let ILog_LogFileName(ByVal value As String)
    If value = vbNullString Then
        m_logFileFullName = m_logFilePath & DEFAULT_LOG_FILENAME
    Else
        m_logFileFullName = m_logFilePath & value & DEFAULT_LOG_EXT
    End If
End Property

' -----------------------------------------------------------------------------------
' Function  : PathLogFile (Property Let)
' Purpose   : Sets the base folder for log files and creates it if it does not exist.
'
' Parameters:
'   value [String] - Folder path (trailing backslash optional).
'
' Returns   :
'
' Notes     :
'   - Uses Dir to detect; calls MkDir to create.
' -----------------------------------------------------------------------------------
Private Property Let ILog_LogFilePath(ByVal value As String)

    m_logFilePath = value
    If Dir(m_logFilePath, vbDirectory) = "" Then MkDir (m_logFilePath)

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
Private Property Get ILog_Buffer() As String

    ILog_Buffer = m_buffer

End Property

' -----------------------------------------------------------------------------------
' Function  : Buffer (Property Let)
' Purpose   : Replaces the in-memory log buffer content.
'
' Parameters:
'   value [String] - New buffer text.
'
' Returns   :
'
' Notes     :
' -----------------------------------------------------------------------------------
Private Property Let ILog_Buffer(ByVal value As String)

    m_buffer = value

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
Private Property Get ILog_MaxLogSizeBeforeSave() As Long

    ILog_MaxLogSizeBeforeSave = m_maxLogSizeBeforeSave

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
Private Property Let ILog_MaxLogSizeBeforeSave(ByVal value As Long)

    m_maxLogSizeBeforeSave = value

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
Private Property Get ILog_LogSeparator() As String

    ILog_LogSeparator = m_logSeparator

End Property

' -----------------------------------------------------------------------------------
' Function  : LogSeparator (Property Let)
' Purpose   : Sets the separator string used around log entries.
'
' Parameters:
'   value [String] - New separator text (e.g., "====" & vbCrLf).
'
' Returns   :
'
' Notes     :
' -----------------------------------------------------------------------------------
Private Property Let ILog_LogSeparator(ByVal value As String)

    m_logSeparator = value

End Property
' -----------------------------------------------------------------------------------
' Function  : WriteError
' Purpose   : Writes a structured error record into the log buffer, including the
'             procedure name, error number, line number, and description.
'
' Parameters:
'   functionName [String]  - Procedure/function name (optional).
'   errorNumber  [Long]    - Error number (optional; -9999 means "none").
'   errorLine    [Integer] - Line number (optional; -1 means "none").
'
' Returns   :
'
' Notes     :
'   - Raises Err internally to retrieve its Description text (if given).
'   - Surrounds the entry with separators (before and after).
' -----------------------------------------------------------------------------------
Private Sub ILog_WriteError(Optional ByVal functionName As String = vbNullString, Optional ByVal errorNumber As Long = -9999, Optional ByVal errorLine As Long = -1)
    Dim eventText As String
    Dim infoText As String
    Dim errorDescription As String
    
    If functionName = vbNullString Then functionName = "The procedure name is not set"
    
    If errorNumber <> -9999 Then
        On Error Resume Next
        Err.Clear
        Err.Raise errorNumber
        If Err.Number <> 0 Then errorDescription = Err.Description
        On Error GoTo 0
    End If

    eventText = vbNewLine & "Function/procedure name:" & functionName & vbNewLine
    eventText = eventText & "Error number:" & IIf(errorLine <> -9999, CStr(errorNumber), "-") & vbNewLine
    
    eventText = eventText & "Line number:" & IIf(errorLine <> -1, CStr(errorLine), "-")
    
    infoText = String(2, vbTab) & "Description:" & errorDescription
    
    Call ILog_AddRecord(eventText, infoText, 0, LST_BEFORE_AND_AFTER, True)
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
Private Sub ILog_SaveLog()
    If AddIntoTXTfile(m_logFileFullName, m_buffer) Then
        m_buffer = Empty
    Else
        Debug.Print "Error writing the log to the file" & m_logFileFullName
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
Private Sub ILog_ClearLog()
    m_buffer = vbNullString                                        ' Clearing the buffer
    SaveTXTfile m_logFileFullName, vbNullString                    ' Saving an empty file
End Sub

' -----------------------------------------------------------------------------------
' Function  : OpenLog
' Purpose   : Starts a logging session: optionally truncates the log file, clears
'             the buffer, and writes a "Starting Log Record..." entry.
'
' Parameters:
'   additionalInfo  [String]  - Optional extra info appended to the start entry.
'   isNewLog   [Boolean] - If True, creates a new empty log file first.
'
' Returns   :
'
' Notes     :
' -----------------------------------------------------------------------------------
Private Sub ILog_OpenLog(Optional additionalInfo As String = vbNullString, Optional ByVal isNewLog As Boolean = False)
    m_buffer = vbNullString                                             ' Clearing the buffer
    If isNewLog Then SaveTXTfile m_logFileFullName, vbNullString  ' Saving a new empty file
    ILog_AddRecord IIf(additionalInfo <> vbNullString, "Starting Log Record..." & String(46 - Len(additionalInfo), " ") & additionalInfo, "Starting Log Record..."), , , LST_BEFORE, True ' start recording
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
Private Sub ILog_CloseLog()
    ILog_AddRecord "Finishing Log Record...", , , LST_AFTER, True    ' last entry
    m_buffer = vbNullString                                             ' clear the buffer
End Sub

' -----------------------------------------------------------------------------------
' Function  : ShowLog
' Purpose   : Displays the current log file either in the Immediate Window (debug)
'             or opens it with the associated application (Shell).
'
' Parameters:
'   showAsDebugWindow [Boolean] - True to print to Immediate Window; otherwise shell open.
'
' Returns   :
'
' Notes     :
' -----------------------------------------------------------------------------------
Private Sub ILog_ShowLog(Optional showAsDebugWindow As Boolean = False)
    On Error Resume Next
    Dim path As String
    If showAsDebugWindow Then
        Debug.Print ReadTXTfile(m_logFileFullName)
    Else
        path = Chr(34) & m_logFileFullName & Chr(34)
        CreateObject("wscript.shell").Run path
    End If
End Sub

' -----------------------------------------------------------------------------------
' Function  : AddRecord
' Purpose   : Adds a formatted log entry to the in-memory buffer, optionally placing
'             separators before/after and forcing an immediate save when thresholds
'             are met or requested.
'
' Parameters:
'   eventText        [String]            - Main event text (left column).
'   info             [String]            - Optional, Additional info (right column).
'   level            [Integer]           - Optional, Indentation level (tabs before timestamp).
'   LogSeparatorType [LogSeparatorType]- Where to place separator lines.
'   forceSave        [Boolean]           - Optional, If True, calls SaveLog immediately.
'
' Returns   :
'
' Notes     :
'   - Pads and aligns text; timestamp format "YYYY-MM-DD HH:NN:SS".
'   - Auto-saves when buffer exceeds maxLogSizeBeforeSave.
' -----------------------------------------------------------------------------------
Private Sub ILog_AddRecord(ByVal eventText As String, Optional ByVal info As String, _
    Optional ByVal level As Long = 0, _
    Optional ByVal LogSeparatorType As LogSeparatorType = LST_NONE, _
    Optional ByVal forceSave As Boolean = False)
    Dim Message As String
    Dim separator As String
    Dim wrappedInfo As String
    Dim infoStartPos As Long
    
    ' Add separator if needed
    If LogSeparatorType = LST_BEFORE Or LogSeparatorType = LST_BEFORE_AND_AFTER Then
        separator = m_logSeparator
    End If
    
    ' Calculate the start position for info (usually position 100)
    ' But after the timestamp and tabs, this position is reached
    infoStartPos = 36  ' Number of spaces for indentation at line breaks
    
    ' If info contains text and is longer than 60 characters, wrap with correct indentation
    If info <> vbNullString And Len(info) > 60 Then
        wrappedInfo = WrapTextWithIndent(info, 60, String(infoStartPos, " "))
    Else
        wrappedInfo = info
    End If
    
    ' base padding
    Message = String(160, " ")
    Mid(Message, 1) = Left(eventText, 99)
    Mid(Message, 100) = wrappedInfo
    
    ' add timestamp
    Message = String(level, vbTab) & Format(Now, "YYYY-MM-DD HH:NN:SS") & String(2, vbTab) & Trim(Message) & vbNewLine

    ' add trailing separator
    Message = separator & Message
    If LogSeparatorType = LST_AFTER Or LogSeparatorType = LST_BEFORE_AND_AFTER Then
        Message = Message & m_logSeparator
    End If

    m_buffer = m_buffer & Message
    If forceSave Or (Len(m_buffer) > m_maxLogSizeBeforeSave) Then ILog_SaveLog
End Sub

' -----------------------------------------------------------------------------------
' Function  : SaveLogAs
' Purpose   : Saves the current log (file content) as another file name (with default
'             extension), either overwriting or appending.
'
' Parameters:
'   fileName  [String]           - Target file name (without path/extension).
'   overwrite [Boolean]          - Optional, True = Overwrite; False = append.
'
' Returns   :
'
' Notes     :
'   - Reads from the current log file (m_logFileFullName).
'   - Ensures DEFAULT_LOG_EXT on the target file.
' -----------------------------------------------------------------------------------
Private Sub ILog_SaveLogAs(ByVal fileName As String, Optional ByVal overwrite As Boolean = False)
    On Error Resume Next
    Dim Message    As String
    fileName = m_logFilePath & fileName
    If Not fileName Like "*" & DEFAULT_LOG_EXT Then fileName = fileName & DEFAULT_LOG_EXT
    Message = ReadTXTfile(m_logFileFullName)
    If overwrite Then
        SaveTXTfile fileName, Message
    Else
        AddIntoTXTfile fileName, Message
    End If
End Sub

' ================================================================================================================================
' Private Member
' ================================================================================================================================

' -----------------------------------------------------------------------------------
' Function  : WrapTextWithIndent
' Purpose   : Wraps long text into multiple lines with proper indentation for
'             subsequent lines
'
' Parameters:
'   text       [String] - The text to wrap
'   maxWidth   [Long]   - Maximum characters per line
'   indent     [String] - Indentation string for wrapped lines
'
' Returns   : String - Wrapped text with line breaks and indentation
'
' Notes     :
'   - First line has no indent, subsequent lines are indented
'   - Tries to break at word boundaries (spaces) when possible
' -----------------------------------------------------------------------------------
Private Function WrapTextWithIndent(ByVal text As String, ByVal maxWidth As Long, ByVal indent As String) As String
    Dim result As String
    Dim currentLine As String
    Dim words() As String
    Dim i As Long
    Dim firstLine As Boolean
    
    ' Remove leading/trailing spaces
    text = Trim(text)
    
    ' If the text is shorter than maxWidth, return it directly.
    If Len(text) <= maxWidth Then
        WrapTextWithIndent = text
        Exit Function
    End If
    
    ' Split text into words
    words = Split(text, " ")
    firstLine = True
    currentLine = ""
    
    For i = LBound(words) To UBound(words)
        ' Check whether the word fits into the current line
        If Len(currentLine) + Len(words(i)) + 1 <= maxWidth Then
            ' Add word to current line
            If currentLine = "" Then
                currentLine = words(i)
            Else
                currentLine = currentLine & " " & words(i)
            End If
        Else
            ' Line is full, add it to the result
            If result = "" Then
                result = currentLine
                firstLine = False
            Else
                result = result & vbNewLine & indent & currentLine
            End If
            currentLine = words(i)
        End If
    Next i
    
    ' Add the last line
    If currentLine <> "" Then
        If result = "" Then
            result = currentLine
        Else
            result = result & vbNewLine & indent & currentLine
        End If
    End If
    
    WrapTextWithIndent = result
End Function

' -----------------------------------------------------------------------------------
' Function  : ReadTXTfile
' Purpose   : Reads the entire content of a text file and returns it as a string.
'
' Parameters:
'   fileName [String] - Full path to the file to read.
'
' Returns   : String - File content (may be empty on failure).
'
' Notes     :
'   - Creates a late-bound Scripting.FileSystemObject and TextStream.
' -----------------------------------------------------------------------------------
Private Function ReadTXTfile(ByVal filepath As String) As String
    Dim fso    As Object
    Dim textStream     As Object
    Set fso = CreateObject("scripting.filesystemobject")
    Set textStream = fso.OpenTextFile(filepath, 1, True): ReadTXTfile = textStream.ReadAll: textStream.Close
    Set textStream = Nothing: Set fso = Nothing
End Function

' -----------------------------------------------------------------------------------
' Function  : SaveTXTfile
' Purpose   : Overwrites a text file with the provided content.
'
' Parameters:
'   fileName [String] - Full path to the file to write.
'   Message      [String] - Content to write.
'
' Returns   : Boolean - True if no error occurred; otherwise False.
'
' Notes     :
'   - Creates or overwrites the file.
' -----------------------------------------------------------------------------------
Private Function SaveTXTfile(ByVal fileName As String, ByVal Message As String) As Boolean
    Dim fso    As Object
    Dim textStream     As Object
    On Error Resume Next: Err.Clear
    Set fso = CreateObject("scripting.filesystemobject")
    Set textStream = fso.CreateTextFile(fileName, True)
    textStream.Write Message: textStream.Close
    SaveTXTfile = Err = 0
    Set textStream = Nothing: Set fso = Nothing
End Function

' -----------------------------------------------------------------------------------
' Function  : AddIntoTXTfile
' Purpose   : Appends content to a text file; creates the file if it does not exist.
'
' Parameters:
'   fileName [String] - Full path to the file to append to.
'   Message      [String] - Content to append.
'
' Returns   : Boolean - True if no error occurred; otherwise False.
'
' Notes     :
'   - Opens the file with ForAppending mode (8).
' -----------------------------------------------------------------------------------
Private Function AddIntoTXTfile(ByVal fileName As String, ByVal Message As String) As Boolean
    Dim fso    As Object
    Dim textStream     As Object
    On Error Resume Next: Err.Clear
    Set fso = CreateObject("scripting.filesystemobject")
    Set textStream = fso.OpenTextFile(fileName, 8, True)
    textStream.Write Message: textStream.Close
    Set textStream = Nothing: Set fso = Nothing
    AddIntoTXTfile = Err = 0
End Function

