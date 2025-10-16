VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Hinzufügen von Trennlinien zum Protokoll
Public Enum LOG_SEPARATOR_TYPE
    LOG_SEPARATOR_NONE = 0
    LOG_SEPARATOR_BEFORE = 1
    LOG_SEPARATOR_AFTER = 2
    LOG_SEPARATOR_BEFORE_AND_AFTER = 3
End Enum

Private Const DefaultTypeExtenFile As String = ".log"
Private Const DefaultNameLogFile As String = "common" & DefaultTypeExtenFile

Private m_strPathLogFile As String
Private m_strLogFileFullName As String

Public Buffer  As String
Public maxLogSizeBeforeSave As Long
Public LogSeparator As String

Public Property Let LogFilename(ByVal LogFilename As String)
    If LogFilename = vbNullString Then
        m_strLogFileFullName = m_strPathLogFile & DefaultNameLogFile
    Else
        m_strLogFileFullName = m_strPathLogFile & LogFilename & DefaultTypeExtenFile
    End If
End Property

Friend Sub WriteErrorLog(Optional ByVal sNameFunction As String = vbNullString, Optional ByVal iErrNum As Long = -9999, Optional ByVal iErrLine As Integer = -1)
    Dim txt    As String
    Dim sErrDescription As String
    If sNameFunction = vbNullString Then sNameFunction = "The procedure name is not set"
    txt = String(2, vbNewLine) & "Function/procedure name:" & sNameFunction & vbNewLine
    If iErrNum <> -9999 Then txt = txt & "Error number:" & CStr(iErrNum) & vbNewLine Else txt = txt & "Error number:-" & vbNewLine
    If iErrLine <> -1 Then txt = txt & "Line number:" & CStr(iErrLine) Else txt = txt & "Line number:-"
    If iErrNum <> -9999 Then
        On Error Resume Next
        Err.Clear
        Err.Raise iErrNum
        If Err.Number <> 0 Then sErrDescription = Err.Description
        On Error GoTo 0
    End If
    Call AddRecord(txt, vbTab & vbTab & "Description:" & sErrDescription, 0, LOG_SEPARATOR_BEFORE_AND_AFTER, True)
End Sub

Friend Sub SaveLog()
    If AddIntoTXTfile(m_strLogFileFullName, Buffer) Then
        Buffer = Empty
    Else
        Debug.Print "Error writing the log to the file" & m_strLogFileFullName
    End If
End Sub

Friend Sub ClearLog()
    Buffer = vbNullString                                        ' Löschen des Puffers
    SaveTXTfile m_strLogFileFullName, vbNullString                    ' Speichern einer leeren Datei
End Sub

Friend Sub OpenLog(Optional strAddInfo As String = "", Optional ByVal blnNewLog As Boolean = False)
    Buffer = vbNullString                                             ' Löschen des Puffers
    If blnNewLog Then SaveTXTfile m_strLogFileFullName, vbNullString  ' Speichern eine neue leeren Datei
    AddRecord "Starting Log Record...", strAddInfo, , LOG_SEPARATOR_BEFORE, True ' Aufzeichnung starten
End Sub

Friend Sub CloseLog()
    AddRecord "Finishing Log Record...", , , LOG_SEPARATOR_AFTER, True    ' Letzter Eintrag
    Buffer = vbNullString                                             ' Löschen des Puffers
End Sub

Friend Sub ShowLog(Optional ShowAsDebugWindow As Boolean = False)
    On Error Resume Next
    Dim sPath$
    If ShowAsDebugWindow Then
        Debug.Print ReadTXTfile(m_strLogFileFullName)
    Else
        sPath$ = Chr(34) & m_strLogFileFullName & Chr(34)
        CreateObject("wscript.shell").Run sPath$
    End If
End Sub

Friend Sub AddRecord(ByVal EventX As String, Optional ByVal Info As String, _
    Optional ByVal Level As Integer = 0, _
    Optional ByVal LogSeparatorType As LOG_SEPARATOR_TYPE = LOG_SEPARATOR_NONE, _
    Optional ByVal ForceSavingLog As Boolean = False)
    Dim txt    As String
    txt = String(160, " "): Mid(txt, 1) = Left(EventX, 200): Mid(txt, 100) = Info
    txt = String(Level, vbTab) & Format(Now, "YYYY-MM-DD HH:NN:SS") & String(2, vbTab) & Trim(txt) & vbNewLine

    txt = IIf(LogSeparatorType = LOG_SEPARATOR_BEFORE Or LogSeparatorType = LOG_SEPARATOR_BEFORE_AND_AFTER, LogSeparator, "") & _
        txt & IIf(LogSeparatorType = LOG_SEPARATOR_AFTER Or LogSeparatorType = LOG_SEPARATOR_BEFORE_AND_AFTER, LogSeparator, "")

    Buffer = Buffer & txt
    If ForceSavingLog Or (Len(Buffer) > maxLogSizeBeforeSave) Then SaveLog
End Sub


Friend Sub SaveLogAs(ByVal sFileName As String, Optional ByVal Overwrite As Boolean = False)
    On Error Resume Next
    Dim txt    As String
    sFileName = m_strPathLogFile & sFileName
    If Not sFileName Like "*" & DefaultTypeExtenFile Then sFileName = sFileName & DefaultTypeExtenFile
    txt = ReadTXTfile(m_strLogFileFullName)
    If Overwrite Then
        SaveTXTfile sFileName, txt
    Else
        AddIntoTXTfile sFileName, txt
    End If
End Sub

' ================================================================================================================================
Private Sub Class_Terminate()
    Buffer = vbNullString    'CloseLog
End Sub

Private Sub Class_Initialize()
    m_strPathLogFile = ThisWorkbook.Path & "\"
    LogSeparator = String(100, "=") & vbNewLine
    If maxLogSizeBeforeSave = 0 Then maxLogSizeBeforeSave = 5000
    m_strLogFileFullName = m_strPathLogFile & DefaultNameLogFile
End Sub

Private Function ReadTXTfile(ByVal sFileName As String) As String
    Dim FSO    As Object
    Dim ts     As Object
    Set FSO = CreateObject("scripting.filesystemobject")
    Set ts = FSO.OpenTextFile(sFileName, 1, True): ReadTXTfile = ts.ReadAll: ts.Close
    Set ts = Nothing: Set FSO = Nothing
End Function

Private Function SaveTXTfile(ByVal sFileName As String, ByVal txt As String) As Boolean
    Dim FSO    As Object
    Dim ts     As Object
    On Error Resume Next: Err.Clear
    Set FSO = CreateObject("scripting.filesystemobject")
    Set ts = FSO.CreateTextFile(sFileName, True)
    ts.Write txt: ts.Close
    SaveTXTfile = Err = 0
    Set ts = Nothing: Set FSO = Nothing
End Function

Private Function AddIntoTXTfile(ByVal sFileName As String, ByVal txt As String) As Boolean
    Dim FSO    As Object
    Dim ts     As Object
    On Error Resume Next: Err.Clear
    Set FSO = CreateObject("scripting.filesystemobject")
    Set ts = FSO.OpenTextFile(sFileName, 8, True)
    ts.Write txt: ts.Close
    Set ts = Nothing: Set FSO = Nothing
    AddIntoTXTfile = Err = 0
End Function

Public Property Let PathLogFile(ByVal strNewValue As String)

    m_strPathLogFile = strNewValue
    If Dir(m_strPathLogFile, vbDirectory) = "" Then MkDir (m_strPathLogFile)

End Property
