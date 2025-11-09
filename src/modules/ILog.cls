VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ILog.cls — Interface for logging service (implemented by clsLog)
' Purpose: Defines the logging contract for file path/name, buffer handling,
'          size limit, separators, and write/display operations.

'========================
' Properties
'========================
Public Property Let LogFilePath(ByVal value As String)
End Property

Public Property Let LogFileName(ByVal value As String)
End Property

Public Property Get Buffer() As String
End Property

Public Property Let Buffer(ByVal value As String)
End Property

Public Property Get MaxLogSizeBeforeSave() As Long
End Property

Public Property Let MaxLogSizeBeforeSave(ByVal value As Long)
End Property

Public Property Get LogSeparator() As String
End Property

Public Property Let LogSeparator(ByVal value As String)
End Property

'========================
' Operations
'========================
Public Sub OpenLog(Optional additionalInfo As String = vbNullString, Optional ByVal isNewLog As Boolean = False)
End Sub

Public Sub CloseLog()
End Sub

Public Sub SaveLog()
End Sub

Public Sub ClearLog()
End Sub

Public Sub ShowLog(Optional showAsDebugWindow As Boolean = False)
End Sub

Public Sub AddRecord(ByVal eventText As String, Optional ByVal info As String, _
    Optional ByVal level As Long = 0, _
    Optional ByVal logSeparatorType As LOG_SEPARATOR_TYPE = LOG_SEPARATOR_NONE, _
    Optional ByVal forceSave As Boolean = False)
End Sub

Public Sub SaveLogAs(ByVal fileName As String, Optional ByVal overwrite As Boolean = False)
End Sub

Public Sub WriteError(Optional ByVal functionName As String = vbNullString, Optional ByVal errorNumber As Long = -9999, Optional ByVal errorLine As Long = -1)
End Sub


