Attribute VB_Name = "modLog"
' -----------------------------------------------------------------------------------
' Module    : modLog
' Purpose   : Provides shared logging types and constants used across the project,
'             in particular the LogSeparatorType enum for log record formatting.
'
' Remarks   :
'   - Central place for logging-related enums and constants.
'   - Referenced by the logging interface (ILog) and its implementations (e.g. clsLog).
'
' Dependencies :
'   - None
'
' -----------------------------------------------------------------------------------
Option Explicit

' Adds visual separators around log records (where to place them relative to the entry)
Public Enum LogSeparatorType
    LST_NONE = 0                 ' No separator line
    LST_BEFORE = 1               ' Insert separator before the entry
    LST_AFTER = 2                ' Insert separator after the entry
    LST_BEFORE_AND_AFTER = 3     ' Insert separator before and after the entry
End Enum



