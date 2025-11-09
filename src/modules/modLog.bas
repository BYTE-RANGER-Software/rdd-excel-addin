Attribute VB_Name = "modLog"
' modLog.bas
Option Explicit

' Adds visual separators around log records (where to place them relative to the entry)
Public Enum LOG_SEPARATOR_TYPE
    LOG_SEPARATOR_NONE = 0                 ' No separator line
    LOG_SEPARATOR_BEFORE = 1               ' Insert separator before the entry
    LOG_SEPARATOR_AFTER = 2                ' Insert separator after the entry
    LOG_SEPARATOR_BEFORE_AND_AFTER = 3     ' Insert separator before and after the entry
End Enum


