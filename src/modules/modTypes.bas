Attribute VB_Name = "modTypes"
Option Explicit
Option Private Module

Public Enum CellCtxMnu
    CCM_Default      ' Default menu behavior
    CCM_Rooms        ' Context menu for cells validated against room IDs
    CCM_Objects      ' Context menu for cells validated against objects list
    CCM_Actors       ' Context menu for cells validated against actors list
End Enum

' Match modes for sheet name selection
Public Enum SheetNameMatchMode
    SNMM_Exact = 0      ' exact match
    SNMM_Prefix = 1     ' starts with pattern
    SNMM_Suffix = 2     ' ends with pattern
    SNMM_Contains = 3   ' substring
    SNMM_Wildcard = 4   ' VBA Like pattern, e.g., "Room*"
    SNMM_Regex = 5      ' VBScript.RegExp (optional)
End Enum
