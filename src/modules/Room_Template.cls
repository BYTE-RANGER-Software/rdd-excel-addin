VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Room_Template"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' ====================================================================================
' Class     : Room_Template
' Purpose   : Template worksheet class for room sheets. Serves as the code-behind
'             for the Room_Template worksheet which is used as a blueprint when
'             creating new room sheets.
'
' Public API:
'   - (None - worksheet template class)
'
' Dependencies:
'   - Used by modRooms.AddRoom to create new room instances
'
' Notes     :
'   - This is a worksheet class module
'   - DO NOT DELETE OR RENAME THIS SHEET
'   - The physical sheet serves as template; this module can contain
'     worksheet-specific event handlers if needed
'
' ====================================================================================
Option Explicit

