VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RDDAddInWkBk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'* * * * * * * * * * * * * * * * * * * *  VSC VBATools * * * * * * * * * * * * * * * * * * * *
'* Path       : C:\Users\thiemann\Documents\Programming\Games\AGS\indy-foy\tools\dev\rdd-excel-addin\src\_internal\Backup\
'* Version    : v2
'* Created    : 11.12.2025_14.04.03
'* OldVersion : v1
'* * * * * * * * * * * * * * * * * * * *  VSC VBATools * * * * * * * * * * * * * * * * * * * *
' ====================================================================================
' Class     : RDDAddInWkBk (ThisWorkbook)
' Purpose   : Encapsulates Add-In Workbook event handling and delegates to modMain.
'             This is the ThisWorkbook class for the Add-In, handling installation,
'             opening, closing, and worksheet events.
'
' Public API:
'   - (Event Handlers Only)
'     * Workbook_AddinInstall
'     * Workbook_Open
'     * Workbook_BeforeClose
'
' Dependencies:
'   - modMain : All event handling logic is delegated to modMain
'
' Notes     :
'   - This class should remain thin; all business logic lives in modMain
'   - SheetBeforeRightClick is for testing; covered by App_SheetBeforeRightClick
'     in add-in mode via clsAppEvents
'   - Event handlers print debug info for troubleshooting
'
' ====================================================================================
Option Explicit

Private Sub Workbook_AddinInstall()
    Debug.Print "Workbook_AddinInstall()"
    Call modMain.HandleAddInWorkbookInstall
End Sub

Private Sub Workbook_Open()
    Debug.Print "Workbook_Open()"
    modMain.HandleAddInWorkbookOpen
End Sub


Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Debug.Print "Workbook_BeforeClose()"
    modMain.HandleAddInWorkbookBeforeClose Cancel
End Sub

