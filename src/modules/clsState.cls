VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ====================================================================================
' Class     : clsState
' Purpose   : Centralized application state manager. Maintains runtime state that
'             needs to be shared across modules and coordinates UI updates.
'             Implements Singleton pattern (PredeclaredId = True) for global access.
'
' Public API:
'   === Properties (Get/Let) ===
'   - RoomSheetChanged            : Tracks whether Room sheet data has changed
'   - RoomsValidated              : Tracks whether room data validation has passed
'   - RoomsValidationIssueCount   : Number of validation issues found (0 = success)
'   - CellCtxMnuNeedsPrepare      : Indicates if context menu needs rebuilding
'   - CellCtxMenuType             : Stores resolved context menu type
'   - CellCtxMnuHideDefault       : Controls visibility of built-in context menu items
'   - RibbonUI                    : Reference to Excel Ribbon UI interface
'
'   === Methods ===
'   - InvalidateRibbon()       : Forces entire Ribbon UI to refresh
'   - InvalidateControl(id)    : Refreshes specific Ribbon control by ID
'   - Init()                   : One-time startup initialization
'   - Cleanup()                : Releases resources and resets state
'
' Usage     :
'   ' Access singleton instance directly (PredeclaredId = True)
'   clsState.RoomSheetChanged = True
'   clsState.InvalidateRibbon
'
'   ' Store Ribbon reference during Ribbon_OnLoad callback
'   Set clsState.RibbonUI = ribbon
'
' Dependencies:
'   - IRibbonUI : Excel Ribbon interface for UI updates
'   - Used by  : modMain, modCellCtxMnu, modRibbon, modRooms
'
' Notes     :
'   - This class holds ONLY runtime state, not business logic
'   - No persistence logic (use modOptions for that)
'   - No user preferences (use modOptions for that)
'   - No constants (use modConst for that)
'   - State is transient; reset via Cleanup() on shutdown
'   - Thread-safe not required (Excel VBA is single-threaded)
'
' Lifecycle :
'   1. Ribbon_OnLoad: Set RibbonUI reference
'   2. Workbook_Open: Call Init() if needed
'   3. During operation: Read/write state properties as needed
'   4. Workbook_BeforeClose: Call Cleanup() to release references
'
' ====================================================================================
Option Explicit

' Module-scope fields
Private m_roomSheetChanged As Boolean       ' Tracks whether Room sheet changed
Private m_roomsValidated As Boolean         ' Tracks whether room data validation has passed
Private m_roomsValidationIssueCount As Long ' Number of validation issues found (0 = success)
Private m_cellCtxMnuNeedsPrepare As Boolean ' Tracks context menu prepare need. True to (re)prepare the ctx menu on next call when Excel changed the "Cell" CommandBar
Private m_cellCtxMenuType As Integer        ' Stores context menu type. Resolved context menu type according to CellCtxMnu
Private m_cellCtxMnuHideDefault As Boolean  ' Hides default context entries. If True, built-in entries will be hidden during the next prepare pass
Private m_ribbonUI As IRibbonUI             ' Ribbon reference

' --- Properties ---
Public Property Get RoomSheetChanged() As Boolean
    RoomSheetChanged = m_roomSheetChanged
End Property

Public Property Let RoomSheetChanged(ByVal value As Boolean)
    m_roomSheetChanged = value
End Property

Public Property Get CellCtxMnuNeedsPrepare() As Boolean
    CellCtxMnuNeedsPrepare = m_cellCtxMnuNeedsPrepare
End Property

Public Property Let CellCtxMnuNeedsPrepare(ByVal value As Boolean)
    m_cellCtxMnuNeedsPrepare = value
End Property

Public Property Get CellCtxMenuType() As Integer
    CellCtxMenuType = m_cellCtxMenuType
End Property

Public Property Let CellCtxMenuType(ByVal value As Integer)
    m_cellCtxMenuType = value
End Property

Public Property Get CellCtxMnuHideDefault() As Boolean
    CellCtxMnuHideDefault = m_cellCtxMnuHideDefault
End Property

Public Property Let CellCtxMnuHideDefault(ByVal value As Boolean)
    m_cellCtxMnuHideDefault = value
End Property

Public Property Get RibbonUI() As IRibbonUI
    Set RibbonUI = m_ribbonUI
End Property

Public Property Set RibbonUI(ByVal obj As IRibbonUI)
    Set m_ribbonUI = obj
End Property

' --- Helpers ---
Public Sub InvalidateRibbon()
    On Error Resume Next
    If Not m_ribbonUI Is Nothing Then m_ribbonUI.Invalidate
    On Error GoTo 0
End Sub

Public Sub InvalidateControl(ByVal controlId As String)
    On Error Resume Next
    m_ribbonUI.InvalidateControl controlId
    On Error GoTo 0
End Sub

Public Sub Init()
    ' One-time startup logic if needed
End Sub

Public Sub Cleanup()
    ' Release object refs, reset flags
    Set m_ribbonUI = Nothing
    m_roomSheetChanged = False
    m_cellCtxMnuNeedsPrepare = False
    m_cellCtxMenuType = 0
    m_cellCtxMnuHideDefault = False
    m_roomsValidated = False
    m_roomsValidationIssueCount = 0
End Sub

Public Property Get RoomsValidated() As Boolean
    RoomsValidated = m_roomsValidated
End Property

Public Property Let RoomsValidated(ByVal value As Boolean)
    m_roomsValidated = value
End Property

Public Property Get RoomsValidationIssueCount() As Long
    RoomsValidationIssueCount = m_roomsValidationIssueCount
End Property

Public Property Let RoomsValidationIssueCount(ByVal value As Long)
    m_roomsValidationIssueCount = value
End Property
