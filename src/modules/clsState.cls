VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Class Module: clsState
Option Explicit

' Module-scope fields
Private m_blnRoomSheetChanged As Boolean       ' Tracks whether Room sheet changed
Private m_blnCellCtxMnuNeedsPrepare As Boolean ' Tracks context menu prepare need. True to (re)prepare the ctx menu on next call when Excel changed the "Cell" CommandBar
Private m_intCellCtxMenuType As Integer        ' Stores context menu type. Resolved context menu type according to CellCtxMnu
Private m_blnCellCtxMnuHideDefault As Boolean  ' Hides default context entries. If True, built-in entries will be hidden during the next prepare pass
Private m_objRibbonUI As IRibbonUI             ' Ribbon reference

' --- Properties ---
Public Property Get RoomSheetChanged() As Boolean
    RoomSheetChanged = m_blnRoomSheetChanged
End Property

Public Property Let RoomSheetChanged(ByVal blnValue As Boolean)
    m_blnRoomSheetChanged = blnValue
End Property

Public Property Get CellCtxMnuNeedsPrepare() As Boolean
    CellCtxMnuNeedsPrepare = m_blnCellCtxMnuNeedsPrepare
End Property

Public Property Let CellCtxMnuNeedsPrepare(ByVal blnValue As Boolean)
    m_blnCellCtxMnuNeedsPrepare = blnValue
End Property

Public Property Get CellCtxMenuType() As Integer
    CellCtxMenuType = m_intCellCtxMenuType
End Property

Public Property Let CellCtxMenuType(ByVal intValue As Integer)
    m_intCellCtxMenuType = intValue
End Property

Public Property Get CellCtxMnuHideDefault() As Boolean
    CellCtxMnuHideDefault = m_blnCellCtxMnuHideDefault
End Property

Public Property Let CellCtxMnuHideDefault(ByVal blnValue As Boolean)
    m_blnCellCtxMnuHideDefault = blnValue
End Property

Public Property Get RibbonUI() As IRibbonUI
    Set RibbonUI = m_objRibbonUI
End Property

Public Property Set RibbonUI(ByVal obj As IRibbonUI)
    Set m_objRibbonUI = obj
End Property

' --- Helpers ---
Public Sub InvalidateRibbon()
    On Error Resume Next
    If Not m_objRibbonUI Is Nothing Then m_objRibbonUI.Invalidate
    On Error GoTo 0
End Sub

Public Sub InvalidateControl(ByVal controlId As String)
    On Error Resume Next
    m_objRibbonUI.InvalidateControl controlId
    On Error GoTo 0
End Sub

Public Sub Init()
    ' One-time startup logic if needed
End Sub

Public Sub Cleanup()
    ' Release object refs, reset flags
    Set m_objRibbonUI = Nothing
    m_blnRoomSheetChanged = False
    m_blnCellCtxMnuNeedsPrepare = False
    m_intCellCtxMenuType = 0
    m_blnCellCtxMnuHideDefault = False
End Sub

