VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RDDAddInWkBk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private mwbActiveWorkbook

Private Sub Workbook_AddinInstall()
'If the add-in is activated when a workbook is opened,
'save the reference to this workbook for Workbook_Open,
'since the add-in is set as the active workbook in Workbook_Open.
If Not ActiveWorkbook Is Nothing Then Set mwbActiveWorkbook = ActiveWorkbook
End Sub

Private Sub Workbook_Open()
    Call ConnectEventHandler
    If Not modRibbon.gRibbonUI Is Nothing Then gRibbonUI.Invalidate
End Sub


Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Call DisconnectEventHandler
    If Not modRibbon.gRibbonUI Is Nothing Then gRibbonUI.Invalidate
End Sub

