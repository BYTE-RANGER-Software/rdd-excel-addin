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
Private m_wbActiveWorkbook As Workbook

Private Sub Workbook_AddinInstall()
'If the add-in is activated when a workbook is opened,
'save the reference to this workbook for Workbook_Open,
'since the add-in is set as the active workbook in Workbook_Open.
If Not ActiveWorkbook Is Nothing Then Set m_wbActiveWorkbook = ActiveWorkbook
End Sub

Private Sub Workbook_Open()
    modMain.AppStart
End Sub


Private Sub Workbook_BeforeClose(Cancel As Boolean)
    modMain.AppStop
End Sub

