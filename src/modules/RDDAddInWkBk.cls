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


Private Sub Workbook_AddinInstall()
    Debug.Print "Workbook_AddinInstall()"
    modMain.AppInstall
End Sub

Private Sub Workbook_Open()
    Debug.Print "Workbook_Open()"
    modMain.AppStart
End Sub


Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Debug.Print "Workbook_BeforeClose()"
    modMain.AppStop
End Sub

Private Sub Workbook_SheetBeforeRightClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    clsState.CellCtxMnuNeedsPrepare = True

    EvaluateCellCtxMenu Sh, Target
  
    clsState.InvalidateControl "RB75dd2c44_btnDynCtxMnu1"
    clsState.InvalidateControl "RB75dd2c44_btnDynCtxMnu2"
    
End Sub
