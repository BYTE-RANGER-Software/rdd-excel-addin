VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' -----------------------------------------------------------------------------------
' Module    : clsAppEvents
' Purpose   : Encapsulates application-level event handling and delegates to modMain.
'
' Public API:
'   - App [WithEvents Application] : Exposes Excel application events to this class.
'
' Dependencies:
'   - modMain, clsState, modTags, modCellCtxMnu, modOptions
'
' Notes     :
'   - This class should not contain business logic; it only forwards events.
' -----------------------------------------------------------------------------------

Option Explicit

Public WithEvents App As Application
Attribute App.VB_VarHelpID = -1

' -----------------------------------------------------------------------------------
' Procedure : App_SheetActivate
' Purpose   : Handles the Application.SheetActivate event and delegates to
'             modMain.HandleSheetActivate for non-add-in workbooks.
'
' Parameters:
'   sh   [Object] - Activated sheet object.
'
' Returns   : (none)
'
' Notes     :
'   - Only reacts if the sheet is not in the add-in workbook.
' -----------------------------------------------------------------------------------
Private Sub App_SheetActivate(ByVal Sh As Object)
   
    If IsSheetFromNonAddIn(Sh) Then
        
        modMain.HandleSheetActivate Sh

    End If
    
End Sub

' -----------------------------------------------------------------------------------
' Procedure : App_SheetChange
' Purpose   : Handles the Application.SheetChange event and delegates to
'             modMain.HandleSheetChange for non-add-in workbooks.
'
' Parameters:
'   sh       [Object] - Sheet object where the change occurred.
'   target   [Range]  - Changed cell or range.
'
' Returns   : (none)
'
' Notes     :
'   - Only reacts if the sheet is not in the add-in workbook.
' -----------------------------------------------------------------------------------
Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)

    If IsSheetFromNonAddIn(Sh) Then
        modMain.HandleSheetChange Sh, Target
    End If
End Sub

' -----------------------------------------------------------------------------------
' Procedure : App_SheetBeforeRightClick
' Purpose   : Handles the Application.SheetBeforeRightClick event and delegates to
'             modMain.HandleSheetBeforeRightClick.
'
' Parameters:
'   sh            [Object] - Sheet object where the right-click occurs.
'   target        [Range]  - Target cell or range.
'   Cancel        [Boolean] - Allows canceling the context menu (ByRef).
'
' Returns   : (none)
'
' Notes     :
'   - Only reacts if the sheet is not in the add-in workbook.
' -----------------------------------------------------------------------------------
Private Sub App_SheetBeforeRightClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    If IsSheetFromNonAddIn(Sh) Then
        modMain.HandleSheetBeforeRightClick Sh, Target, Cancel
    End If
End Sub

Private Sub App_SheetDeactivate(ByVal Sh As Object)
    If IsSheetFromNonAddIn(Sh) Then
    
        modMain.HandleSheetDeactivate Sh
    End If
End Sub

Private Sub App_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    If IsSheetFromNonAddIn(Sh) Then
        modMain.HandleSheetSelectionChange Sh, Target
    End If
End Sub

' -----------------------------------------------------------------------------------
' Procedure : App_WorkbookBeforeSave
' Purpose   : Handles the Application.WorkbookBeforeSave event and delegates to
'             modMain.HandleWorkbookBeforeSave for non-add-in workbooks.
'
' Parameters:
'   wb               [Workbook] - Workbook that is about to be saved.
'   SaveAsUi         [Boolean]  - Indicates whether the Save As UI is displayed.
'   Cancel           [Boolean]  - Allows canceling the save operation (ByRef).
'
' Returns   : (none)
'
' Notes     :
'   - Does not handle the add-in workbook itself.
' -----------------------------------------------------------------------------------
Private Sub App_WorkbookBeforeSave(ByVal wb As Workbook, ByVal SaveAsUI As Boolean, Cancel As Boolean)
    If IsNonAddInWorkbook(wb) Then

        modMain.HandleWorkbookBeforeSave wb, SaveAsUI, Cancel

    End If
End Sub

' -----------------------------------------------------------------------------------
' Function  : IsSheetFromNonAddIn
' Purpose   : Checks whether the given Worksheet belongs to a Workbook
'             that is neither the Add-In Workbook nor marked as an Add-In.
'
' Parameters:
'   sh              [Worksheet]  - Worksheet to be checked
'   returnSrcBook   [Workbook]   - (Optional , ByRef) Workbook reference (only set if function returns True)
'
' Returns   : Boolean - True if business logic may be applied; raises error otherwise.
' -----------------------------------------------------------------------------------
Private Function IsSheetFromNonAddIn(ByVal Sh As Worksheet, Optional ByRef returnSrcBook As Workbook) As Boolean

    Set returnSrcBook = Sh.Parent
    

    IsSheetFromNonAddIn = (Not (returnSrcBook Is RDDAddInWkBk) And Not returnSrcBook.IsAddin)

End Function

' -----------------------------------------------------------------------------------
' Function  : IsNonAddInWorkbook
' Purpose   : Checks whether the given Workbook is neither an Add-In Workbook nor marked as an Add-In.
'
' Parameters:
'   wb   [Workbook]   - Workbook reference
'
' Returns   : Boolean - True if business logic may be applied; raises error otherwise.
' -----------------------------------------------------------------------------------
Private Function IsNonAddInWorkbook(ByVal wb As Workbook) As Boolean
    

    IsNonAddInWorkbook = (Not (wb Is RDDAddInWkBk) And Not wb.IsAddin)

End Function
