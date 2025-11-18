Attribute VB_Name = "modFormDropCallbacks"
' -----------------------------------------------------------------------------------
' Module    : modFormDropCallbacks
' Purpose   : Hosts callback procedures for FormDrop category and sub dropdowns.
'             Called via Application.Run from clsFormDrop when a FormDrop selection
'             changes.
'
' Public API:
'   - OnFormDropCatSelected  : Category dropdown selection callback.
'   - OnFormDropSubSelected  : Sub dropdown selection callback.
'
' Dependencies:
'   - modErr     : Error reporting helper.
'   - modMain    : Provides AppProjectName for dialog captions.
'
' Notes     :
'   - Procedure names must match the callbacks configured in
'     clsFormDropManager.SetCallbacks.
' -----------------------------------------------------------------------------------

Option Explicit


' === callbacks for drop-downs ==========

' -----------------------------------------------------------------------------------
' Procedure : OnFormDropCatSelected
' Purpose   : Business logic for FormDrop category selection.
'             Called via Application.Run from clsFormDrop when user selects a category.
'
' Parameters:
'   ddCat   [DropDown] - The category dropdown control that triggered the event
'   cell    [Range]    - cell associated with the dropdown
'
' Returns   : (none)
'
' Notes     :
'   - This is the callback implementation for FormDrop category changes
'   - Method name must match the string passed to SetCallbacks in clsFormDropManager
' -----------------------------------------------------------------------------------
Public Sub OnFormDropCatSelected(ByVal ddCat As DropDown, cell As Range)
    On Error GoTo ErrHandler
    
    'TODO:
    
    Debug.Print "FormDrop Category selected: " & ddCat.List(ddCat.Value)
    
CleanExit:
    On Error GoTo 0
    Exit Sub
    
ErrHandler:
    modErr.ReportError "OnFormDropCatSelected", Err.Number, Erl, caption:=AppProjectName
    Resume CleanExit
End Sub


' -----------------------------------------------------------------------------------
' Procedure : OnFormDropSubSelected
' Purpose   : Business logic for FormDrop sub-item selection.
'             Called via Application.Run from clsFormDrop when user selects a sub-item.
'
' Parameters:
'   ddSub   [DropDown] - The sub dropdown control that triggered the event
'   cell    [Range]    - cell associated with the dropdown
' Returns   : (none)
'
' Notes     :
'   - This is the callback implementation for FormDrop sub-item changes
'   - Method name must match the string passed to SetCallbacks in clsFormDropManager
' -----------------------------------------------------------------------------------
Public Sub OnFormDropSubSelected(ByVal ddSub As DropDown, cell As Range)
    On Error GoTo ErrHandler

    
    ' TODO:
    
    Debug.Print "FormDrop Sub selected: " & ddSub.List(ddSub.Value)
    
CleanExit:
    On Error GoTo 0
    Exit Sub
    
ErrHandler:
    modErr.ReportError "OnFormDropSubSelected", Err.Number, Erl, caption:=AppProjectName
    Resume CleanExit
End Sub
