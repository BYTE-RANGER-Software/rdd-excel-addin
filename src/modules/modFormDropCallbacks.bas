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
    
    Debug.Print "FormDrop Category selected: " & ddCat.list(ddCat.value)
    
CleanExit:
    On Error GoTo 0
    Exit Sub
    
ErrHandler:
    modErr.ReportError "OnFormDropCatSelected", Err.Number, Erl, caption:=AppProjectName
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

    Dim selectedIndex As Integer
    Dim selectedValue As String
    Dim arrVals() As String
    Dim result As String
    Dim i As Integer
    Dim found As Boolean
    Dim ownerSheet As Worksheet: Set ownerSheet = ddSub.Parent
    ' TODO:
    
    'Debug.Print "FormDrop Sub selected: " & ddSub.List(ddSub.Value)

    selectedIndex = ddSub.ListIndex

    ' Wenn nichts ausgewählt wurde, abbrechen
    If selectedIndex = 0 Then Exit Sub

    selectedValue = ddSub.list(selectedIndex)

    ' Vorhandene Werte aufteilen
    If cell.value <> "" Then
        arrVals = Split(cell.value, ", ")
    Else
        ReDim arrVals(0)
        arrVals(0) = ""
    End If

    ' Prüfen, ob Wert bereits vorhanden ist
    result = ""
    found = False
    For i = LBound(arrVals) To UBound(arrVals)
        If arrVals(i) = selectedValue Then
            found = True ' Wert wird entfernt
        Else
            If arrVals(i) <> "" Then
                If result = "" Then
                    result = arrVals(i)
                Else
                    result = result & ", " & arrVals(i)
                End If
            End If
        End If
    Next i

    ' Wenn nicht vorhanden, hinzufügen
    If Not found Then
        If result = "" Then
            result = selectedValue
        Else
            result = result & ", " & selectedValue
        End If
    End If

    cell.value = result
    
    'Reset DropDown value, so that the same selection triggers OnAction again next time
    If ddSub.ListCount > 0 Then ddSub.value = selectedIndex - 1
        
CleanExit:
    On Error GoTo 0
    Exit Sub
    
ErrHandler:
    modErr.ReportError "OnFormDropSubSelected", Err.Number, Erl, caption:=AppProjectName
End Sub
