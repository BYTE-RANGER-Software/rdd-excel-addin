Attribute VB_Name = "modPDCCallbacks"
' -----------------------------------------------------------------------------------
' Modules: modPDCCallbacks
' Purpose: Public dispatcher for OnAction callbacks from PDC shapes.
'          Forwards calls to the corresponding private modules.
'
' Pattern:
'   Shape.OnAction -> modCallbacks.OnPdcNodeClick -> (on strg + click) modPDC.NavigateToPuzzle
'
' Notes:
'   - This module must NOT have the “Private Module” option.
'   - OnAction can only call public subs in non-private modules
'   - strg + click detection via GetAsyncKeyState
'
' -----------------------------------------------------------------------------------

Option Explicit

' ===== Windows API ==================================================================
Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer


Private Const VK_CONTROL As Long = &H11

' -----------------------------------------------------------------------------------
' Procedure : OnPdcNodeClick
' Purpose   : Dispatcher for clicks on PDC node shapes.
'             Detects double clicks and then forwards to modPDC.NavigateToPuzzle.
'
' Parameters: (none) - Shape name is determined via Application.Caller
'
' Notes:
'   - Called by Shape.OnAction
'   - Application.Caller contains the shape name (= NodeID)
'   - Single click: No action (shape can be moved)
'   - Strg pressed + click: Navigation to the puzzle
' ---------------------- -------------------------------------------------------------
Public Sub OnPdcNodeClick()
    On Error GoTo ErrHandler

    Dim nodeID As String
    
    nodeID = Application.Caller
        
    If IsCtrlKeyDown() Then
        modPDC.NavigateToPuzzle nodeID
    End If
    
    Exit Sub
    
ErrHandler:
    modErr.ReportError "OnPdcNodeClick", Err.Number, Erl, caption:=modMain.AppProjectName
End Sub

' -----------------------------------------------------------------------------------
' Function: IsCtrlKeyDown
' Purpose: Checks whether the Ctrl key is pressed.
' Returns: Boolean - True if Ctrl is pressed
' ----------------------------------------------------- ------------------------------
Private Function IsCtrlKeyDown() As Boolean
    IsCtrlKeyDown = (GetAsyncKeyState(VK_CONTROL) And &H8000) <> 0
End Function
