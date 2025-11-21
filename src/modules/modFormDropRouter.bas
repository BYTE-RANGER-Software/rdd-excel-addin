Attribute VB_Name = "modFormDropRouter"
' -----------------------------------------------------------------------------------
' Module    : modFormDropRouter
' Purpose   : Routes Excel form control events (DropDown and arrow Shape) to
'             clsFormDrop (clsFormDrop) instances via a shared registry.
'
' Public API:
'   - g_formDropRegistryDict  : Global registry mapping instance id to clsFormDrop instances.
'   - FormDrop_OnAction      : Macro called by Form Control DropDown.OnAction.
'   - FormDrop_Arrow_OnClick : Macro called by arrow Shape.OnAction.
'
' Dependencies:
'   - clsFormDrop : Event target class handling the form drop logic.
'
' Notes     :
'   - modFormDropRouter and clsFormDrop belong together and are mutually dependent.
'   - Keep this module thin and limited to the event routing of the Form Control DropDown and class registry access.
' -----------------------------------------------------------------------------------
Option Explicit

' ===== Public State ================================================================
' Global clsFormDrop instance registry: instance id -> class instance (clsFormDrop).
' (only reference) = real dict is hold in clsFormDropManager
Public g_formDropRegistryDict As Scripting.Dictionary

' ===== Public API ==================================================================
' Entry points called from form control events (DropDown/Shape OnAction).

' -----------------------------------------------------------------------------------
' Procedure : FormDrop_OnAction
' Purpose   : Routes DropDown.OnAction calls to the corresponding clsFormDrop instance
'             based on the encoded instance id in the control name.
'
' Parameters:
'   (none)
'
' Returns   : (none)
'
' Notes     :
'   - Called by the Form Control DropDown via its OnAction property.
'   - Exits silently if the instance id or registry entry cannot be resolved.
' -----------------------------------------------------------------------------------
Public Sub FormDrop_OnAction()
    On Error GoTo ErrHandler
    
    Dim callerCtrlName As String: callerCtrlName = CStr(Application.Caller)

    Dim instanceId As String: instanceId = clsFormDrop.ParseInstanceId(callerCtrlName)
    If Len(instanceId) = 0 Then Exit Sub

    If g_formDropRegistryDict Is Nothing Then Exit Sub
    If Not g_formDropRegistryDict.Exists(instanceId) Then Exit Sub

    Dim FormDropInst As clsFormDrop
    Set FormDropInst = g_formDropRegistryDict(instanceId)
    FormDropInst.HandleOnAction callerCtrlName
    
CleanExit:
    Exit Sub
    
ErrHandler:
    modErr.ReportError "FormDrop_OnAction", Err.Number, Erl
    Resume CleanExit
End Sub

' -----------------------------------------------------------------------------------
' Procedure : FormDrop_Arrow_OnClick
' Purpose   : Routes arrow Shape.OnAction calls to the corresponding clsFormDrop
'             instance based on the encoded instance id in the shape name.
'
' Parameters:
'   (none)
'
' Returns   : (none)
'
' Notes     :
'   - Called by the arrow Shape via its OnAction property.
'   - Exits silently if the instance id or registry entry cannot be resolved.
' -----------------------------------------------------------------------------------
Public Sub FormDrop_Arrow_OnClick()
    On Error GoTo ErrHandler
    
    Dim callerCtrlName As String
    callerCtrlName = CStr(Application.Caller)

    Dim instanceId As String
    instanceId = clsFormDrop.ParseInstanceId(callerCtrlName)
    If Len(instanceId) = 0 Then Exit Sub

    If g_formDropRegistryDict Is Nothing Then Exit Sub
    If Not g_formDropRegistryDict.Exists(instanceId) Then Exit Sub

    Dim FormDropInst As clsFormDrop
    Set FormDropInst = g_formDropRegistryDict(instanceId)
    FormDropInst.HandleArrowClick
    
CleanExit:
    Exit Sub
    
ErrHandler:
     modErr.ReportError "FormDrop_Arrow_OnClick", Err.Number, Erl
    Resume CleanExit
End Sub

