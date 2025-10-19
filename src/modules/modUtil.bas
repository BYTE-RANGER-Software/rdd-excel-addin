Attribute VB_Name = "modUtil"
' modUtil
Option Explicit
Option Private Module

' ---------------------------------------------------------------
' Procedure : HideOpMode
' Purpose   : Enables or disables Excel's interactive features
'             to optimize performance and suppress distractions
'             during automated operations.
'
' Parameters:
'   blnEnable [Boolean] - If True, disables events, screen updates,
'                         animations, and alerts (silent mode).
'                         If False, restores normal behavior.
'
' Usage     : Call HideOpMode True before automation,
'             and HideOpMode False afterwards.
' ---------------------------------------------------------------
Public Sub HideOpMode(ByVal blnEnable As Boolean)
    Application.EnableEvents = Not blnEnable
    Application.ScreenUpdating = Not blnEnable
    Application.EnableAnimations = Not blnEnable
    Application.DisplayAlerts = Not blnEnable
End Sub

'--- join a Collection of strings with a separator ---
Public Function JoinCollection(ByVal col As Collection, ByVal sSep As String) As String
    Dim arr() As String, i As Long
    If col.Count = 0 Then Exit Function
    ReDim arr(1 To col.Count)
    For i = 1 To col.Count
        arr(i) = CStr(col(i))
    Next
    JoinCollection = Join(arr, sSep)
End Function

Public Function GetTempFolder() As String
    GetTempFolder = Environ("Temp")
End Function

