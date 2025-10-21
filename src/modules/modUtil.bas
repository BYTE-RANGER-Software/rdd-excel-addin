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

Public Function ReplaceWildCards(ByVal InputStr As String) As String
    Dim OutputStr As String
    If InStr(1, InputStr, modConst.WILDCARD_APP_PATH) = 1 Then
        OutputStr = Replace$(InputStr, modConst.WILDCARD_APP_PATH, ThisWorkbook.Path)
    Else
        OutputStr = InputStr
    End If
            
    If InStr(1, OutputStr, modConst.WILDCARD_MY_DOCUMENTS) = 1 Then
        OutputStr = Replace$(OutputStr, modConst.WILDCARD_MY_DOCUMENTS, GetMyDocumentPath)
    End If
    
    ReplaceWildCards = OutputStr
End Function

Public Function AddWildCards(ByVal InputStr As String) As String
    Dim OutputStr As String
    If InStr(1, InputStr, ThisWorkbook.Path) = 1 Then
        OutputStr = Replace$(InputStr, ThisWorkbook.Path, modConst.WILDCARD_APP_PATH)
    Else
        OutputStr = InputStr
    End If
            
    If InStr(1, OutputStr, GetMyDocumentPath) = 1 Then
        OutputStr = Replace$(OutputStr, GetMyDocumentPath, modConst.WILDCARD_MY_DOCUMENTS)
    End If
    
    AddWildCards = OutputStr
End Function

Public Function GetMyDocumentPath() As String
    Dim WshShell As Object
    Set WshShell = CreateObject("WScript.Shell")
    GetMyDocumentPath = WshShell.SpecialFolders("MyDocuments")
    Set WshShell = Nothing
End Function
