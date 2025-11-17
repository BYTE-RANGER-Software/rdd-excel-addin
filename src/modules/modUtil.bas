Attribute VB_Name = "modUtil"
' ===================================================================================
' Module    : modUtil
' Purpose   : Utility helpers for UI-silent mode, string/path substitutions,
'             collection joining, temp/document paths, and simple string sorting.
'
' Notes     :
'   - Public API section exposes helpers used across modules.
'   - Sorting routines operate in-place on String arrays.
' ===================================================================================
Option Explicit
Option Private Module

' ---------------------------------------------------------------
' Procedure : HideOpMode
' Purpose   : Enable a "silent" mode for automation and restore it safely.
'
' Parameters:
'   blnSilentMode      (Boolean)           : True enables silent mode, False restores.
'   blnAffectEvents    (Boolean, Optional) : Toggle Application.EnableEvents. Default True.
'   blnAffectScreen    (Boolean, Optional) : Toggle ScreenUpdating. Default True.
'   blnAffectAnimat'ns (Boolean, Optional) : Toggle EnableAnimations. Default True.
'   blnAffectAlerts    (Boolean, Optional) : Toggle DisplayAlerts. Default True.
'
' Notes:
'   - Captures original values on first entry, restores on last exit (depth-aware).
'   - Restores only features that were actually changed in this scope.
'   - Uses Static locals to track depth and touched flags.
' ---------------------------------------------------------------
Public Sub HideOpMode( _
    ByVal blnSilentMode As Boolean, _
    Optional ByVal blnAffectEvents As Boolean = True, _
    Optional ByVal blnAffectScreen As Boolean = True, _
    Optional ByVal blnAffectAnimations As Boolean = True, _
    Optional ByVal blnAffectAlerts As Boolean = True)

    Static lngOpDepth                As Long
    Static blnPrevEnableEvents       As Boolean
    Static blnPrevScreenUpdating     As Boolean
    Static blnPrevEnableAnimations   As Boolean
    Static blnPrevDisplayAlerts      As Boolean
    Static blnTouchedEvents          As Boolean
    Static blnTouchedScreen          As Boolean
    Static blnTouchedAnimations      As Boolean
    Static blnTouchedAlerts          As Boolean

    If blnSilentMode Then
        If lngOpDepth = 0 Then
            blnPrevEnableEvents = Application.EnableEvents
            blnPrevScreenUpdating = Application.ScreenUpdating
            blnPrevEnableAnimations = Application.EnableAnimations
            blnPrevDisplayAlerts = Application.DisplayAlerts
            blnTouchedEvents = False
            blnTouchedScreen = False
            blnTouchedAnimations = False
            blnTouchedAlerts = False
        End If

        lngOpDepth = lngOpDepth + 1

        If blnAffectEvents Then
            Application.EnableEvents = False
            blnTouchedEvents = True
        End If
        If blnAffectScreen Then
            Application.ScreenUpdating = False
            blnTouchedScreen = True
        End If
        If blnAffectAnimations Then
            Application.EnableAnimations = False
            blnTouchedAnimations = True
        End If
        If blnAffectAlerts Then
            Application.DisplayAlerts = False
            blnTouchedAlerts = True
        End If

    Else
        If lngOpDepth > 0 Then
            lngOpDepth = lngOpDepth - 1
            If lngOpDepth = 0 Then
                If blnTouchedEvents Then Application.EnableEvents = blnPrevEnableEvents
                If blnTouchedScreen Then Application.ScreenUpdating = blnPrevScreenUpdating
                If blnTouchedAnimations Then Application.EnableAnimations = blnPrevEnableAnimations
                If blnTouchedAlerts Then Application.DisplayAlerts = blnPrevDisplayAlerts
            End If
        End If
    End If
End Sub


' -----------------------------------------------------------------------------------
' Function  : JoinCollection
' Purpose   : Join all items of a string-typed Collection with a separator.
'
' Parameters:
'   col  [Collection] - Source items (string-coercible).
'   sSep [String]     - Separator.
'
' Returns   : String - Joined representation; empty when collection is empty.
' -----------------------------------------------------------------------------------
Public Function JoinCollection(ByVal col As Collection, ByVal sSep As String) As String
    Dim arr() As String, i As Long
    If col.Count = 0 Then Exit Function
    ReDim arr(1 To col.Count)
    For i = 1 To col.Count
        arr(i) = CStr(col(i))
    Next
    JoinCollection = Join(arr, sSep)
End Function

' -----------------------------------------------------------------------------------
' Function  : GetTempFolder
' Purpose   : Return the system temp folder path.
' Parameters: -
' Returns   : String - Temp path.
' -----------------------------------------------------------------------------------
Public Function GetTempFolder() As String
    GetTempFolder = Environ("Temp")
End Function

' -----------------------------------------------------------------------------------
' Function  : ReplaceWildcards
' Purpose   : Replace project/user wildcards with absolute paths.
'
' Parameters:
'   InputStr [String] - Source string possibly containing wildcards.
'
' Returns   : String - Expanded string.
'
' Notes     :
'   - Supports modConst.WILDCARD_APP_PATH and modConst.WILDCARD_MY_DOCUMENTS.
' -----------------------------------------------------------------------------------
Public Function ReplaceWildcards(ByVal InputStr As String) As String
    Dim OutputStr As String
    
    If LenB(InputStr) = 0 Then Exit Function
    
    If InStr(1, InputStr, modConst.WILDCARD_APP_PATH) = 1 Then
        OutputStr = Replace$(InputStr, modConst.WILDCARD_APP_PATH, ThisWorkbook.path)
    Else
        OutputStr = InputStr
    End If
            
    If InStr(1, OutputStr, modConst.WILDCARD_MY_DOCUMENTS) = 1 Then
        OutputStr = Replace$(OutputStr, modConst.WILDCARD_MY_DOCUMENTS, GetMyDocumentsPath)
    End If
    
    ReplaceWildcards = OutputStr
End Function

' -----------------------------------------------------------------------------------
' Function  : AddWildcards
' Purpose   : Collapse absolute paths back to configured wildcards.
'
' Parameters:
'   InputStr [String] - String containing absolute paths.
'
' Returns   : String - String with wildcards re-applied.
' -----------------------------------------------------------------------------------
Public Function AddWildcards(ByVal InputStr As String) As String
    Dim OutputStr As String
    
    If LenB(InputStr) = 0 Then Exit Function
    
    If InStr(1, InputStr, ThisWorkbook.path) = 1 Then
        OutputStr = Replace$(InputStr, ThisWorkbook.path, modConst.WILDCARD_APP_PATH)
    Else
        OutputStr = InputStr
    End If
            
    If InStr(1, OutputStr, GetMyDocumentsPath) = 1 Then
        OutputStr = Replace$(OutputStr, GetMyDocumentsPath, modConst.WILDCARD_MY_DOCUMENTS)
    End If
    
    AddWildcards = OutputStr
End Function

' -----------------------------------------------------------------------------------
' Function  : GetMyDocumentsPath
' Purpose   : Return the current user's "My Documents" folder path.
' Parameters: -
' Returns   : String - Path to "My Documents".
' -----------------------------------------------------------------------------------
Public Function GetMyDocumentsPath() As String
    Dim WshShell As Object
    Set WshShell = CreateObject("WScript.Shell")
    GetMyDocumentsPath = WshShell.SpecialFolders("MyDocuments")
    Set WshShell = Nothing
End Function

' -----------------------------------------------------------------------------------
' Procedure : BubbleSortStringArray
' Purpose   : Sort a String array in ascending order using Bubble Sort (in-place).
'
' Parameters:
'   arr() [String] - Target array (must be dimensioned).
'
' Returns   : (none)
' -----------------------------------------------------------------------------------
Public Sub BubbleSortStringArray(ByRef arr() As String)
    Dim i As Long, j As Long
    Dim strTemp As String
    Dim upper As Long
    
    If (Not arr) Or UBound(arr) < LBound(arr) Then Exit Sub
    
    upper = UBound(arr)

    For i = LBound(arr) To upper - 1
        For j = LBound(arr) To upper - i - 1
            If arr(j) > arr(j + 1) Then
                strTemp = arr(j)
                arr(j) = arr(j + 1)
                arr(j + 1) = strTemp
            End If
        Next j
    Next i
End Sub

' -----------------------------------------------------------------------------------
' Procedure : QuickSortStringArray
' Purpose   : Sort a String array in ascending order using QuickSort (in-place).
'
' Parameters:
'   arr()  [String] - Target array (must be dimensioned).
'   first  [Long]   - Start index.
'   last   [Long]   - End index.
'
' Returns   : (none)
' -----------------------------------------------------------------------------------
Public Sub QuickSortStringArray(ByRef arr() As String, ByVal first As Long, ByVal last As Long)
    Dim low As Long, high As Long
    Dim strPivot As String, strTemp As String

    low = first
    high = last
    
    If (Not arr) Or UBound(arr) < LBound(arr) Then Exit Sub
    
    If first >= last Then Exit Sub
    
    strPivot = arr((first + last) \ 2)

    Do While low <= high
        Do While arr(low) < strPivot
            low = low + 1
        Loop
        Do While arr(high) > strPivot
            high = high - 1
        Loop
        If low <= high Then
            strTemp = arr(low)
            arr(low) = arr(high)
            arr(high) = strTemp
            low = low + 1
            high = high - 1
        End If
    Loop

    If first < high Then QuickSortStringArray arr, first, high
    If low < last Then QuickSortStringArray arr, low, last
End Sub


