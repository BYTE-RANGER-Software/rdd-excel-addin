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
'   silentMode      (Boolean)           : True enables silent mode, False restores.
'   affectEvents    (Boolean, Optional) : Toggle Application.EnableEvents. Default True.
'   affectScreen    (Boolean, Optional) : Toggle ScreenUpdating. Default True.
'   affectAnimat'ns (Boolean, Optional) : Toggle EnableAnimations. Default True.
'   affectAlerts    (Boolean, Optional) : Toggle DisplayAlerts. Default True.
'
' Notes:
'   - Captures original values on first entry, restores on last exit (depth-aware).
'   - Restores only features that were actually changed in this scope.
'   - Uses Static locals to track depth and touched flags.
' ---------------------------------------------------------------
Public Sub HideOpMode( _
    ByVal silentMode As Boolean, _
    Optional ByVal affectEvents As Boolean = True, _
    Optional ByVal affectScreen As Boolean = True, _
    Optional ByVal affectAnimations As Boolean = True, _
    Optional ByVal affectAlerts As Boolean = True, _
    Optional ByVal affectCalc As Boolean = True)

    Static opDepth                As Long
    Static prevEnableEvents       As Boolean
    Static prevScreenUpdating     As Boolean
    Static prevEnableAnimations   As Boolean
    Static prevDisplayAlerts      As Boolean
    Static prevCalculation     As XlCalculation
    Static touchedEvents          As Boolean
    Static touchedScreen          As Boolean
    Static touchedAnimations      As Boolean
    Static touchedAlerts          As Boolean
    Static touchedCalc            As Boolean

    If silentMode Then
        If opDepth = 0 Then
            prevEnableEvents = Application.EnableEvents
            prevScreenUpdating = Application.ScreenUpdating
            prevEnableAnimations = Application.EnableAnimations
            prevDisplayAlerts = Application.DisplayAlerts
            prevCalculation = Application.Calculation
            touchedEvents = False
            touchedScreen = False
            touchedAnimations = False
            touchedAlerts = False
            touchedCalc = False
        End If

        opDepth = opDepth + 1

        If affectEvents Then
            Application.EnableEvents = False
            touchedEvents = True
        End If
        If affectScreen Then
            Application.ScreenUpdating = False
            touchedScreen = True
        End If
        If affectAnimations Then
            Application.EnableAnimations = False
            touchedAnimations = True
        End If
        If affectAlerts Then
            Application.DisplayAlerts = False
            touchedAlerts = True
        End If
        If affectCalc Then
            Application.Calculation = xlCalculationManual
            touchedCalc = True
        End If
        
    Else
        If opDepth > 0 Then
            opDepth = opDepth - 1
            If opDepth = 0 Then
                If touchedEvents Then Application.EnableEvents = prevEnableEvents
                If touchedScreen Then Application.ScreenUpdating = prevScreenUpdating
                If touchedAnimations Then Application.EnableAnimations = prevEnableAnimations
                If touchedAlerts Then Application.DisplayAlerts = prevDisplayAlerts
                If touchedCalc Then Application.Calculation = prevCalculation
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
    If col.count = 0 Then Exit Function
    ReDim arr(1 To col.count)
    For i = 1 To col.count
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

' -----------------------------------------------------------------------------------
' Function  : SplitTrim
' Purpose   : Splits a string by a separator and trims each entry.
'
' Parameters:
'   sourceText   [String] - Text to split.
'   separator    [String] - Separator string.
'
' Returns   : Variant - One-dimensional array (LB=0) of trimmed strings.
' -----------------------------------------------------------------------------------
Public Function SplitTrim(ByVal sourceText As String, ByVal separator As String) As Variant
    Dim splitArray As Variant: splitArray = Split(sourceText, separator)
    Dim i As Long
    For i = LBound(splitArray) To UBound(splitArray)
        splitArray(i) = Trim$(CStr(splitArray(i)))
    Next
    SplitTrim = splitArray
End Function

