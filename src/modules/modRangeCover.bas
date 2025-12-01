Attribute VB_Name = "modRangeCover"
' -----------------------------------------------------------------------------------
' Module    : modRangeCover
' Purpose   : Provides helpers for creating, aligning, showing/hiding, and removing
'             rectangular cover shapes over ranges, and for locking/unlocking ranges.
'
' Notes     :
'   - No UI or orchestration logic here; pure helpers.
' -----------------------------------------------------------------------------------

Option Explicit
Option Private Module

' ===== Public API ===================================================================

' -----------------------------------------------------------------------------------
' Procedure : EnsureRangeCover
' Purpose   : Creates or realigns a rectangular cover shape over a given range.
'
' Parameters:
'   targetSheet   [Worksheet] - Target worksheet.
'   coverRange    [Range]     - Range to cover.
'   coverName     [String]    - Shape name of the cover.
'   bringToFront  [Boolean]   - If True, bring shape to front (default True).
'
' Returns   : (none)
'
' Notes     :
'   - If the shape already exists, it is resized/repositioned to match coverRange.
'   - If not, it is created and initialized via InitCoverShape.
' -----------------------------------------------------------------------------------
Public Sub EnsureRangeCover( _
    ByVal targetSheet As Worksheet, _
    ByVal coverRange As Range, _
    ByVal coverName As String, _
    Optional ByVal bringToFront As Boolean = True)

    If targetSheet Is Nothing Then Exit Sub
    If coverRange Is Nothing Then Exit Sub
    
    Dim coverShape As Shape
    Set coverShape = TryGetShape(targetSheet, coverName)
    
    targetSheet.Unprotect
    If coverShape Is Nothing Then
        Set coverShape = targetSheet.Shapes.AddShape( _
            msoShapeRectangle, _
            coverRange.Left, _
            coverRange.Top, _
            coverRange.Width, _
            coverRange.Height)
        InitCoverShape coverShape, coverRange, coverName
    Else
        ' Realign to current range geometry
        coverShape.Left = coverRange.Left
        coverShape.Top = coverRange.Top
        coverShape.Width = coverRange.Width
        coverShape.Height = coverRange.Height
        coverShape.Placement = xlMoveAndSize
    End If
    
    If bringToFront Then coverShape.ZOrder msoBringToFront
    targetSheet.Protect
End Sub

' -----------------------------------------------------------------------------------
' Procedure : ApplyCoverVisibilityByDropdown
' Purpose   : Hides or shows the cover based on a dropdown cell value.
'             Cover is hidden when a valid selection (not the noneToken) is present.
'
' Parameters:
'   targetSheet     [Worksheet] - Target worksheet.
'   dropdownAddress [String]    - Address of the dropdown cell.
'   noneToken       [String]    - Token in dropdown indicating "no selection"
'                                 (case-insensitive).
'   coverName       [String]    - Shape name of the cover.
'
' Returns   : (none)
'
' Notes     :
'   - If the cover shape does not exist, the procedure exits quietly.
' -----------------------------------------------------------------------------------
Public Sub ApplyCoverVisibilityByDropdown( _
    ByVal targetSheet As Worksheet, _
    ByVal dropdownAddress As String, _
    ByVal noneToken As String, _
    ByVal coverName As String)

    If targetSheet Is Nothing Then Exit Sub
    
    Dim cellValue As String
    Dim showBlock As Boolean
    Dim coverShape As Shape
    
    cellValue = Trim$(CStr(targetSheet.Range(dropdownAddress).value2))
    showBlock = (Len(cellValue) > 0 And LCase$(cellValue) <> LCase$(noneToken))
    
    Set coverShape = TryGetShape(targetSheet, coverName)
    If coverShape Is Nothing Then Exit Sub
    
    ' Range should be visible? cover must be hidden
    coverShape.Visible = Not showBlock
End Sub

' -----------------------------------------------------------------------------------
' Procedure : SetCoverVisible
' Purpose   : Explicitly sets the visibility of the cover shape.
'
' Parameters:
'   targetSheet [Worksheet] - Target worksheet.
'   coverName   [String]    - Shape name of the cover.
'   isVisible   [Boolean]   - Desired visibility state.
'
' Returns   : (none)
'
' Notes     :
'   - Does nothing if the shape does not exist.
' -----------------------------------------------------------------------------------
Public Sub SetCoverVisible( _
    ByVal targetSheet As Worksheet, _
    ByVal coverName As String, _
    ByVal isVisible As Boolean)

    Dim coverShape As Shape
    Set coverShape = TryGetShape(targetSheet, coverName)
    If coverShape Is Nothing Then Exit Sub
    coverShape.Visible = isVisible
End Sub


' -----------------------------------------------------------------------------------
' Procedure : LockRangeAndHideFormulas
' Purpose   : Locks a range and hides formulas, optionally protected with a password.
'
' Parameters:
'   targetSheet [Worksheet] - Target worksheet.
'   targetRange [Range]     - Range to lock and hide formulas.
'   password    [String]    - Optional sheet password (default vbNullString).
'
' Returns   : (none)
'
' Notes     :
'   - Sheet is unprotected, edited, and protected again with UserInterfaceOnly:=True.
'   - Any error during protect/unprotect causes a silent exit via CleanExit label.
' -----------------------------------------------------------------------------------
Public Sub LockRangeAndHideFormulas( _
    ByVal targetSheet As Worksheet, _
    ByVal targetRange As Range, _
    Optional ByVal password As String = vbNullString)

    If targetSheet Is Nothing Then Exit Sub
    If targetRange Is Nothing Then Exit Sub
    
    On Error GoTo CleanExit
    targetSheet.Unprotect password:=password
    targetRange.Locked = True
    targetRange.FormulaHidden = True
    targetSheet.Protect password:=password, UserInterfaceOnly:=True

CleanExit:
End Sub

' -----------------------------------------------------------------------------------
' Procedure : UnlockRangeAndShowFormulas
' Purpose   : Unlocks a range and shows formulas again, optionally protected afterwards.
'
' Parameters:
'   targetSheet [Worksheet] - Target worksheet.
'   targetRange [Range]     - Range to unlock and show formulas.
'   password    [String]    - Optional sheet password (default vbNullString).
'
' Returns   : (none)
'
' Notes     :
'   - Sheet is unprotected, edited, and protected again with UserInterfaceOnly:=True.
'   - Any error during protect/unprotect causes a silent exit via CleanExit label.
' -----------------------------------------------------------------------------------
Public Sub UnlockRangeAndShowFormulas( _
    ByVal targetSheet As Worksheet, _
    ByVal targetRange As Range, _
    Optional ByVal password As String = vbNullString)

    If targetSheet Is Nothing Then Exit Sub
    If targetRange Is Nothing Then Exit Sub
    
    On Error GoTo CleanExit
    targetSheet.Unprotect password:=password
    targetRange.Locked = False
    targetRange.FormulaHidden = False
    targetSheet.Protect password:=password, UserInterfaceOnly:=True

CleanExit:
End Sub

' -----------------------------------------------------------------------------------
' Procedure : RemoveRangeCover
' Purpose   : Deletes the cover shape if present.
'
' Parameters:
'   targetSheet [Worksheet] - Target worksheet.
'   coverName   [String]    - Shape name of the cover to delete.
'
' Returns   : (none)
'
' Notes     :
'   - No error is raised when the shape is missing.
' -----------------------------------------------------------------------------------
Public Sub RemoveRangeCover( _
    ByVal targetSheet As Worksheet, _
    ByVal coverName As String)

    Dim coverShape As Shape
    Set coverShape = TryGetShape(targetSheet, coverName)
    If Not coverShape Is Nothing Then coverShape.Delete
End Sub

' ===== Private Helpers =============================================================

' -----------------------------------------------------------------------------------
' Procedure : InitCoverShape
' Purpose   : Initializes visual and behavioral properties for a newly created cover.
'
' Parameters:
'   coverShape [Shape]  - Shape to initialize.
'   coverRange [Range]  - Range whose top-left color is mirrored.
'   coverName  [String] - Name to assign to the shape.
'
' Returns   : (none)
'
' Notes     :
'   - Fills the shape with the top-left cell color.
'   - Sets no border, movable & resizable with cells, locked and printable.
' -----------------------------------------------------------------------------------
Private Sub InitCoverShape( _
    ByVal coverShape As Shape, _
    ByVal coverRange As Range, _
    ByVal coverName As String)

    Dim colorRgb As Long
    colorRgb = TopLeftDisplayColor(coverRange)
    
    With coverShape
        .name = coverName
        .Fill.Visible = msoTrue
        .Fill.Solid
        .Fill.ForeColor.RGB = colorRgb
        .Fill.Transparency = 0
        .Line.Visible = msoFalse
        .Placement = xlMoveAndSize
        .Locked = True
        .ControlFormat.PrintObject = True
        .Visible = True
    End With
End Sub

' -----------------------------------------------------------------------------------
' Function  : TryGetShape
' Purpose   : Returns a shape by name from a worksheet or Nothing if not found.
'
' Parameters:
'   targetSheet [Worksheet] - Worksheet to search in.
'   coverName   [String]    - Shape name to find.
'
' Returns   : Shape - The found shape; Nothing if not found or on error.
'
' Notes     :
'   - Uses On Error Resume Next intentionally to suppress missing-name errors.
' -----------------------------------------------------------------------------------
Private Function TryGetShape( _
    ByVal targetSheet As Worksheet, _
    ByVal coverName As String) As Shape

    On Error Resume Next
    Set TryGetShape = targetSheet.Shapes(coverName)
    On Error GoTo 0
End Function

' -----------------------------------------------------------------------------------
' Function  : TopLeftDisplayColor
' Purpose   : Gets the display color of the top-left cell, falling back to
'             Interior.Color if DisplayFormat is not available.
'
' Parameters:
'   targetRange [Range] - Range whose top-left cell color is read.
'
' Returns   : Long (RGB color)
'
' Notes     :
'   - Uses DisplayFormat.Interior.Color first (includes conditional formatting).
'   - Falls back to Interior.Color when DisplayFormat is not supported.
' -----------------------------------------------------------------------------------
Private Function TopLeftDisplayColor( _
    ByVal targetRange As Range) As Long

    Dim colorRgb As Long
    On Error Resume Next
    colorRgb = targetRange.Cells(1, 1).DisplayFormat.Interior.Color
    If Err.Number <> 0 Then
        Err.Clear
        colorRgb = targetRange.Cells(1, 1).Interior.Color
    End If
    On Error GoTo 0

    TopLeftDisplayColor = colorRgb
End Function

