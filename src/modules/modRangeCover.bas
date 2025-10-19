Attribute VB_Name = "modRangeCover"
Option Explicit
Option Private Module

' -----------------------------------------------------------------------------------
' Function  : EnsureRangeCover
' Purpose   : Creates or realigns a rectangular cover shape over a given range.
'
'
' Parameters:
'   wks            [Worksheet] - Target worksheet.
'   rngBlock       [Range]     - Range to cover.
'   strCoverName   [String]    - Shape name of the cover.
'   blnBringToFront[Boolean]   - If True, bring shape to front (default True).
'
' Returns   :
'
' Notes     :
' -----------------------------------------------------------------------------------
Public Sub EnsureRangeCover(ByVal wks As Worksheet, ByVal rngBlock As Range, _
                            ByVal strCoverName As String, _
                            Optional ByVal blnBringToFront As Boolean = True)
    If wks Is Nothing Then Exit Sub
    If rngBlock Is Nothing Then Exit Sub
    
    Dim shpCover As Shape
    Set shpCover = TryGetShape(wks, strCoverName)
    
    If shpCover Is Nothing Then
        Set shpCover = wks.Shapes.AddShape(msoShapeRectangle, rngBlock.Left, rngBlock.Top, rngBlock.Width, rngBlock.Height)
        InitCoverShape shpCover, rngBlock, strCoverName
    Else
        ' Realign to current range geometry
        shpCover.Left = rngBlock.Left
        shpCover.Top = rngBlock.Top
        shpCover.Width = rngBlock.Width
        shpCover.Height = rngBlock.Height
        shpCover.Placement = xlMoveAndSize
    End If
    
    If blnBringToFront Then shpCover.ZOrder msoBringToFront
End Sub

' -----------------------------------------------------------------------------------
' Function  : ApplyCoverVisibilityByDropdown
' Purpose   : Hides or shows the cover based on a dropdown cell value.
'             Cover is hidden when a valid selection (not the strNoneToken) is present.
'
' Parameters:
'   wks              [Worksheet] - Target worksheet.
'   strDropdownAddr  [String]    - Address of the dropdown cell.
'   strNoneToken     [String]    - Token in Dropdown that indicating "no selection" (case-insensitive).
'   strCoverName     [String]    - Shape name of the cover.
'
' Returns   :
'
' Notes     :
' -----------------------------------------------------------------------------------
Public Sub ApplyCoverVisibilityByDropdown(ByVal wks As Worksheet, _
                                          ByVal strDropdownlAddr As String, _
                                          ByVal strNoneToken As String, _
                                          ByVal strCoverName As String)
    If wks Is Nothing Then Exit Sub
    
    Dim strValue As String
    Dim bShowBlock As Boolean
    Dim shpCover As Shape
    
    strValue = Trim$(CStr(wks.Range(strDropdownlAddr).Value2))
    bShowBlock = (Len(strValue) > 0 And LCase$(strValue) <> LCase$(strNoneToken))
    
    Set shpCover = TryGetShape(wks, strCoverName)
    If shpCover Is Nothing Then Exit Sub
    
    ' Range should be visible ? cover must be hidden
    shpCover.Visible = Not bShowBlock
End Sub

' -----------------------------------------------------------------------------------
' Function  : SetCoverVisible
' Purpose   : Explicitly sets the visibility of the cover shape.
'
'
' Parameters:
'   wks           [Worksheet] - Target worksheet.
'   strCoverName  [String]    - Shape name of the cover.
'   blnVisible    [Boolean]   - Desired visibility state.
'
' Returns   :
'
' Notes     :
' -----------------------------------------------------------------------------------
Public Sub SetCoverVisible(ByVal wks As Worksheet, ByVal strCoverName As String, ByVal blnVisible As Boolean)
    Dim shpCover As Shape
    Set shpCover = TryGetShape(wks, strCoverName)
    If shpCover Is Nothing Then Exit Sub
    shpCover.Visible = blnVisible
End Sub

' -----------------------------------------------------------------------------------
' Function  : LockRangeAndHideFormulas
' Purpose   : Locks a range and hides formulas, optionally protected with a password.
'
'
' Parameters:
'   wks          [Worksheet] - Target worksheet.
'   rngBlock     [Range]     - Range to lock and hide formulas.
'   strPassword  [String]    - Optional sheet password (default vbNullString).
'
' Returns   :
'
' Notes     :
' -----------------------------------------------------------------------------------
Public Sub LockRangeAndHideFormulas(ByVal wks As Worksheet, ByVal rngBlock As Range, _
                                    Optional ByVal strPassword As String = vbNullString)
    If wks Is Nothing Then Exit Sub
    If rngBlock Is Nothing Then Exit Sub
    
    On Error GoTo CleanExit
    wks.Unprotect Password:=strPassword
    rngBlock.Locked = True
    rngBlock.FormulaHidden = True
    wks.Protect Password:=strPassword, UserInterfaceOnly:=True

CleanExit:
End Sub

' -----------------------------------------------------------------------------------
' Function  : UnlockRangeAndShowFormulas
' Purpose   : Unlocks a range and shows formulas again, optionally protected afterwards.
'
'
' Parameters:
'   wks          [Worksheet] - Target worksheet.
'   rngBlock     [Range]     - Range to unlock and show formulas.
'   strPassword  [String]    - Optional sheet password (default vbNullString).
'
' Returns   :
'
' Notes     :
' -----------------------------------------------------------------------------------
Public Sub UnlockRangeAndShowFormulas(ByVal wks As Worksheet, ByVal rngBlock As Range, _
                                      Optional ByVal strPassword As String = vbNullString)
    If wks Is Nothing Then Exit Sub
    If rngBlock Is Nothing Then Exit Sub
    
    On Error GoTo CleanExit
    wks.Unprotect Password:=strPassword
    rngBlock.Locked = False
    rngBlock.FormulaHidden = False
    wks.Protect Password:=strPassword, UserInterfaceOnly:=True

CleanExit:
End Sub

' -----------------------------------------------------------------------------------
' Function  : RemoveRangeCover
' Purpose   : Deletes the cover shape if present.
'
'
' Parameters:
'   wks           [Worksheet] - Target worksheet.
'   strCoverName  [String]    - Shape name of the cover to delete.
'
' Returns   :
'
' Notes     :
' -----------------------------------------------------------------------------------
Public Sub RemoveRangeCover(ByVal wks As Worksheet, ByVal strCoverName As String)
    Dim shpCover As Shape
    Set shpCover = TryGetShape(wks, strCoverName)
    If Not shpCover Is Nothing Then shpCover.Delete
End Sub

'==================== Helpers ====================

' -----------------------------------------------------------------------------------
' Function  : InitCoverShape
' Purpose   : Initializes visual and behavioral properties for a newly created cover.
'
'
' Parameters:
'   shpCover      [Shape]    - Shape to initialize.
'   rngBlock      [Range]    - Range whose top-left color is mirrored.
'   strCoverName  [String]   - Name to assign to the shape.
'
' Returns   :
'
' Notes     :
' -----------------------------------------------------------------------------------
Private Sub InitCoverShape(ByVal shpCover As Shape, ByVal rngBlock As Range, ByVal strCoverName As String)
    Dim lngRGB As Long
    lngRGB = TopLeftDisplayColor(rngBlock)
    
    With shpCover
        .Name = strCoverName
        .Fill.Visible = msoTrue
        .Fill.Solid
        .Fill.ForeColor.RGB = lngRGB
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
'
' Parameters:
'   wks           [Worksheet] - Worksheet to search in.
'   strCoverName  [String]    - Shape name to find.
'
' Returns   : Shape or Nothing
'
' Notes     :
' -----------------------------------------------------------------------------------
Private Function TryGetShape(ByVal wks As Worksheet, ByVal strCoverName As String) As Shape
    On Error Resume Next
    Set TryGetShape = wks.Shapes(strCoverName)
    On Error GoTo 0
End Function

' -----------------------------------------------------------------------------------
' Function  : TopLeftDisplayColor
' Purpose   : Gets the display color of the top-left cell, fallback to Interior.Color.
'
'
' Parameters:
'   rngBlock   [Range]  - Range whose top-left cell color is read.
'
' Returns   : Long (RGB color)
'
' Notes     :
' -----------------------------------------------------------------------------------
Private Function TopLeftDisplayColor(ByVal rngBlock As Range) As Long
    Dim lngRGB As Long
    On Error Resume Next
    lngRGB = rngBlock.Cells(1, 1).DisplayFormat.Interior.Color
    If Err.Number <> 0 Then
        Err.Clear
        lngRGB = rngBlock.Cells(1, 1).Interior.Color
    End If
    On Error GoTo 0
    TopLeftDisplayColor = lngRGB
End Function

