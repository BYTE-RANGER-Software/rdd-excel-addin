VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dispatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' ================================================
' Sheet Name: Dispatcher (Macro base)
' Purpose   : Hosts shared macro logic for buttons and automation
' Notes     : This sheet contains critical VBA code used by other sheets.
'             All buttons across the workbook reference macros stored here.
'             DO NOT DELETE OR RENAME THIS SHEET!
'             Doing so will break macro functionality.
'
'             Macros are stored in this sheet module intentionally
'             to avoid requiring "Trust access to the VBA project object model".
'             This approach ensures compatibility with restricted environments.
'
' Author    : [Cal Dymos]
' Date      : [07.10.2025]
' ================================================

Option Explicit
Public Sub InsertPicture_Button()
    ' Opens a file picker, inserts the chosen image into a named range "ImageArea",
    ' scales it to fit while keeping aspect ratio, centers it, and replaces any previous image.

    Const TARGET_RANGE_NAME As String = "Range_ImageArea"
    Const PICTURE_NAME As String = "Scene image"

    Dim fd As FileDialog
    Dim filePath As String
    Dim ws As Worksheet
    Dim rng As Range
    Dim pic As Picture
    Dim maxW As Double, maxH As Double, scaleF As Double
    Dim shp As Shape
    
    Set ws = ActiveSheet

    ' Resolve target range by name, fall back to a hardcoded area if needed
    On Error Resume Next
    Set rng = ws.Range(TARGET_RANGE_NAME)
    On Error GoTo 0
    If rng Is Nothing Then
        ' Fallback, adjust if you do not use the named range
        Set rng = ws.Range("H5:N22")
    End If
    Set rng = rng.MergeArea                      ' use full merge area when cell is merged

    ' File picker for images
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .title = "Select Image"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "image files", "*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.tif;*.tiff"
        If .Show <> -1 Then Exit Sub
        filePath = .SelectedItems(1)
    End With

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' Remove previous picture with the known name
    On Error Resume Next
    ws.Shapes(PICTURE_NAME).Delete
    On Error GoTo 0

    ' Insert picture at default size
    Set pic = ws.Pictures.Insert(filePath)

    ' Name, place, and scale while keeping aspect ratio
    With pic
        .Name = PICTURE_NAME
        .Placement = xlMoveAndSize
        ' First set position near target to avoid jumps during scaling
        .Left = rng.Left
        .Top = rng.Top

        ' Lock aspect ratio via ShapeRange
        .ShapeRange.LockAspectRatio = msoTrue

        ' Compute max width and height available in the target range
        maxW = rng.Width
        maxH = rng.Height

        ' Scale down to fit if larger than the target area
        scaleF = WorksheetFunction.Min(maxW / .Width, maxH / .Height)
        If scaleF < 1 Then
            .Width = .Width * scaleF
            .Height = .Height * scaleF
        End If

        ' Center inside the range
        .Left = rng.Left + (maxW - .Width) / 2
        .Top = rng.Top + (maxH - .Height) / 2
    End With

Cleanup:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub



