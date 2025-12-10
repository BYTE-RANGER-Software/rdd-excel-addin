VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFormDrop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' -----------------------------------------------------------------------------------
' Class     : clsFormDrop
' Purpose   : Reusable manager for two Form Controls (DropDown) placed under a target
'             cell. A category dropdown drives a dependent "sub" dropdown. List data
'             lives on a hidden sheet or in named ranges. An optional arrow Shape
'             toggles the dropdown visibility.
'
' Public API:
'   - Init                         : Initializes the instance and registers it.
'   - Destroy                      : Cleans up controls and unregisters the instance.
'   - IsInit                       : Indicates whether the instance is initialized.
'   - id                           : Returns the unique instance id.
'   - SetLists                     : Provides lists (categories + sublists) from code.
'   - SetListsFromNamedRanges      : Provides lists from Range objects.
'   - SetListsFromNames            : Provides lists from workbook Names.
'   - ShowAt                       : Places and prepares dropdowns under a target cell.
'   - HideDropDowns                : Hides dropdowns (and optionally the arrow).
'   - HandleOnAction (Friend)      : Routes category/sub selection changes.
'   - HandleArrowClick (Friend)    : Handles arrow Shape click (toggle behavior).
'   - SetArrowEnabled              : Enables or disables arrow Shape usage.
'   - SetArrowStyle                : Configures arrow Shape geometry and gap.
'   - SetCallbacks                 : Registers callback macros for category/sub choices.
'   - SetPlacement                 : Sets control placement behavior (move / move+size).
'   - ParseInstanceId              : Extracts instance id from control/shape name.
'   - CleanWorkbookArtifacts       : Deletes orphaned controls/shapes named like fd_*.
'   - Logger (Property Set)        : Injects an optional logging implementation.
'
' Dependencies:
'   - modFormDropRouter                  : Calls FormDrop_OnAction and FormDrop_Arrow_OnClick
'                                    as router entry points.
'   - ILog / clsLog                : Optional logger interface/class.
'
' Notes     :
'   - clsFormDrop and modFormDropRouter belong together and depend on each other.
'   - Set Attribute VB_PredeclaredId = True for static-like behavior where required.
' -----------------------------------------------------------------------------------

Option Explicit


' ======= Private state =======
Private m_instId As String
Private m_hostSheet As Worksheet
Private m_hostWorkbook As Workbook
Private m_targetCell As Range

Private m_gapPts As Double
Private m_rowHeight As Double

Private m_catDropDownName As String
Private m_subDropDownName As String

Private m_catItems As Variant                           ' Stores the category names for the drop-down list, e.g., Array("Category1", "Category2", "Category3")
Private m_subItems As Variant                           ' Stores sub-items grouped by category , e.g., Array(Array("apple", "pear", "plum"), Array("carrot", "onion"))
Private m_subListMap As Scripting.Dictionary            ' Scripting.Dictionary: key = category index (1-based) -> address string
Private m_listsInitialized As Boolean

Private m_log As ILog   ' Optional injected logger

Private m_isInitialized As Boolean

' --- List mode (arrays vs ranges) ---
Private Enum FDLM_Mode
    FDLM_None = 0
    FDLM_Arrays = 1
    FDLM_Ranges = 2
End Enum

Private m_listMode As FDLM_Mode

' --- Arrow config/state ---
Private m_arrowEnabled As Boolean
Private m_arrowShapeName As String
Private m_arrowGap As Double
Private m_arrowW As Double
Private m_arrowH As Double

' --- Optional user macros (callbacks) ---
Private m_callbackCatAddress As String  ' "Module.Procedure"
Private m_callbackSubAddress As String  ' "Module.Procedure"

Private Const TRIANGLE_BIG As Long = 9660   ' U+25BC BLACK DOWN-POINTING TRIANGLE
Private Const TRIANGLE_SMALL As Long = 9662 ' U+25BE BLACK DOWN-POINTING SMALL TRIANGLE


' -----------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Purpose   : default initialization of a new class instance.
'
' Parameters:
'   (none)
'
' Returns   : (none)
'
' Notes     :
'   - Actual logical initialization is performed in Init.
'   - keep empty for default instance
' -----------------------------------------------------------------------------------
Private Sub Class_Initialize()

End Sub

' -----------------------------------------------------------------------------------
' Procedure : Class_Terminate
' Purpose   : default cleanup for a class instance when it is destroyed
'
' Parameters:
'   (none)
'
' Returns   : (none)
'
' Notes     :
'   - please use methode dispose() for cleaning
'   - code in this methode is only for failsafe
' -----------------------------------------------------------------------------------
Private Sub Class_Terminate()
    If m_isInitialized Then
        Call Dispose
    End If
End Sub

' ======= Registry (router uses this) ===============================================

' -----------------------------------------------------------------------------------
' Function  : Init
' Purpose   : Initializes the instance, cleans workbook artifacts, sets identifiers,
'             and registers the instance in the provided registry.
'
' Parameters:
'   hostWorkbook              [Workbook]             - Workbook that hosts the controls.
'   formDropRegistry          [Scripting.Dictionary] - Registry mapping instance id to
'                                                      clsFormDrop instances.
'   loggerInst                [clsLog]               - (Optional) clsLog instance
'
' Returns   : Scripting.Dictionary - The updated registry instance.
'
' Notes     :
'   - When hostWorkbook is Nothing, ActiveWorkbook is used.
'   - Safe to call multiple times; initialization runs only once per instance.
' -----------------------------------------------------------------------------------
Public Function Init(ByRef hostWorkbook As Workbook, ByRef formDropRegistry As Scripting.Dictionary, Optional loggerInst As clsLog) As Scripting.Dictionary
    On Error GoTo ErrHandler
    
    If Not m_isInitialized Then
        Randomize
        
        If Not loggerInst Is Nothing Then
            Set m_log = loggerInst
        End If
        
        If hostWorkbook Is Nothing Then Set hostWorkbook = ActiveWorkbook
        Set m_hostWorkbook = hostWorkbook
        
        Dim result As Long
        result = CleanWorkbookArtifacts()
        
        m_instId = GenerateInstanceId
        m_gapPts = 2
        m_rowHeight = 16                                                ' height for Form Control DropDown
        m_catDropDownName = "fd_" & m_instId & "_Cat"
        m_subDropDownName = "fd_" & m_instId & "_Sub"
        Set m_subListMap = CreateObject("Scripting.Dictionary")

        m_listMode = FDLM_None
        m_listsInitialized = False
        
        m_arrowEnabled = True
        m_arrowGap = 2
        m_arrowW = 22
        m_arrowH = 22
        m_arrowShapeName = "fd_" & m_instId & "_Arrow"
    
        'Register instance
        If formDropRegistry Is Nothing Then Set formDropRegistry = New Scripting.Dictionary
        
        If formDropRegistry.Exists(CStr(m_instId)) Then
            Err.Raise ERR_INSTANCE_EXISTS, "clsFormDrop.Init", _
                "Instance ID already registered: " & m_instId
        End If

        Set formDropRegistry(CStr(m_instId)) = Me
        
        m_isInitialized = True
    End If
    
    If Not formDropRegistry Is Nothing Then Set Init = formDropRegistry
    
CleanExit:
    Exit Function
ErrHandler:
    ReportError "Init", Err.Number, Erl
    Resume CleanExit
End Function

' -----------------------------------------------------------------------------------
' Procedure : Dispose
' Purpose   : Handles cleanup for a class instance when it is destroyed; unregisters
'             the instance from the instance registry if it was initialized.
'
' Parameters:
'   (none)
'
' Returns   : (none)
'
' Notes     :
'   - Safe-guards against missing registry or missing key.
' -----------------------------------------------------------------------------------
Public Sub Dispose()
    If Not m_isInitialized Then Exit Sub
        
    ' cleanup controls
    Call CleanupControls
    
    ' reset state
    Set m_hostSheet = Nothing
    Set m_targetCell = Nothing
    Set m_hostWorkbook = Nothing
    
    If Not m_subListMap Is Nothing Then
        m_subListMap.RemoveAll
        Set m_subListMap = Nothing
    End If
    
    Set m_log = Nothing
    
    m_isInitialized = False
End Sub

' -----------------------------------------------------------------------------------
' Procedure : SetCallbacks
' Purpose   : Registers optional callback macros that are executed when a category
'             or sub value is selected.
'
' Parameters:
'   catCallbackAddress   [String] - Optional; full method name for category selection (Module.Procedure).
'   subCallbackAddress   [String] - Optional; full medthod name for sub selection (Module.Procedure).
'
' Returns   : (none)
'
' Notes     :
'   - methods are executed via Application.Run.
'   - methods (macros) must be located in a public module
' -----------------------------------------------------------------------------------
Public Sub SetCallbacks(Optional ByVal catCallbackAddress As String = vbNullString, _
    Optional ByVal subCallbackAddress As String = vbNullString)
                        
    If Not EnsureInitialized Then Exit Sub
    If Not EnsureNotDefaultInstance Then Exit Sub

    m_callbackCatAddress = catCallbackAddress
    m_callbackSubAddress = subCallbackAddress
End Sub

' -----------------------------------------------------------------------------------
' Property  : IsInit (Get)
' Purpose   : Indicates whether the instance has been initialized via Init.
'
' Returns   : Boolean - True if initialized; False otherwise.
' -----------------------------------------------------------------------------------
Public Property Get IsInit() As Boolean

    IsInit = m_isInitialized

End Property

' -----------------------------------------------------------------------------------
' Property  : id (Get)
' Purpose   : Returns the unique instance identifier used in control names and the
'             global registry.
'
' Returns   : String - Instance id.
' -----------------------------------------------------------------------------------
Public Property Get id() As String
    id = m_instId
End Property

' ======= Public API =======

' -----------------------------------------------------------------------------------
' Procedure : SetLists
' Purpose   : Provides category and sublist data from 1-D arrays; writes them to the
'             internal list sheet and builds the sublist mapping.
'
' Parameters:
'   categories   [Variant] - 1-D array of category names (wrapped in Variant).
'   subLists     [Variant] - 1-D array of 1-D arrays; one sublist per category (wrapped in Variant).
'
' Returns   : (none)
'
' Notes     :
'   - Expects both parameters to contain arrays.
' -----------------------------------------------------------------------------------
Public Sub SetLists(ByVal categories As Variant, ByVal subLists As Variant)
    If Not EnsureInitialized Then Exit Sub
    If Not EnsureNotDefaultInstance Then Exit Sub
    
    If Not IsArray(categories) Then Err.Raise ERR_NO_ARRAY, "clsFormDrop.SetList", "categories must be an array of String."
    If Not IsArray(subLists) Then Err.Raise ERR_NO_ARRAY, "clsFormDrop.SetList", "subLists must be an array of arrays."
    
    If (UBound(categories) - LBound(categories)) <> (UBound(subLists) - LBound(subLists)) Then
        Err.Raise ERR_DIF_LENGTH, "clsFormDrop.SetList", "categories and subLists must have the same length."
    End If
    
    ' Store arrays
    m_catItems = categories
    m_subItems = subLists
    
    m_subListMap.RemoveAll

    m_listMode = FDLM_Arrays
    m_listsInitialized = True
End Sub

' -----------------------------------------------------------------------------------
' Procedure : SetListsFromNamedRanges
' Purpose   : Uses Range objects directly as category and sublist sources
'
' Parameters:
'   catRange        [Range]      - Range containing category values (single column or row)
'   subRanges       [Variant]    - Array of Range objects, each representing one sublist (vertical or horizontal)
'
' Returns   : (none)
'
' Notes     :
'   - Builds m_subListMap with addresses pointing to the supplied ranges.
' -----------------------------------------------------------------------------------
Public Sub SetListsFromNamedRanges(ByVal catRange As Range, subRanges As Variant)
    If Not EnsureInitialized Then Exit Sub
    If Not EnsureNotDefaultInstance Then Exit Sub
    
    If (Not IsArray(subRanges)) Then
        Err.Raise ERR_NO_ARRAY, "clsFormDrop.SetListsFromNamedRanges", "subRanges must be an array of Range."
    End If
    
    m_catItems = catRange.value

    m_subListMap.RemoveAll

    Dim i As Long
    Dim r As Range
    For i = LBound(subRanges) To UBound(subRanges)
        Set r = subRanges(i)
        If r Is Nothing Then
            Err.Raise ERR_OBJ_IS_NOTHING, "clsFormDrop.SetListsFromNamedRanges", "subRanges(" & i & ") is Nothing."
        End If
        m_subListMap(CStr(i - LBound(subRanges) + 1)) = "'" & r.Worksheet.name & "'!" & r.Address(False, False)
    Next i
    
    m_listMode = FDLM_Ranges
    m_listsInitialized = True
End Sub

' -----------------------------------------------------------------------------------
' Procedure : SetListsFromNames
' Purpose   : Uses workbook Names to obtain category and sublist ranges, then delegates
'             to SetListsFromNamedRanges.
'
' Parameters:
'   catName        [String]      - Workbook Name for the category range.
'   subNames       [ParamArray]  - One workbook Name per sublist.
'
' Returns   : (none)
'
' Notes     :
'   - Names must refer to valid Range objects.
' -----------------------------------------------------------------------------------
Public Sub SetListsFromNames(ByVal catName As String, ParamArray subNames() As Variant)
    If Not EnsureInitialized Then Exit Sub
    If Not EnsureNotDefaultInstance Then Exit Sub
    
    Dim catRng As Range
    Set catRng = m_hostWorkbook.Names(catName).RefersToRange

    Dim i As Long
    Dim arr() As Range
    ReDim arr(LBound(subNames) To UBound(subNames))
    For i = LBound(subNames) To UBound(subNames)
        Set arr(i) = m_hostWorkbook.Names(CStr(subNames(i))).RefersToRange
    Next i
    Call SetListsFromNamedRanges(catRng, arr)
End Sub

' -----------------------------------------------------------------------------------
' Procedure : SetListsFromTable
' -----------------------------------------------------------------------------------
Public Sub SetListsFromTable(ByVal srcTable As ListObject, ParamArray subNames() As Variant)
    If Not EnsureInitialized Then Exit Sub
    If Not EnsureNotDefaultInstance Then Exit Sub
    
    If Not IsArray(subNames) Then Err.Raise ERR_NO_ARRAY, "clsFormDrop", "subNames must be an array of Strings."

    Dim i As Long
    Dim subRanges() As Range
    Dim catLabels() As String
    ReDim subRanges(LBound(subNames) To UBound(subNames))
    ReDim catLabels(LBound(subNames) To UBound(subNames))
    For i = LBound(subNames) To UBound(subNames)
        catLabels(i) = CStr(subNames(i))
        Set subRanges(i) = srcTable.ListColumns(catLabels(i)).DataBodyRange
    Next i
    Call SetListsFromLabelsAndRanges(catLabels, subRanges)
End Sub

' -----------------------------------------------------------------------------------
' Procedure : SetListsFromLabelsAndRanges
' Purpose   : Use string labels as category list and a Range array as dependent lists.
'
' Parameters:
'   catLabels  [Variant]   - 1-D array of String labels for the categories.
'   subRanges  [Variant]   - 1-D array of Range objects, one per category.
'
' Returns   : (none)
'
' Notes     :
'   - Stores category labels in m_catItems (no internal list sheet).
'   - Builds m_subListMap from the provided ranges (addresses).
'   - Lengths of catLabels and subRanges must match.
' -----------------------------------------------------------------------------------
Public Sub SetListsFromLabelsAndRanges(ByVal catLabels As Variant, ByVal subRanges As Variant)
    If Not EnsureInitialized Then Exit Sub
    If Not EnsureNotDefaultInstance Then Exit Sub

    If Not IsArray(catLabels) Then Err.Raise ERR_NO_ARRAY, "clsFormDrop", "catLabels must be an array of String."
    If Not IsArray(subRanges) Then Err.Raise ERR_NO_ARRAY, "clsFormDrop", "subRanges must be an array of Range."
    If (UBound(catLabels) - LBound(catLabels)) <> (UBound(subRanges) - LBound(subRanges)) Then
        Err.Raise ERR_DIF_LENGTH, "clsFormDrop", "catLabels and subRanges must have the same length."
    End If

    ' Store category labels
    m_catItems = catLabels

    ' Build sub map from ranges
    m_subListMap.RemoveAll
    Dim r As Range
    Dim i As Long
    For i = LBound(subRanges) To UBound(subRanges)
        Set r = subRanges(i)
        If r Is Nothing Then Err.Raise ERR_OBJ_IS_NOTHING, "clsFormDrop", "subRanges(" & i & ") is Nothing."
        m_subListMap(CStr(i - LBound(subRanges) + 1)) = "'" & r.Worksheet.name & "'!" & r.Address(False, False)
    Next i

    ' Range-based list mode
    m_subItems = Empty
    m_listMode = FDLM_Ranges
    m_listsInitialized = True
End Sub

' -----------------------------------------------------------------------------------
' Procedure : ShowAt
' Purpose   : Attaches the dropdown controls to the specified target range on the
'             given worksheet and prepares them for user interaction.
'
' Parameters:
'   hostSheet    [Worksheet] - Worksheet that hosts the controls.
'   targetRange  [Range]     - Target cell range (top-left cell is used).
'
' Returns   : (none)
'
' Notes     :
'   - Requires list data to be initialized beforehand.
'   - Shows only the arrow; dropdowns become visible when the arrow is clicked.
' -----------------------------------------------------------------------------------
Public Sub ShowAt(ByVal hostSheet As Worksheet, ByVal targetRange As Range)
    On Error GoTo ErrHandler:
    
    If Not EnsureInitialized Then Exit Sub
    If Not EnsureNotDefaultInstance Then Exit Sub

    ' Require lists to be initialized
    If Not m_listsInitialized Then
        Err.Raise ERR_LISTS_NOT_SET, "clsFormDrop", _
            "No lists have been initialized. Call SetLists / SetListsFromNamedRanges / SetListsFromNames before ShowAt."
    End If

    Set m_hostSheet = hostSheet
    Set m_targetCell = targetRange.Cells(1, 1)

    ' Ensure controls exist but keep them hidden until arrow is clicked
    Dim ddCat As DropDown: Set ddCat = EnsureDropDown(m_catDropDownName)
    Dim ddSub As DropDown: Set ddSub = EnsureDropDown(m_subDropDownName)

    ' Prepare category dropdown (but do not show yet)
    ddCat.LinkedCell = ""
    ddCat.OnAction = "FormDrop_OnAction"
    ddCat.Left = m_targetCell.Left
    ddCat.Top = m_targetCell.Top + m_targetCell.Height + m_gapPts
    ddCat.Width = m_targetCell.Width
    If IsArray(m_catItems) Then
        If UBound(m_catItems) >= LBound(m_catItems) Then
            ddCat.list = m_catItems
        End If
    End If
    ddCat.Visible = False

    ' Prepare dependent dropdown (hidden)
    ddSub.ListFillRange = ""
    ddSub.LinkedCell = ""
    ddSub.OnAction = ""
    ddSub.Left = ddCat.Left
    ddSub.Top = ddCat.Top + ddCat.Height + 2
    ddSub.Width = ddCat.Width
    ddSub.Visible = False

    ' Only show the arrow next to the target cell
    ShowArrow
    
CleanExit:
    Exit Sub
ErrHandler:
    ReportError "clsFormDrop.ShowAt", Err.Number, Erl
    Resume CleanExit
End Sub

' -----------------------------------------------------------------------------------
' Procedure : HideDropDowns
' Purpose   : Hides both dropdowns while keeping them for reuse; optionally hides the
'             arrow Shape as well.
'
' Parameters:
'   hideArrow   [Boolean] - Optional; when True, also hides the arrow Shape.
'
' Returns   : (none)
'
' Notes     :
'   - Does nothing if the host sheet is not set.
' -----------------------------------------------------------------------------------
Public Sub HideDropDowns(Optional hideArrow As Boolean = False)
    If Not EnsureInitialized Then Exit Sub
    If Not EnsureNotDefaultInstance Then Exit Sub
    
    On Error Resume Next
    m_hostSheet.DropDowns(m_catDropDownName).Visible = False
    m_hostSheet.DropDowns(m_subDropDownName).Visible = False
    On Error GoTo 0
    
    If hideArrow Then HideArrowButton
End Sub

' -----------------------------------------------------------------------------------
' Procedure : HandleOnAction  (Friend)
' Purpose   : Handles category and sub dropdown OnAction events routed from
'             modFormDropRouter; updates dependent lists and calls optional callbacks.
'
' Parameters:
'   callerCtrlName   [String] - Name of the calling dropdown control.
'
' Returns   : (none)
'
' Notes     :
'   - Friend scope: not intended as public API.
' -----------------------------------------------------------------------------------
Friend Sub HandleOnAction(ByVal callerCtrlName As String)
    If StrComp(callerCtrlName, m_catDropDownName, vbTextCompare) = 0 Then
        ' Category changed -> populate sub + optional callback
        Dim ddCat As DropDown: Set ddCat = m_hostSheet.DropDowns(callerCtrlName)
        Dim catIdx As Long: catIdx = ddCat.value
        Dim ddSub As DropDown: Set ddSub = EnsureDropDown(m_subDropDownName)

        ddSub.Visible = False
        ddSub.LinkedCell = ""
        ddSub.OnAction = ""
        
        If catIdx >= 1 Then
            Select Case m_listMode
                Case FDLM_Ranges
                    ' Use address map (Range-based lists)
                    If m_subListMap.Exists(CStr(catIdx)) Then
                        ddSub.ListFillRange = m_subListMap(CStr(catIdx))
                        ddSub.Visible = True
                        ddSub.Left = ddCat.Left
                        ddSub.Top = ddCat.Top + ddCat.Height + 2
                        ddSub.Width = ddCat.Width
                        ddSub.bringToFront
                        ' Enable OnAction only if Callback exists
                        If Len(m_callbackSubAddress) Then
                            ddSub.OnAction = "FormDrop_OnAction"
                        End If
                    End If

                Case FDLM_Arrays
                    ' Use in-memory arrays (SetLists variant)
                    Dim subArr As Variant
                    subArr = m_subItems(LBound(m_subItems) + catIdx - 1)

                    ddSub.ListFillRange = ""
                    ddSub.list = subArr
                    ddSub.Visible = True
                    ddSub.Left = ddCat.Left
                    ddSub.Top = ddCat.Top + ddCat.Height + 2
                    ddSub.Width = ddCat.Width
                    ddSub.bringToFront
                    If Len(m_callbackSubAddress) Then
                        ddSub.OnAction = "FormDrop_OnAction"
                    End If
            End Select
        End If
        
        ' Category-Callback
        If Len(m_callbackCatAddress) Then
            On Error Resume Next
            Application.Run m_callbackCatAddress, ddCat, m_targetCell
            On Error GoTo 0
        End If
        Exit Sub
    End If

    If StrComp(callerCtrlName, m_subDropDownName, vbBinaryCompare) = 0 Then
        ' Sub-Callback
        If Len(m_callbackSubAddress) Then
            Dim dd As DropDown: Set dd = m_hostSheet.DropDowns(callerCtrlName)
            On Error Resume Next
            Application.Run m_callbackSubAddress, dd, m_targetCell
            On Error GoTo 0
            HideDropDowns
        End If
        Exit Sub
    End If

End Sub

' ======= Arrow helpers =============================================================

' -----------------------------------------------------------------------------------
' Function  : EnsureArrowShape
' Purpose   : Ensures that the arrow Shape exists on the host sheet and configures
'             its appearance and behavior.
'
' Parameters:
'   (none)
'
' Returns   : Shape - The arrow Shape instance.
'
' Notes     :
'   - Creates the Shape when it does not exist.
' -----------------------------------------------------------------------------------
Private Function EnsureArrowShape() As Shape
    Dim shp As Shape
    On Error Resume Next
    Set shp = m_hostSheet.Shapes(m_arrowShapeName)
    On Error GoTo 0

    If shp Is Nothing Then
        ' Temporary position; will be placed in ShowArrow
        Set shp = m_hostSheet.Shapes.AddShape(msoShapeRoundedRectangle, _
            m_targetCell.Left, m_targetCell.Top, _
            m_arrowW, m_arrowH)
            
        With shp
            .name = m_arrowShapeName
            .Placement = xlMove
            .OnAction = "FormDrop_Arrow_OnClick"   ' router macro
            .Visible = m_arrowEnabled
            
            .Fill.ForeColor.RGB = RGB(240, 242, 247)
            .Line.ForeColor.RGB = RGB(166, 172, 179)
            .TextFrame2.TextRange.Characters.text = ChrW$(TRIANGLE_BIG)
            .TextFrame2.TextRange.ParagraphFormat.alignment = msoAlignCenter
            .TextFrame2.VerticalAnchor = msoAnchorMiddle
            .TextFrame2.TextRange.Font.Size = 8
            .TextFrame2.TextRange.Font.name = "Segoe UI Symbol"
            ' --- COLOR of the triangle ---
            .TextFrame2.TextRange.Font.Fill.Visible = msoTrue
            .TextFrame2.TextRange.Font.Fill.Solid
            .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(88, 89, 91)
            .Locked = True   ' only works if UserInterfaceOnly:=True
        End With
    End If
    Set EnsureArrowShape = shp
End Function

' -----------------------------------------------------------------------------------
' Procedure : ShowArrow
' Purpose   : Positions and displays the arrow Shape next to the target cell.
'
' Parameters:
'   (none)
'
' Returns   : (none)
'
' Notes     :
'   - Does nothing if the arrow is disabled or the host/target are missing.
' -----------------------------------------------------------------------------------
Private Sub ShowArrow()
    If Not m_arrowEnabled Then Exit Sub
    If m_hostSheet Is Nothing Or m_targetCell Is Nothing Then Exit Sub

    Dim shp As Shape: Set shp = EnsureArrowShape()
    
    With shp
        .Placement = xlMove
        .Height = m_arrowH
        .Width = m_arrowW
        .Left = m_targetCell.Left + m_targetCell.Width + m_arrowGap
        .Top = m_targetCell.Top + (m_targetCell.Height - m_arrowH) - m_arrowGap
        .Visible = True
        .ZOrder msoBringToFront
    End With
End Sub

' -----------------------------------------------------------------------------------
' Procedure : ShowCatOnly
' Purpose   : Shows only the category dropdown below the target cell, hiding the sub
'             dropdown until a category is selected.
'
' Parameters:
'   (none)
'
' Returns   : (none)
'
' Notes     :
'   - Used when the arrow is clicked and dropdowns need to be toggled.
' -----------------------------------------------------------------------------------
Private Sub ShowCatOnly()
    If m_hostSheet Is Nothing Or m_targetCell Is Nothing Then Exit Sub

    Dim ddCat As DropDown: Set ddCat = EnsureDropDown(m_catDropDownName)
    Dim ddSub As DropDown: Set ddSub = EnsureDropDown(m_subDropDownName)

    ' Ensure cat dropdown is properly configured and positioned
    ddCat.ListFillRange = ""
    ddCat.LinkedCell = ""
    ddCat.OnAction = "FormDrop_OnAction"
    ddCat.Left = m_targetCell.Left
    ddCat.Top = m_targetCell.Top + m_targetCell.Height + m_gapPts
    ddCat.Width = m_targetCell.Width
    
    If IsArray(m_catItems) Then
        ddCat.list = m_catItems
    End If
    
    ddCat.Visible = True
    ddCat.bringToFront

    ' Keep sub hidden until a category is chosen
    ddSub.ListFillRange = ""
    ddSub.LinkedCell = ""
    ddSub.OnAction = ""
    ddSub.Left = ddCat.Left
    ddSub.Top = ddCat.Top + ddCat.Height + 2
    ddSub.Width = ddCat.Width
    ddSub.Visible = False
End Sub

' -----------------------------------------------------------------------------------
' Procedure : HideArrowButton
' Purpose   : Hides the arrow Button (Shape) on the host sheet.
'
' Parameters:
'   (none)
'
' Returns   : (none)
'
' Notes     :
'   - Ignores errors if the Shape is missing.
' -----------------------------------------------------------------------------------
Private Sub HideArrowButton()
    On Error Resume Next
    m_hostSheet.Shapes(m_arrowShapeName).Visible = False
    On Error GoTo 0
End Sub

' -----------------------------------------------------------------------------------
' Procedure : HandleArrowClick  (Friend)
' Purpose   : Toggles visibility of the dropdowns when the arrow is clicked.
'
' Parameters:
'   (none)
'
' Returns   : (none)
'
' Notes     :
'   - Recreates the dropdowns when they were deleted.
' -----------------------------------------------------------------------------------
Friend Sub HandleArrowClick()
    ' Toggle visibility of the dropdowns on arrow click
    Dim ddCat As DropDown
    On Error Resume Next
    Set ddCat = m_hostSheet.DropDowns(m_catDropDownName)
    On Error GoTo 0

    If ddCat Is Nothing Then
        ' Controls were deleted; (re)create and show cat dropdown
        ShowCatOnly
        Exit Sub
    End If

    If ddCat.Visible Then
        HideDropDowns
    Else
        ShowCatOnly
    End If
End Sub

' ======= Public options ============================================================

' -----------------------------------------------------------------------------------
' Procedure : SetArrowEnabled
' Purpose   : Enables or disables the arrow Shape; hides the arrow when disabled.
'
' Parameters:
'   enabled   [Boolean] - True to enable the arrow; False to hide and disable it.
'
' Returns   : (none)
'
' Notes     :
'   - When enabling and host/target are known, the arrow is shown immediately.
' -----------------------------------------------------------------------------------
Public Sub SetArrowEnabled(ByVal enabled As Boolean)
    If Not EnsureInitialized Then Exit Sub
    If Not EnsureNotDefaultInstance Then Exit Sub
    
    m_arrowEnabled = enabled
    If Not enabled Then
        HideArrowButton
    Else
        If Not m_hostSheet Is Nothing And Not m_targetCell Is Nothing Then ShowArrow
    End If
End Sub

' -----------------------------------------------------------------------------------
' Procedure : SetArrowStyle
' Purpose   : Configures the arrow Shape gap and dimensions.
'
' Parameters:
'   gapPts   [Double] - Optional; gap in points between cell and arrow.
'   widthPt  [Double] - Optional; arrow width in points.
'   heightPt [Double] - Optional; arrow height in points.
'
' Returns   : (none)
'
' Notes     :
'   - When host/target are known, updates the arrow position immediately.
' -----------------------------------------------------------------------------------
Public Sub SetArrowStyle(Optional ByVal gapPts As Double = 2, _
    Optional ByVal widthPt As Double = 22, _
    Optional ByVal heightPt As Double = 22)
                         
    If Not EnsureInitialized Then Exit Sub
    If Not EnsureNotDefaultInstance Then Exit Sub
    
    m_arrowGap = gapPts
    m_arrowW = widthPt
    m_arrowH = heightPt
    If Not m_hostSheet Is Nothing And Not m_targetCell Is Nothing Then ShowArrow
End Sub

' -----------------------------------------------------------------------------------
' Procedure : SetPlacement
' Purpose   : Sets the placement behavior for dropdown controls (move vs move+size).
'
' Parameters:
'   moveAndSize   [Boolean] - When True, controls move and size with cells.
'
' Returns   : (none)
'
' Notes     :
'   - Does nothing when dropdowns are missing.
' -----------------------------------------------------------------------------------
Public Sub SetPlacement(Optional ByVal moveAndSize As Boolean = False)
    If Not EnsureInitialized Then Exit Sub
    If Not EnsureNotDefaultInstance Then Exit Sub

    On Error Resume Next
    Dim dd As DropDown
    Set dd = m_hostSheet.DropDowns(m_catDropDownName): If Not dd Is Nothing Then dd.Placement = IIf(moveAndSize, xlMoveAndSize, xlMove)
    Set dd = m_hostSheet.DropDowns(m_subDropDownName): If Not dd Is Nothing Then dd.Placement = IIf(moveAndSize, xlMoveAndSize, xlMove)
    On Error GoTo 0
End Sub

' -----------------------------------------------------------------------------------
' Function  : ParseInstanceId
' Purpose   : Parses an instance id out of a control/shape name of the form
'             "fd_<id>_Cat", "fd_<id>_Sub", or "fd_<id>_Arrow".
'
' Parameters:
'   ctrName   [String] - Name of the control/shape.
'
' Returns   : String - The parsed instance id, or an empty string if not found.
'
' Notes     :
'   - Implemented as a public function for use by the router in modFormDropRouter.
' -----------------------------------------------------------------------------------
Public Function ParseInstanceId(ByVal ctrlName As String) As String
    ' ctrlName pattern: "fd_<id>_Cat" or "fd_<id>_Sub" or "fd_<id>_Arrow"
    Dim p1 As Long, p2 As Long
    p1 = InStr(1, ctrlName, "fd_", vbTextCompare)
    If p1 <> 1 Then Exit Function ' must be first part
    p1 = p1 + 3
    p2 = InStr(p1, ctrlName, "_", vbTextCompare)
    If p2 = 0 Then Exit Function
    ParseInstanceId = Mid$(ctrlName, p1, p2 - p1)
End Function

' ======= Helpers ===================================================================

' -----------------------------------------------------------------------------------
' Function  : EnsureDropDown
' Purpose   : Ensures that a DropDown with the given name exists on the host sheet;
'             creates it off-screen when missing.
'
' Parameters:
'   ctrlName   [String] - Name of the DropDown control.
'
' Returns   : DropDown - The DropDown control instance.
'
' Notes     :
'   - Uses cell XFD1048575 as an off-screen position for newly created controls.
' -----------------------------------------------------------------------------------
Private Function EnsureDropDown(ByVal ctrlName As String) As DropDown
    Dim dd As DropDown
    Dim tmpRng As Range
    On Error Resume Next
    Set dd = m_hostSheet.DropDowns(ctrlName)
    On Error GoTo 0

    If dd Is Nothing Then
        ' Temporary position; will be moved in ShowAt
        Set tmpRng = m_hostSheet.Range("XFD1048575")
        Set dd = m_hostSheet.DropDowns.Add(Left:=tmpRng.Left, Top:=tmpRng.Top, Width:=1, Height:=m_rowHeight)
        dd.Visible = False
        
        dd.name = ctrlName
        dd.Placement = xlMove     ' move with cells (no auto-resize)
        dd.PrintObject = False
        dd.Locked = True ' only works if UserInterfaceOnly:=True
    End If

    Set EnsureDropDown = dd
End Function

' -----------------------------------------------------------------------------------
' Function  : CleanWorkbookArtifacts
' Purpose   : Cleans all sheets in the host workbook from orphaned controls and
'             shapes whose names start with the fd_ prefix.
'
' Parameters:
'   (none)
'
' Returns   : Long - Number of deleted objects; negative codes on guard failure.
'
' Notes     :
'   - Returns -2 when EnsureNotDefaultInstance fails.
'   - Returns -1 when m_wbHost is Nothing.
' -----------------------------------------------------------------------------------
Public Function CleanWorkbookArtifacts() As Long
    CleanWorkbookArtifacts = -2 ' Return Code -2
    If Not EnsureNotDefaultInstance Then Exit Function
    CleanWorkbookArtifacts = -1 ' Return Code -1
    Dim ws As Worksheet
    Dim total As Long
    If Not m_hostWorkbook Is Nothing Then
        For Each ws In m_hostWorkbook.Worksheets
            total = total + CleanSheetArtifacts(ws)
        Next ws
        CleanWorkbookArtifacts = total
    End If
End Function

Private Sub CleanupControls()
    
    If m_hostSheet Is Nothing Then Exit Sub
    
    Dim sheet As Worksheet
    On Error Resume Next
    Set sheet = m_hostWorkbook.Worksheets(m_hostSheet.name)
    On Error GoTo 0
    
    If sheet Is Nothing Then
        Set m_hostSheet = Nothing
        Exit Sub
    End If
    
    Dim dd As DropDown
    Set dd = m_hostSheet.DropDowns(m_catDropDownName)
    If Not dd Is Nothing Then dd.Delete
        
    Set dd = m_hostSheet.DropDowns(m_subDropDownName)
    If Not dd Is Nothing Then dd.Delete
        
    Dim shp As Shape
    Set shp = m_hostSheet.Shapes(m_arrowShapeName)
    If Not shp Is Nothing Then shp.Delete

End Sub

' -----------------------------------------------------------------------------------
' Function  : CleanSheetArtifacts
' Purpose   : Cleans a single worksheet from orphaned fd_* controls and arrow shapes.
'
' Parameters:
'   targetSheet [Worksheet] - Worksheet to clean.
'   prefix      [String]    - Optional; name prefix to match (default "fd_").
'
' Returns   : Long - Number of deleted objects on the sheet.
'
' Notes     :
'   - Deletes DropDowns "*_Cat" and "*_Sub" and Shapes "*_Arrow" when prefixed.
' -----------------------------------------------------------------------------------
Private Function CleanSheetArtifacts(ByVal targetSheet As Worksheet, Optional ByVal Prefix As String = "fd_") As Long
    Dim deleted As Long
    Dim i As Long, ctrlNm As String

    On Error Resume Next

    ' 1) Form Controls: DropDowns
    For i = targetSheet.DropDowns.count To 1 Step -1
        ctrlNm = targetSheet.DropDowns(i).name
        If NameHasAllowedPrefix(ctrlNm, Prefix) Then
            If ctrlNm Like "*_Cat" Or ctrlNm Like "*_Sub" Then
                targetSheet.DropDowns(i).Delete
                deleted = deleted + 1
            End If
        End If
    Next i

    ' 2) Shapes: our Arrow button (and any leftover shapes with _Arrow)
    For i = targetSheet.Shapes.count To 1 Step -1
        ctrlNm = targetSheet.Shapes(i).name
        If NameHasAllowedPrefix(ctrlNm, Prefix) Then
            If ctrlNm Like "*_Arrow" Then
                targetSheet.Shapes(i).Delete
                deleted = deleted + 1
            End If
        End If
    Next i

    On Error GoTo 0
    CleanSheetArtifacts = deleted
End Function

' -----------------------------------------------------------------------------------
' Function  : NameHasAllowedPrefix
' Purpose   : Checks whether a given name starts with the specified prefix.
'
' Parameters:
'   nameValue   [String] - Name to test.
'   prefix      [String] - Prefix to compare.
'
' Returns   : Boolean - True when the name starts with the prefix; otherwise False.
'
' Notes     :
'   - Case-insensitive comparison.
' -----------------------------------------------------------------------------------
Private Function NameHasAllowedPrefix(ByVal nameValue As String, ByVal Prefix As String) As Boolean
    If Len(nameValue) >= Len(Prefix) Then
        If StrComp(Left$(nameValue, Len(Prefix)), Prefix, vbTextCompare) = 0 Then
            NameHasAllowedPrefix = True
            Exit Function
        End If
    End If
End Function

' ======= Instance Guard Methods =======================================================
' All instance methods should use these guards. However, "static" methods should not.

' -----------------------------------------------------------------------------------
' Function  : EnsureNotDefaultInstance
' Purpose   : Ensures that a member is not called on the predeclared default instance
'             , but only on real instances created with New.
'
' Parameters:
'   (none)
'
' Returns   : Boolean - True when called on a real instance; raises error otherwise.
'
' Notes     :
'   - Raises ERR_USE_INSTANCE when called on the default instance.
' -----------------------------------------------------------------------------------
Private Function EnsureNotDefaultInstance() As Boolean
#If VBA7 Then
    Dim lpThis As LongPtr, lpDef As LongPtr
#Else
    Dim lpThis As Long, lpDef As Long
#End If
    ' Compare current instance pointer with the predeclared default instance.
    lpThis = ObjPtr(Me)
    lpDef = ObjPtr(clsFormDrop)
    If lpThis = lpDef Then
        Err.Raise ERR_USE_INSTANCE, "clsFormDrop", "Use an instance to call this member."
    End If
    EnsureNotDefaultInstance = True
End Function

' -----------------------------------------------------------------------------------
' Function  : EnsureInitialized
' Purpose   : Ensures that Init has been called before using the instance.
'
' Parameters:
'   (none)
'
' Returns   : Boolean - True when initialized; raises error otherwise.
'
' Notes     :
'   - Raises ERR_NOT_INITIALIZED when the instance is not initialized.
' -----------------------------------------------------------------------------------
Private Function EnsureInitialized() As Boolean
    EnsureInitialized = m_isInitialized
    If Not EnsureInitialized Then
        Err.Raise ERR_NOT_INITIALIZED, "clsFormDrop", "Call .Init first."
    End If
End Function

' ======= Logging ===================================================================

' -----------------------------------------------------------------------------------
' Property  : Logger (Set)
' Purpose   : Injects an optional logger instance implementing ILog/clsLog.
'
' Parameters:
'   value   [clsLog]     - Logger (clsLog) instance.
'
' Returns   : (none)
'
' Notes     :
'   - Logging is optional; guards silently if no logger is provided.
' -----------------------------------------------------------------------------------
Public Property Set Logger(ByVal value As clsLog)
    Set m_log = value
End Property

' -----------------------------------------------------------------------------------
' Procedure : ReportError
' Purpose   : Log an error and optionally show a MsgBox in one call.
'
' Parameters:
'   procedureName [String]              - Procedure where error occurred.
'   errNo         [Long]       (Opt)    - Err.Number; if 0, uses current Err.Number.
'   errLine       [Integer]    (Opt)    - Erl; pass 0 if unknown.
'   showMsg       [Boolean]    (Opt)    - True to show MsgBox; default True.
'   caption       [String]     (Opt)    - MsgBox caption; default procedureName.
'   buttons       [VbMsgBoxStyle] (Opt) - MsgBox style; default Critical + OK.
'   customMessage [String]     (Opt)    - Extra information appended to MsgBox.
'
' Returns   : (none)
'
' Notes     :
' -----------------------------------------------------------------------------------
Private Sub ReportError( _
    ByVal procedureName As String, _
    Optional ByVal errNo As Long = 0, _
    Optional ByVal errLine As Integer = 0, _
    Optional ByVal showMsg As Boolean = True, _
    Optional ByVal caption As String = vbNullString, _
    Optional ByVal buttons As VbMsgBoxStyle = vbOKOnly Or vbCritical, _
    Optional ByVal customMessage As String = vbNullString)

    Dim desc As String, boxCaption As String, body As String
    
    If errNo = 0 And Err.Number <> 0 Then errNo = Err.Number
    If LenB(Err.Description) > 0 Then desc = Err.Description

    On Error Resume Next
    If Not m_log Is Nothing Then m_log.WriteError procedureName, errNo, errLine
    On Error GoTo 0

    If showMsg Then
        If LenB(caption) = 0 Then boxCaption = procedureName Else boxCaption = caption
        body = "Error (" & CStr(errNo) & "): " & desc
        If LenB(customMessage) > 0 Then body = body & vbCrLf & customMessage
        MsgBox body, buttons, boxCaption
    End If
End Sub


Private Function GenerateInstanceId() As String
    Dim timestamp As String
    Dim random1 As Long, random2 As Long, random3 As Long
    Dim instanceId As String
    
    ' micro seconds time stamp (fixed length: 12 characters)
    timestamp = Format$(Now, "yymmddhhnnssms")
    
    ' Randomize Timer + (Rnd() * 10000)
    random1 = Int(Rnd() * 10000)
    random2 = Int(Rnd() * 10000)
    random3 = Int(Rnd() * 10000)
    
    ' Combine to form the full instance ID (24 characters max for the instance ID)
    instanceId = timestamp & _
        Format$(random1, "0000") & _
        Format$(random2, "0000") & _
        Format$(random3, "0000")
    
    ' Ensure the instance ID is 24 characters or less (leave space for "fd_" and "_Cat" or "_Sub")
    If Len(instanceId) > 24 Then
        instanceId = Left(instanceId, 24)
    End If
    
    GenerateInstanceId = instanceId
End Function



