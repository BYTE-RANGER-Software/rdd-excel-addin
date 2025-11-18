Attribute VB_Name = "modMain"
' -----------------------------------------------------------------------------------
' Module    : modMain
' Purpose   : Orchestrates use cases; handles UI, validation, selection, confirmation,
'             navigation, logging, error display, and app lifecycle.
'
' Public API:
'   - AppProjectName              : Returns the VBA project name.
'   - AppTempPath                 : Gets or sets the temp/log path.
'   - AppVersion                  : Returns the current add-in version.
'   - HandleWorkbookAddinInstall  : Performs first-time installation initialization.
'   - HandleWorkbookOpen          : Handles add-in startup tasks.
'   - HandleWorkbookBeforeClose   : Handles add-in shutdown tasks.
'   - EnsureWorkbookIsTagged      : Marks a workbook as compatible with this add-in.
'   - IsRDDWorkbook               : Tests whether a workbook is tagged for this add-in.
'   - ShowLog                     : Displays the current log.
'   - ShowManual                  : Opens the manual file.
'   - ShowOptions                 : Displays the options dialog.
'   - ShowAbout                   : Displays the about dialog.
'   - AddNewRoom                  : Creates a new room sheet.
'   - AddNewRoomFromCellCtxMnu    : Creates a new room from the cell context menu.
'   - RemoveCurrentRoom           : Deletes the active room sheet.
'   - GotoRoomFromCell            : Navigates to a room sheet based on the active cell.
'
' Dependencies:
'   - modUtil, modErr, modOptions, modTags, modRooms, modCellCtxMnu, modProps, clsState
'
' Notes     :
'   - Keep this module focused on orchestration, application events, and UI flows.
'   - Business logic should live in feature modules/classes (e.g., modRooms, modOptions).
' -----------------------------------------------------------------------------------

Option Explicit
Option Private Module

' ===== Private State =================================================================
' Module-level private state and WithEvents references used across procedures.

Private m_appTempPath As String
Private m_appProjectName As String

Private m_appEvents As clsAppEvents ' keeps WithEvents sink alive
Private m_activeWorkbookOnInstall As Workbook ' holds ActiveWorkbook on install

' ===== Public API ====================================================================
' Public entry points, properties, and Ribbon callback targets used by the add-in.

' -----------------------------------------------------------------------------------
' Function  : AppProjectName (Get)
' Purpose   : Returns the VBA project name.
' Parameters: (none)
' Returns   : String - Project name
' Notes     : Ensure that SetAppProjectName was executed before the first query.
' -----------------------------------------------------------------------------------
Public Property Get AppProjectName() As String
    AppProjectName = m_appProjectName
End Property

' -----------------------------------------------------------------------------------
' Function  : AppTempPath (Get/Let)
' Purpose   : Gets/Sets the path used for temp/log files.
' Parameters: (none)
' Returns   : String (Get)
' Notes     : Ensure trailing "\" when setting.
' -----------------------------------------------------------------------------------
Public Property Get AppTempPath() As String

    AppTempPath = m_appTempPath

End Property

Public Property Let AppTempPath(ByVal value As String)

    ' Ensure trailing backslash
    If Len(value) > 0 Then
        If Right$(value, 1) <> "\" Then
            value = value & "\"
        End If
    End If
    m_appTempPath = value

End Property

' -----------------------------------------------------------------------------------
' Function  : AppVersion (Get)
' Purpose   : Returns version string from the Add-In, holds in custom document property.
' Parameters: (none)
' Returns   : String - e.g., "1.2.3"
' Notes     : Uses GetDocumentPropertyValue("RDD_AddInVersion").
' -----------------------------------------------------------------------------------
Public Property Get AppVersion() As String
    AppVersion = GetDocumentPropertyValue(ThisWorkbook, "RDD_AddInVersion", "0.0.0")
End Property

' -----------------------------------------------------------------------------------
' Procedure : HandleWorkbookAddinInstall
' Purpose   : Initializes application-specific settings and resources during first-time add-in installation.
'             Setts default properties, creating required named ranges,
'             registering document tags, or preparing the workbook for use with the add-in.
'             Stores ActiveWorkbook reference on installation for later use..
' Parameters: (none)
' Returns   : (none)
' Notes     : Must be called from Workbook_AddinInstall()
' -----------------------------------------------------------------------------------
Public Sub HandleWorkbookAddinInstall()
    'If the add-in is activated when a workbook is opened,
    'save the reference to this workbook for Workbook_Open,
    'since the add-in is set as the active workbook in Workbook_Open.
    If Not ActiveWorkbook Is Nothing Then Set m_activeWorkbookOnInstall = ActiveWorkbook
End Sub

' -----------------------------------------------------------------------------------
' Function  : HandleWorkbookOpen
' Purpose   : Application startup: init logging, wire App events, init state, refresh UI,
'             validating workbook structure.
'
' Parameters: (none)
' Returns   : (none)
' Notes     : must called from Workbook_open(). Requires clsAppEvents and clsState
' -----------------------------------------------------------------------------------
Public Sub HandleWorkbookOpen()
    
    ' Ensure temp path exists before logging
    m_appTempPath = modUtil.GetTempFolder & "\BYTE RANGER"
    If Dir(m_appTempPath, vbDirectory) = "" Then MkDir m_appTempPath
    m_appTempPath = m_appTempPath & "\" & AppProjectName & "\"
    If Dir(m_appTempPath, vbDirectory) = "" Then MkDir m_appTempPath
        
    '
    SetAppProjectName
    
    ' Error Logger
    modErr.InitLogger m_appTempPath, AppProjectName, (AppProjectName & " " & AppVersion)
    
    ' load options
    modOptions.ReadGeneralOptions

    modFormDropManager.FD_InitMngrs
    
    ' wire application events when running as add-in
    If RDDAddInWkBk.IsAddin Then
        ConnectEventHandler
    End If

    ' init State
    clsState.Init
    clsState.InvalidateRibbon
End Sub

' -----------------------------------------------------------------------------------
' Function  : HandleWorkbookBeforeClose
' Purpose   : Application shutdown: saving settings, releasing resources,
'             unhook events, cleanup state, close log.
' Parameters: (none)
' Returns   : (none)
' Notes     : must called from Workbook_BeforeClose(). Safe to call multiple times.
' -----------------------------------------------------------------------------------
Public Sub HandleWorkbookBeforeClose()

    Call SaveGeneralOptions
    
    ' Detach events
    DisconnectEventHandler
    
    modFormDropManager.FD_DisposeMngrs

    '  Update ribbon/UI and clear state
    clsState.InvalidateRibbon
    clsState.Cleanup

    ' close Log
    modErr.CloseLogger
End Sub

' -----------------------------------------------------------------------------------
' Procedure : HandleSheetActivate
' Purpose   : Handles sheet activation for non-add-in workbooks
'
' Parameters:
'   activatedSheet   [Worksheet] - Activated worksheet.
'
' Returns   : (none)
'
' Notes     :
'   - Only acts when SHEET_DISPATCHER is present in the target workbook.
' -----------------------------------------------------------------------------------
Public Sub HandleSheetActivate(ByVal activatedSheet As Worksheet)
    Dim wb As Workbook: Set wb = activatedSheet.Parent
    
    ' Check whether the target workbook has a sheet with the Tag SHEET_LISTS
    If modTags.SheetWithTagExists(wb, SHEET_DISPATCHER) Then
        'Don't update on sheet Lists
        If activatedSheet.CodeName <> SHEET_DISPATCHER And Not modTags.HasSheetTag(activatedSheet, SHEET_DISPATCHER) Then
            If clsState.RoomSheetChanged Then
                clsState.RoomSheetChanged = False
                modRooms.UpdateLists wb
            End If
            If modRooms.IsRoomSheet(activatedSheet) Then
                modRooms.ApplyParallaxRangeCover activatedSheet
            End If
        End If
    End If
    
    clsState.InvalidateRibbon
End Sub

' -----------------------------------------------------------------------------------
' Procedure : HandleSheetChange
' Purpose   : Handles sheet changes for non-add-in worksheets
'
' Parameters:
'   changedSheet  [Worksheet] - Worksheet where the change occurred.
'   targetRng     [Range]     - Changed cell or range.
'
' Returns   : (none)
'
' Notes     :
'   - Only acts on room sheets when the dispatcher tag sheet exists.
' -----------------------------------------------------------------------------------
Public Sub HandleSheetChange(ByVal changedSheet As Worksheet, ByVal targetRng As Range)
    Dim srcBook As Workbook: Set srcBook = changedSheet.Parent

    ' Check whether the TargetRng workbook has a sheet with the Tag SHEET_LISTS
    If modTags.SheetWithTagExists(srcBook, SHEET_DISPATCHER) Then
        ' Only set the change flag if a room sheet has been changed.
        If modTags.HasSheetTag(changedSheet, ROOM_SHEET_ID_TAG_NAME) Then
            clsState.RoomSheetChanged = True
            If Not Intersect(targetRng, changedSheet.Range(NAME_CELL_PARALLAX)) Is Nothing Then
                modRooms.ApplyParallaxRangeCover changedSheet
            End If
        End If
    End If
End Sub

' -----------------------------------------------------------------------------------
' Procedure : HandleSheetBeforeRightClick
' Purpose   : Prepares cell context menu state and invalidates related Ribbon controls
'             before a right-click on a non-add-in worksheet.
'
' Parameters:
'   clickedOnSheet    [Worksheet] - Worksheet where the right-click occurs.
'   targetRng         [Range]     - Target cell or range.
'   shouldCancel      [Boolean]   - (ByRef) Indicates whether the context menu should be
'                                   canceled (passed by reference).
'
' Returns   : (none)
'
' Notes     :
'   - Does not modify shouldCancel in the current implementation.
' -----------------------------------------------------------------------------------
Public Sub HandleSheetBeforeRightClick(ByVal clickedOnSheet As Worksheet, ByVal targetRng As Range, ByRef shouldCancel As Boolean)
    clsState.CellCtxMnuNeedsPrepare = True

    modCellCtxMnu.EvaluateCellCtxMenu clickedOnSheet, targetRng
  
    clsState.InvalidateControl "RB75dd2c44_btnDynCtxMnu1"
    clsState.InvalidateControl "RB75dd2c44_btnDynCtxMnu2"
End Sub

' -----------------------------------------------------------------------------------
' Procedure : HandleWorkbookBeforeSave
' Purpose   : Handles workbook-related tasks before saving, such as persisting
'             workbook-specific options. Only for non-Add-in Workbooks
'
' Parameters:
'   targetBook       [Workbook] - Workbook being saved.
'   showSaveAsUi     [Boolean]  - Indicates whether the Save As UI is shown.
'   shouldCancel     [Boolean]  - (ByRef) Allows canceling the save operation (ByRef).
'
' Returns   : (none)
'
' Notes     :
'   - Currently only persists workbook options, does not alter shouldCancel.
' -----------------------------------------------------------------------------------
Public Sub HandleWorkbookBeforeSave(ByVal targetBook As Workbook, ByVal showSaveAsUi As Boolean, ByRef shouldCancel As Boolean)
    Call modOptions.SaveWorkbookOptions(targetBook)
End Sub
' -----------------------------------------------------------------------------------
' Procedure : EnsureWorkbookIsTagged
' Purpose   : Marks workbook as compatible with this add-in if not already tagged.
' Parameters:
'   wb [Workbook] - Target workbook.
' Returns   : (none)
' Notes     : Tag key/value defined by APP_DOC_TAG_KEY/APP_DOC_TAG_VAL.
' -----------------------------------------------------------------------------------
Public Sub EnsureWorkbookIsTagged(ByVal wb As Workbook)
    If Not modProps.DocumentPropertyExists(wb, APP_DOC_TAG_KEY) Then
        modProps.SetDocumentProperty wb, APP_DOC_TAG_KEY, APP_DOC_TAG_VAL
    End If
End Sub

' -----------------------------------------------------------------------------------
' Function  : IsRDDWorkbook
' Purpose   : Tests whether the workbook is tagged as compatible with this add-in.
'
' Parameters:
'   targetWorkbook   [Workbook] - Workbook to test.
'
' Returns   : Boolean - True when tag matches, otherwise False.
'
' Notes     :
'   - Uses APP_DOC_TAG_KEY/APP_DOC_TAG_VAL.
'   - Returns False defensively when targetWorkbook is Nothing.
' -----------------------------------------------------------------------------------
Public Function IsRDDWorkbook(ByVal targetWorkbook As Workbook) As Boolean
    If targetWorkbook Is Nothing Then Exit Function

    Dim tagValue As String
    tagValue = modProps.GetDocumentPropertyValue(targetWorkbook, APP_DOC_TAG_KEY, "")
    IsRDDWorkbook = (StrComp(tagValue, APP_DOC_TAG_VAL, vbBinaryCompare) = 0)
End Function

' ===== Ribbon Callback Targets ======================================================
' Ribbon callbacks and UI entry points triggered from the custom UI.

' -----------------------------------------------------------------------------------
' Procedure : ShowLog
' Purpose   : Displays the current log to the user.
'
' Parameters: (none)
' Returns   : (none)
'
' Notes     :
'   - Delegates UI/display to modErr.ShowLog.
' -----------------------------------------------------------------------------------
Public Sub ShowLog()
    modErr.ShowLog
End Sub

' -----------------------------------------------------------------------------------
' Procedure : ShowManual
' Purpose   : Opens the manual (PDF/HTML) from the configured path. Shows a message
'             if the file cannot be found and logs any runtime errors.
'
' Parameters: (none)
'
' Returns   : (none)
'
' Notes     :
'   - Uses FILENAME_MANUAL; adjust path provider as needed.
' -----------------------------------------------------------------------------------
Public Sub ShowManual()

    On Error GoTo ErrHandler
    
    Dim manualPath As String
    manualPath = ReplaceWildcards(modOptions.Opt_ManualPath) & "\"  ' TODO: Add options Formular


    If Dir(manualPath & FILENAME_MANUAL) <> "" Then
        ThisWorkbook.FollowHyperlink manualPath & FILENAME_MANUAL
    Else
        ' File not found message
        MsgBox "File " & manualPath & FILENAME_MANUAL & _
            " not found!", vbExclamation, AppProjectName
    End If
    

    On Error GoTo 0
    Exit Sub

ErrHandler:
    modErr.ReportError "ShowManual", Err.Number, Erl, caption:=modMain.AppProjectName
End Sub

' -----------------------------------------------------------------------------------
' Procedure : ShowOptions
' Purpose   : Displays the options form, passing the active workbook context.
' Parameters: (none)
' Returns   : (none)
' Notes     : Form lifetime is scoped to the procedure.
' -----------------------------------------------------------------------------------
Public Sub ShowOptions()

    On Error GoTo ErrHandler
    
    Dim currentOptions As tOptions
    currentOptions = modOptions.GetAllOptions
        
    Dim currentWorkbook As Workbook: Set currentWorkbook = ActiveWorkbook
    
    Dim optionsForm As frmOptions
    Set optionsForm = New frmOptions: optionsForm.Init AppProjectName, currentOptions
        
    optionsForm.Show
        

    If optionsForm.Confirmed Then
        Dim newOptions As tOptions
        newOptions = optionsForm.ResultOptions

        Dim validationError As String
        validationError = modOptions.ValidateOptions(newOptions)
        If LenB(validationError) > 0 Then
            MsgBox validationError, vbExclamation, AppProjectName
            Set optionsForm = Nothing
            Exit Sub
        End If

        modOptions.SetAllOptions newOptions
        modOptions.SaveGeneralOptions
        modOptions.SaveWorkbookOptions currentWorkbook
    End If
    
    Set optionsForm = Nothing
    
    On Error GoTo 0
    Exit Sub

ErrHandler:
    modErr.ReportError "ShowOptions", Err.Number, Erl, caption:=modMain.AppProjectName

End Sub

' -----------------------------------------------------------------------------------
' Procedure : ShowAbout
' Purpose   : Displays the about dialog.
' Parameters: (none)
' Returns   : (none)
' Notes     : Form lifetime is scoped to the procedure.
' -----------------------------------------------------------------------------------
Public Sub ShowAbout()
    On Error GoTo ErrHandler
    
    Dim currentWorkbook As Workbook: Set currentWorkbook = ActiveWorkbook
    Dim aboutForm As frmAbout
    
    Set aboutForm = New frmAbout
    
    aboutForm.Show
    
    Set aboutForm = Nothing
    
    On Error GoTo 0
    Exit Sub
        
ErrHandler:
    modErr.ReportError "ShowAbout", Err.Number, Erl, caption:=modMain.AppProjectName
End Sub
' -----------------------------------------------------------------------------------
' Function  : AddNewRoom
' Purpose   : Runs the "New Room Sheet" dialog, creates the sheet via modRooms,
'             post-processes visuals, and navigates optionally.
' Parameters:
'   shouldGoToNewRoom [Boolean] - If True, jumps to A1 of the created sheet.
' Returns   : String - The created Room ID (empty if cancelled).
' Notes     : Uses frmNewItem and modRooms.
' -----------------------------------------------------------------------------------
Public Function AddNewRoom(Optional ByVal shouldGoToNewRoom As Boolean = True) As String
    On Error GoTo ErrHandler

    Dim currentSheet As Worksheet: Set currentSheet = ActiveSheet
    Dim currentWorkbook As Workbook: Set currentWorkbook = ActiveWorkbook
    Dim currentCell As Range: Set currentCell = ActiveCell

    Dim newSheet As Worksheet
    Dim roomIndex As Long
    Dim roomId As String
    
    Dim newItemForm As frmNewItem: Set newItemForm = New frmNewItem
            
    Application.StatusBar = False
    
    With newItemForm
        .FormCaption = "New Room Sheet"
        .NameLabel = "Room Name"
        .IDLabel = "Room ID"
        .IDVisible = True
        roomIndex = modRooms.GetNextRoomIndex(currentWorkbook)
        roomId = modRooms.GetFormattedRoomID(roomIndex)
        .IDText = roomId
        .nameText = roomId

        .Show                       ' modal
        If Not .Cancelled Then
            
            EnsureWorkbookIsTagged currentWorkbook
     
            Set newSheet = modRooms.AddRoom(currentWorkbook, .nameText, roomIndex)
            If Not newSheet Is Nothing Then
                modUtil.HideOpMode True
                modRooms.ApplyParallaxRangeCover newSheet
                If shouldGoToNewRoom Then
                    Application.GoTo newSheet.Range("A1"), True
                Else
                    currentSheet.Activate
                    If Not currentCell Is Nothing Then currentCell.Select
                End If
                modUtil.HideOpMode False
                AddNewRoom = roomId
            End If
        End If
        Unload newItemForm
    End With
                        
    Set newItemForm = Nothing
                
    On Error GoTo 0
    Exit Function
    
ErrHandler:
    modErr.ReportError "AddNewRoom", Err.Number, Erl, caption:=modMain.AppProjectName
    modUtil.HideOpMode False
End Function

' -----------------------------------------------------------------------------------
' Procedure : AddNewRoomFromCellCtxMnu
' Purpose   : Creates a new room via dialog and writes the new Room ID into ActiveCell.
' Parameters: (none)
' Returns   : (none)
' Notes     : Safe when there is no active cell value.
' -----------------------------------------------------------------------------------
Public Sub AddNewRoomFromCellCtxMnu()
    On Error GoTo ErrHandler
    
    Dim targetCell As Range: Set targetCell = ActiveCell
    
    Dim roomId As String
    
    roomId = AddNewRoom(False)
        
    If Len(roomId) > 0 Then
        If Not targetCell Is Nothing Then targetCell.value = roomId
    End If
    
    
ErrHandler:
    modErr.ReportError "AddNewRoomFromCellCtxMnu", Err.Number, Erl, caption:=modMain.AppProjectName
End Sub

' -----------------------------------------------------------------------------------
' Procedure : RemoveCurrentRoom
' Purpose   : Deletes the active room sheet after confirmation and safety checks.
' Parameters: (none)
' Returns   : (none)
' Notes     : Delegates the deletion to modRooms.RemoveRoom.
' -----------------------------------------------------------------------------------
Public Sub RemoveCurrentRoom()
    On Error GoTo ErrHandler

    Dim roomSheet As Worksheet
    Set roomSheet = ActiveSheet

    Application.StatusBar = False
    
    If Not modRooms.IsRoomSheet(roomSheet) Then
        MsgBox "Active sheet is not a 'Room' sheet.", vbInformation, AppProjectName
        Exit Sub
    End If
    
    ' Confirm with the user
    If MsgBox("Are you sure you want to delete the sheet '" & roomSheet.Name & "'?" & vbCrLf & _
        "This action cannot be undone.", vbYesNo + vbExclamation, "Confirm Sheet Deletion") <> vbYes Then
        Application.StatusBar = "Deletion cancelled."
        Exit Sub
    End If

    Call modRooms.RemoveRoom(roomSheet)

    Exit Sub
ErrHandler:
    modErr.ReportError "RemoveCurrentRoom", Err.Number, Erl, caption:=modMain.AppProjectName
End Sub

' -----------------------------------------------------------------------------------
' Procedure : GotoRoomFromCell
' Purpose   : Jumps to the room sheet referenced by the active cell value.
' Parameters: (none)
' Returns   : (none)
' Notes     : Requires room sheets to be discoverable via modRooms.HasRoomSheet.
' -----------------------------------------------------------------------------------
Public Sub GotoRoomFromCell()
    On Error GoTo ErrHandler
    
    Dim roomId As String
    Dim currentWorkbook As Workbook: Set currentWorkbook = ActiveWorkbook
    Dim currentCell As Range: Set currentCell = ActiveCell
    
    roomId = Trim$(CStr(currentCell.value))
    If Len(roomId) = 0 Then
        MsgBox "No Room ID in the selected cell.", vbInformation, AppProjectName
        Exit Sub
    End If
    
    Dim roomSheet As Worksheet
    If modRooms.HasRoomSheet(currentWorkbook, roomId, roomSheet) Then
        Application.GoTo roomSheet.Range("A1"), True
        Exit Sub
    End If
    
    MsgBox "Room '" & roomId & "' not found.", vbInformation, AppProjectName
    Exit Sub

ErrHandler:
    modErr.ReportError "GotoRoomFromCell", Err.Number, Erl, caption:=modMain.AppProjectName
End Sub

' ===== Private Methods ===============================================================

' -----------------------------------------------------------------------------------
' Function  : ConnectEventHandler
' Purpose   : Enables application-level event handling by assigning the Excel
'             Application object to the clsAppEvents instance.
'
' Params    : (none)
' Returns   : Boolean - True on success; False on failure.
'
' Notes     :
'   - Requires a class module `clsAppEvents` exposing an `App` property (WithEvents).
'   - Logs an error and shows a critical message box on failure.
'   - Keeps a private instance alive in this module.
' -----------------------------------------------------------------------------------
Private Function ConnectEventHandler() As Boolean

    On Error GoTo ErrHandler
    
    If m_appEvents Is Nothing Then Set m_appEvents = New clsAppEvents
    Set m_appEvents.App = Application
    ConnectEventHandler = True
    Exit Function
    
ErrHandler:
    On Error Resume Next
    modErr.ReportError "ConnectEventHandler", Err.Number, Erl, caption:=modMain.AppProjectName, customMessage:="Unable to enable application events."
           
    ' Ensure instance is detached/cleared on failure
    If Not m_appEvents Is Nothing Then Set m_appEvents.App = Nothing
    Set m_appEvents = Nothing
End Function

' -----------------------------------------------------------------------------------
' Function  : DisconnectEventHandler
' Purpose   : Disables application-level event handling by releasing the reference
'             to the Excel Application object.
'
' Params    : (none)
' Returns   : (none)
'
' Notes     :
'   - Safe to call multiple times; sets `AppEvents.App = Nothing`.
'   - Clears instance reference so events stop firing.
' -----------------------------------------------------------------------------------
Private Sub DisconnectEventHandler()
    On Error Resume Next
    If Not m_appEvents Is Nothing Then Set m_appEvents.App = Nothing
    Set m_appEvents = Nothing
    On Error GoTo 0
End Sub

' -----------------------------------------------------------------------------------
' Procedure : SetAppProjectName
' Purpose   : Best-effort retrieval of the project name using Err.Source trick.
' Parameters: (none)
' Returns   : (none)
' Notes     : Consider moving to a fixed constant or document property if available.
' -----------------------------------------------------------------------------------
Private Sub SetAppProjectName()
    On Error Resume Next
    Err.Raise 999
    m_appProjectName = Err.Source
    On Error GoTo 0
End Sub
