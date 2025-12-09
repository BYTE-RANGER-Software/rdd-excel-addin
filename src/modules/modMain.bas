Attribute VB_Name = "modMain"
' -----------------------------------------------------------------------------------
' Module    : modMain
' Purpose   : Central application controller and business logic hub.
'             Manages app lifecycle, handles application-level events (delegated from
'             clsAppEvents), orchestrates feature workflows, provides Ribbon callbacks,
'             and implements business logic for user interactions.
'
' Public API:
'   === Configuration ===
'   - AppProjectName              : Returns the VBA project name
'   - AppTempPath                 : Gets/sets temp/log path
'   - AppVersion                  : Returns add-in version
'
'   === Lifecycle (Add-In) ===
'   - HandleAddInWorkbookInstall       : First-time installation initialization
'   - HandleAddInWorkbookOpen          : Startup tasks (logging, events, state)
'   - HandleAddInWorkbookBeforeClose   : Shutdown tasks (cleanup, save settings)
'
'   === Event Business Logic (RDD Workbooks) ===
'   - HandleSheetActivate         : Sheet activation logic
'   - HandleSheetDeactivate       :
'   - HandleSheetChange           : Sheet change logic
'   - HandleSheetBeforeRightClick : Right-click menu preparation
'   - HandleWorkbookBeforeSave    : Pre-save operations
'   - HandleWorkbookOpen          :
'
'   === FormDrop Callbacks ===
'   - OnFormDropCatSelected       : Category dropdown selection logic
'   - OnFormDropSubSelected       : Sub dropdown selection logic
'
'   === Workbook Management ===
'   - EnsureWorkbookIsTagged      : Marks workbook as add-in compatible
'   - IsRDDWorkbook               : Tests workbook compatibility tag
'
'   === UI Entry Points (Ribbon) ===
'   - ShowLog                     : Displays log viewer
'   - ShowManual                  : Opens manual file
'   - ShowOptions                 : Displays options dialog
'   - ShowAbout                   : Displays about dialog
'
'   === Feature Orchestration ===
'   - AddNewRoom                  : Creates new room sheet with dialog
'   - AddNewRoomFromCellCtxMnu    : Creates room and writes ID to cell
'   - RemoveCurrentRoom           : Deletes active room sheet
'   - GotoRoomFromCell            : Navigates to room referenced in cell
'   - EditRoomIdentity            : Edits room ID and alias with dialog
'
' Dependencies:
'   - clsAppEvents    : Event sink (delegates to this module)
'   - clsState        : Application state management
'   - modUtil         : Utility functions
'   - modErr          : Error handling and logging
'   - modOptions      : Settings management
'   - modTags         : Sheet/workbook tagging
'   - modRooms        : Room sheet operations
'   - modCellCtxMnu   : Cell context menu
'   - modProps        : Document properties
'
' Notes:
'   - This module acts as the **central controller** for the application
'   - Event handlers contain business logic (not just delegation)
'   - Keeps orchestration logic separate from low-level feature implementation
'   - Maintains single responsibility: "What should happen when..." not "How to do..."
' -----------------------------------------------------------------------------------

Option Explicit
Option Private Module


'Enum for the different cell/range change categories
Private Enum ChangeCategory
    CC_None = 0
    CC_Parallax = 1
    CC_RoomMetadata = 2        ' Room ID, No, Alias
    CC_SceneMetadata = 3       ' Scene ID
    CC_GeneralSettings = 4     ' Game Heigth, Width, BG Heigth, Width, UI Heigth
    CC_Actors = 5
    CC_Sounds = 6
    CC_SpecialFX = 7
    CC_TouchableObjects = 8 ' Hotspot ID + Name
    CC_PickupableObjects = 9 ' Item ID + Name
    CC_MultiStateObjects = 10 ' State Object ID + Name
    CC_Flags = 11
End Enum

' ===== Private State =================================================================
' Module-level private state and WithEvents references used across procedures.

Private m_appTempPath As String
Private m_appProjectName As String

Private m_appEvents As clsAppEvents ' keeps WithEvents sink alive
Private m_formDropMgr As clsFormDropManager
Private m_activeWorkbookOnInstall As Workbook ' holds ActiveWorkbook on install

' ===== Public API ====================================================================
' Public entry points, properties, and Ribbon callback targets used by the add-in.

' -----------------------------------------------------------------------------------
' Property  : AppProjectName (Get)
' Purpose   : Returns the VBA project name.
' Parameters: (none)
' Returns   : String - Project name
' Notes     : Ensure that SetAppProjectName was executed before the first query.
' -----------------------------------------------------------------------------------
Public Property Get AppProjectName() As String
    AppProjectName = m_appProjectName
End Property

' -----------------------------------------------------------------------------------
' Property  : AppTempPath (Get/Let)
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
' Property  : AppVersion (Get)
' Purpose   : Returns version string from the Add-In, holds in custom document property.
' Parameters: (none)
' Returns   : String - e.g., "1.2.3"
' Notes     : Uses GetDocumentPropertyValue("RDD_AddInVersion").
' -----------------------------------------------------------------------------------
Public Property Get AppVersion() As String
    AppVersion = GetDocumentPropertyValue(ThisWorkbook, "RDD_AddInVersion", "0.0.0")
End Property

' -----------------------------------------------------------------------------------
' Procedure : HandleAddInWorkbookInstall
' Purpose   : Initializes application-specific settings and resources during first-time add-in installation.
'             Sets default properties, creating required named ranges,
'             registering document tags, or preparing the workbook for use with the add-in.
'             Stores ActiveWorkbook reference on installation for later use..
' Parameters: (none)
' Returns   : (none)
' Notes     : Must be called from Workbook_AddinInstall()
' -----------------------------------------------------------------------------------
Public Sub HandleAddInWorkbookInstall()
    'If the add-in is activated when a workbook is opened,
    'save the reference to this workbook for Workbook_Open,
    'since the add-in is set as the active workbook in Workbook_Open.
    If Not ActiveWorkbook Is Nothing Then Set m_activeWorkbookOnInstall = ActiveWorkbook
End Sub

' -----------------------------------------------------------------------------------
' Function  : HandleAddInWorkbookOpen
' Purpose   : Application startup: init logging, wire App events, init state, refresh UI,
'             validating workbook structure.
'
' Parameters: (none)
' Returns   : (none)
' Notes     : must called from Workbook_open(). Requires clsState
' -----------------------------------------------------------------------------------
Public Sub HandleAddInWorkbookOpen()
    On Error GoTo ErrHandler
    
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
    
    If m_formDropMgr Is Nothing Then Set m_formDropMgr = New clsFormDropManager
    m_formDropMgr.Init _
        onCatCallback:="modFormDropCallbacks.OnFormDropCatSelected", _
        onSubCallback:="modFormDropCallbacks.OnFormDropSubSelected"
        
    'init frmWait
    frmWait.Init
    
    ' wire application events when running as add-in
    If RDDAddInWkBk.IsAddin Then
        ConnectEventHandler
    End If

    ' init State
    clsState.Init
    clsState.InvalidateRibbon
    
    Exit Sub

ErrHandler:
    
    modErr.ReportError "HandleAddInWorkbookOpen", Err.Number, Erl, _
        caption:=modMain.AppProjectName
End Sub

' -----------------------------------------------------------------------------------
' Function  : HandleAddInWorkbookBeforeClose
' Purpose   : Application shutdown: saving settings, releasing resources,
'             unhook events, cleanup state, close log.
' Parameters: (none)
' Returns   : (none)
' Notes     : must called from Workbook_BeforeClose(). Safe to call multiple times.
' -----------------------------------------------------------------------------------
Public Sub HandleAddInWorkbookBeforeClose(Cancel As Boolean)

    Call SaveGeneralOptions
    
    ' Detach events
    DisconnectEventHandler
    
    Set m_formDropMgr = Nothing

    '  Update ribbon/UI and clear state
    clsState.InvalidateRibbon
    clsState.Cleanup

    ' close Log
    modErr.CloseLogger
End Sub

' -----------------------------------------------------------------------------------
' Function  : HandleWorkbookOpen
' Purpose   : handles opening of non RDD-AddIn workbooks
'
' Parameters: (none)
' Returns   : (none)
' Notes     : must called from clsAppEvent App_WorkbookOpen(). Requires clsAppEvents
' -----------------------------------------------------------------------------------
Public Sub HandleWorkbookOpen(ByVal targetBook As Workbook)
    On Error GoTo ErrHandler
    ' only on RDD Workbooks
    If Not IsRDDWorkbook(targetBook) Then Exit Sub
    modOptions.ReadWorkbookOptions targetBook
    
    Exit Sub
ErrHandler:
    modErr.ReportError "HandleWorkbookOpen", Err.Number, Erl, caption:=modMain.AppProjectName
End Sub

' -----------------------------------------------------------------------------------
' Procedure : HandleSheetDeactivate
' Purpose   : Handles business logic when a worksheet is deactivated.
'             Updates application state, refreshes UI elements, and manages
'             context-sensitive features based on the deactivated sheet.
' Parameters: sh  [Worksheet] - Worksheet that was deactivated
' Returns   : (none)
' Notes     : Called by clsAppEvents.App_SheetDeactivate event handler.
' -----------------------------------------------------------------------------------
Public Sub HandleSheetDeactivate(ByVal sh As Worksheet)
    On Error GoTo ErrHandler
    
    ' Code that must run on all Sheets
    ' ...
    
    
    'only on RoomSheets
    If Not IsRoomSheet(sh) Then Exit Sub
    ' Code That runs only on Room sheets
    m_formDropMgr.HandleSheetDeactivate sh
    
    Exit Sub
ErrHandler:
    modErr.ReportError "HandleSheetDeactivate", Err.Number, Erl, caption:=modMain.AppProjectName
End Sub

' -----------------------------------------------------------------------------------
' Procedure : HandleSheetSelectionChange
' Purpose   : Handles business logic when worksheet cells are changed.
'             Implements data validation, cascading updates, or other change-triggered
'             workflows based on changed ranges.
' Parameters: sh      [Worksheet]   - Worksheet where changes occurred
'             Target  [Range]    - Range that was changed
' Returns   : (none)
' Notes     : Called by clsAppEvents.App_SheetChange event handler.
' -----------------------------------------------------------------------------------
Public Sub HandleSheetSelectionChange(ByVal sh As Worksheet, ByVal Target As Range)
    On Error GoTo ErrHandler
    'only on RoomSheets
    If Not IsRoomSheet(sh) Then Exit Sub
    
    m_formDropMgr.HandleSelectionChange sh, Target
    
    Exit Sub
ErrHandler:
    modErr.ReportError "HandleSheetSelectionChange", Err.Number, Erl, caption:=modMain.AppProjectName
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
    On Error GoTo ErrHandler
    
    ' Code that must run on all Sheets
    clsState.InvalidateRibbon
        
    'only on RDD Workbooks and on RoomSheets
    If Not IsRDDWorkbook(activatedSheet.Parent) Then Exit Sub
    If Not modRooms.IsRoomSheet(activatedSheet) Then Exit Sub
                        
    modRooms.ApplyParallaxRangeCover activatedSheet
    
    Exit Sub
    
ErrHandler:
    modErr.ReportError "HandleSheetActivate", Err.Number, Erl, caption:=modMain.AppProjectName
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
    On Error GoTo ErrHandler
    
    Dim srcBook As Workbook: Set srcBook = changedSheet.Parent
    
    ' Only process on RDD Workbooks and on room sheets
    If Not IsRDDWorkbook(srcBook) Then Exit Sub
    If Not modRooms.IsRoomSheet(changedSheet) Then Exit Sub
            
    ' Set the general change flag
    clsState.RoomSheetChanged = True
    clsState.RoomsValidated = False
    
    clsState.InvalidateControl RIBBON_BTN_BUILD_DATA
    clsState.InvalidateControl RIBBON_BTN_SYNC_LISTS
    clsState.InvalidateControl RIBBON_BTN_NEED_SYNC_LISTS
    
    ' Determine what type of change occurred
    Dim changeType As ChangeCategory
    changeType = DetermineChangeCategory(changedSheet, targetRng)

    ' Handle the change with optimal SYNC/APPEND strategy
    Select Case changeType
        Case CC_Parallax
            modRooms.ApplyParallaxRangeCover changedSheet
            
        Case CC_RoomMetadata
            modRooms.UpdateRoomsMetadataLists srcBook, changedSheet, LUM_Sync
            
        Case CC_SceneMetadata
            modRooms.UpdateScenesMetadataLists srcBook, changedSheet, LUM_Sync
            
        Case CC_GeneralSettings
            modRooms.UpdateGeneralSettingsLists srcBook, changedSheet, LUM_Sync
            
        Case CC_Actors
            modRooms.UpdateActorsLists srcBook, changedSheet, LUM_Append
            
        Case CC_Sounds
            modRooms.UpdateSoundsLists srcBook, changedSheet, LUM_Append
            
        Case CC_SpecialFX
            modRooms.UpdateSpecialFXLists srcBook, changedSheet, LUM_Append
            
        Case CC_Flags
            modRooms.UpdateFlagsLists srcBook, changedSheet, LUM_Append
            
        Case CC_PickupableObjects
            modRooms.UpdateItemsLists srcBook, changedSheet, LUM_Append
            
        Case CC_MultiStateObjects
            modRooms.UpdateStateObjectsLists srcBook, changedSheet, LUM_Append
            
        Case CC_TouchableObjects
            modRooms.UpdateHotspotsLists srcBook, changedSheet, LUM_Append
            
        Case CC_None
            ' No list update needed for this change
            
    End Select
    
    Exit Sub
    
ErrHandler:
    modErr.ReportError "HandleSheetActivate", Err.Number, Erl, caption:=modMain.AppProjectName
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
    On Error GoTo ErrHandler
    
    If modRooms.IsRoomSheet(clickedOnSheet) Then
        clsState.CellCtxMnuNeedsPrepare = True

        modCellCtxMnu.EvaluateCellCtxMenu clickedOnSheet, targetRng
  
        clsState.InvalidateControl "RB75dd2c44_btnDynCtxMnu1"
        clsState.InvalidateControl "RB75dd2c44_btnDynCtxMnu2"
    Else
        modCellCtxMnu.ResetToDefaultCtxMenu
    End If
    
CleanExit:
    Exit Sub
ErrHandler:
    modCellCtxMnu.ResetToDefaultCtxMenu  ' Fallback auf Standard
    modErr.ReportError "HandleSheetBeforeRightClick", Err.Number, Erl, _
        caption:=modMain.AppProjectName
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
    On Error GoTo ErrHandler
    
    If Not IsRDDWorkbook(targetBook) Then Exit Sub
    modOptions.SaveWorkbookOptions targetBook
    
    Exit Sub
    
ErrHandler:
    modErr.ReportError "HandleWorkbookBeforeSave", Err.Number, Erl, caption:=modMain.AppProjectName
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
        modProps.SetDocumentPropertyValue wb, APP_DOC_TAG_KEY, APP_DOC_TAG_VAL
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
    manualPath = ReplaceWildcards(modOptions.Opt_ManualPath) & "\"


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
    Dim optionsForm As frmOptions
    Dim currentWorkbook As Workbook
    Dim hasRDDWorkbook As Boolean
    Dim dataDict As Scripting.Dictionary
    Dim lo As ListObject
    
    ' Check whether an RDD workbook is active
    Set currentWorkbook = ActiveWorkbook
    hasRDDWorkbook = IsRDDWorkbook(currentWorkbook)
    
    '  Load current options
    currentOptions = modOptions.GetAllOptions()
    
    ' If RDD workbook is active, also load workbook options
    If hasRDDWorkbook Then
        modOptions.ReadWorkbookOptions currentWorkbook
        currentOptions = modOptions.GetAllOptions()  ' Retrieve again after ReadWorkbookOptions
        
        ' Load data lists for combo boxes
        Set dataDict = New Scripting.Dictionary
    
        ' get ListObject
        Set lo = currentWorkbook.Worksheets(modConst.SHEET_DISPATCHER).ListObjects(modConst.NAME_DATA_TABLE)
        If Not lo Is Nothing Then
   
            ' Extract columns in arrays
            dataDict(modConst.LISTS_HEADER_PERSPECTIVE) = modLists.CollectTableColumnValuesToArray(lo, modConst.LISTS_HEADER_PERSPECTIVE, True)
            dataDict(modConst.LISTS_HEADER_PARALLAX) = modLists.CollectTableColumnValuesToArray(lo, modConst.LISTS_HEADER_PARALLAX, True)
            dataDict(modConst.LISTS_HEADER_SCENE_MODE) = modLists.CollectTableColumnValuesToArray(lo, modConst.LISTS_HEADER_SCENE_MODE, True)

        End If
    End If
    
    'Initialize and display dialog
    Set optionsForm = New frmOptions
    With optionsForm
        .Init AppProjectName, currentOptions, hasRDDWorkbook, dataDict
        .Show vbModal
        
        If .Confirmed Then
            ' Apply changes
            modOptions.SetAllOptions .ResultOptions
            
            ' Speichere General Options (Registry)
            modOptions.SaveGeneralOptions
            
            ' Speichere Workbook Options (Document Properties)
            If hasRDDWorkbook Then
                modOptions.SaveWorkbookOptions currentWorkbook
            End If
            
        End If
    End With
    
CleanExit:
    Unload optionsForm: Set optionsForm = Nothing
    Set dataDict = Nothing
    Exit Sub
    
ErrHandler:
    modErr.ReportError "ShowOptions", Err.Number, Erl, caption:=AppProjectName
    Resume CleanExit
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
' Notes     : Uses frmObjectEdit and modRooms.
' -----------------------------------------------------------------------------------
Public Function AddNewRoom(Optional ByVal shouldGoToNewRoom As Boolean = True) As String
    On Error GoTo ErrHandler

    Dim currentSheet As Worksheet: Set currentSheet = ActiveSheet
    Dim currentWorkbook As Workbook: Set currentWorkbook = ActiveWorkbook
    Dim currentCell As Range: Set currentCell = ActiveCell

    Dim newSheet As Worksheet
    Dim roomIndex As Long
    Dim roomID As String
    Dim roomNo As Long
    Dim roomName As String
    Dim sceneID As String
    
    Dim newItemForm As frmObjectEdit: Set newItemForm = New frmObjectEdit
            
    Application.StatusBar = False
    
    With newItemForm
        roomIndex = modRooms.GetNextRoomIndex(currentWorkbook)
        roomID = modRooms.GetFormattedRoomID(roomIndex)
        
        .FormCaption = "New Room Sheet"
        .Field4Visible = True
        .Field5Visible = False
        
        ' Field 1: Scene ID/Name
        .Label1Text = "Scene ID/Name"
        .Text1Locked = False
        .Text1NumericOnly = False
        .Text1Value = vbNullString
        .Text1Tip = "Scene identifier or name for this room (optional), e.g., Temple"
        .Text1RequiresValue = False
        
        ' Field 2: Room Name
        .Label2Text = "Room Name"
        .Text2Locked = False
        .Text2NumericOnly = False
        .Text2Value = modConst.ROOM_SHEET_DEFAULT_PREFIX & " " & CStr(roomIndex)
        .Text2Tip = "Descriptive alias for the room, e.g., Temple Entrance"
        .Text2RequiresValue = True
                
        ' Field 3: Room No
        .Label3Text = "Room No"
        .Text3Locked = False
        .Text3NumericOnly = True
        .Text3Value = CStr(roomIndex)
        .Text3Tip = "AGS room number e.g., 1 or 20"
        .Text3RequiresValue = True
                
        ' Field 4: Room ID (locked)
        .Label4Text = "Room ID"
        .Text4Locked = True
        .Text4NumericOnly = False
        .Text4Value = roomID
        .Text4Tip = "This is the short unique ID for the room"
        .Text4RequiresValue = False
                        
        Do
            .Show vbModal
            If .Cancelled Then
                Unload newItemForm: Set newItemForm = Nothing
                Exit Function
            End If
        
            sceneID = Trim$(.Text1Value)
            roomName = Trim$(.Text2Value)
            roomNo = CLng(Trim$(.Text3Value))
            
            If modRooms.IsValidAGSRoomNo(roomNo) Then
                If modRooms.HasRoomNo(currentWorkbook, roomNo) Then
                    MsgBox "Room No '" & CStr(roomNo) & "' already exists !" & vbCrLf & _
                        "Please choose a different room no.", _
                        vbExclamation, AppProjectName
                Else 'New room No is unique
                    Exit Do
                End If
            Else
                MsgBox "Room No '" & CStr(roomNo) & "' is not a valid AGS room no !" & vbCrLf & _
                    "Please choose a different room no.", _
                    vbExclamation, AppProjectName
            End If
        Loop
        
    End With
    
    Unload newItemForm: Set newItemForm = Nothing
    
    frmWait.ShowDialog
    modUtil.HideOpMode True
            
    Set newSheet = modRooms.AddRoom(currentWorkbook, roomName, roomIndex, roomNo, sceneID)
    If Not newSheet Is Nothing Then
    
        EnsureWorkbookIsTagged currentWorkbook
        
        modRooms.ApplyParallaxRangeCover newSheet
        If shouldGoToNewRoom Then
            Application.GoTo newSheet.Range("A1"), True
        Else
            currentSheet.Activate
            If Not currentCell Is Nothing Then currentCell.Select
        End If
        modUtil.HideOpMode False
        AddNewRoom = roomID
    End If
    
    clsState.InvalidateRibbon
                
    frmWait.Hide
    On Error GoTo 0
    Exit Function
    
ErrHandler:
    modUtil.HideOpMode False
    frmWait.Hide
    modErr.ReportError "AddNewRoom", Err.Number, Erl, caption:=modMain.AppProjectName
End Function

' -----------------------------------------------------------------------------------
' Procedure : AddNewRoomFromCellCtxMnu
' Purpose   : Creates a new room sheet and writes the Room ID into the currently
'             selected cell. Triggered from cell context menu.
' Parameters: (none)
' Returns   : (none)
' Notes     : Context menu callback. Similar to AddNewRoom but auto-populates cell.
'             Requires active cell selection.
' -----------------------------------------------------------------------------------
Public Sub AddNewRoomFromCellCtxMnu()
    On Error GoTo ErrHandler
    
    Dim targetCell As Range: Set targetCell = ActiveCell
    
    Dim roomID As String
    
    roomID = AddNewRoom(False)
        
    If Len(roomID) > 0 Then
        If Not targetCell Is Nothing Then targetCell.value = roomID
    End If
    
    Exit Sub
ErrHandler:
    modErr.ReportError "AddNewRoomFromCellCtxMnu", Err.Number, Erl, caption:=modMain.AppProjectName
End Sub

' -----------------------------------------------------------------------------------
' Procedure : RemoveCurrentRoom
' Purpose   : Deletes the currently active room sheet after user confirmation.
'             Validates sheet is a room sheet before deletion.
' Parameters: (none)
' Returns   : (none)
' Notes     : Requires room sheet validation via modRooms.IsRoomSheet.
'             Shows confirmation dialog before deletion.
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
    If MsgBox("Are you sure you want to delete the sheet '" & roomSheet.name & "'?" & vbCrLf & _
        "This action cannot be undone.", vbYesNo + vbExclamation, "Confirm Sheet Deletion") <> vbYes Then
        Application.StatusBar = "Deletion cancelled."
        Exit Sub
    End If

    frmWait.ShowDialog
    modUtil.HideOpMode True
     
    modRooms.RemoveRoom roomSheet
    
    modUtil.HideOpMode False
    frmWait.Hide

    Exit Sub
ErrHandler:
    frmWait.Hide
    modUtil.HideOpMode False
    modErr.ReportError "RemoveCurrentRoom", Err.Number, Erl, caption:=modMain.AppProjectName
End Sub

' -----------------------------------------------------------------------------------
' Procedure : BuildPDCDataFromRooms
' Purpose   : Orchestrates PDC data generation from all room sheets.
'             Validates preconditions, shows progress dialog, calls modPDC.BuildPdcData,
'             updates application state, and provides user feedback.
'
' Parameters: (none)
' Returns   : (none)
'
' Notes     :
'   - Requires validation to pass first (RoomsValidated = True, RoomsValidationIssueCount = 0)
'   - Creates/updates "PDCData" sheet with nodes and edges
'   - Updates clsState.PDCDataBuilt on success
' -----------------------------------------------------------------------------------
Public Sub BuildPDCDataFromRooms()
    On Error GoTo ErrHandler
    
    ' Verify we have a valid workbook
    If Workbooks.count = 0 Then Exit Sub
    
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    
    ' Verify this is an RDD workbook
    If Not modMain.IsRDDWorkbook(wb) Then
        MsgBox "This is not an RDD workbook.", vbExclamation, modMain.AppProjectName
        Exit Sub
    End If
    
    ' Verify validation has passed
    If Not clsState.RoomsValidated Or clsState.RoomsValidationIssueCount > 0 Then
        MsgBox "Please validate room data first." & vbCrLf & _
               "Use 'Validate' button before building PDC data.", _
               vbExclamation, modMain.AppProjectName
        Exit Sub
    End If
    
    ' Check if there are any room sheets
    Dim hasRoomSheets As Boolean
    Dim ws As Worksheet
    
    hasRoomSheets = False
    For Each ws In wb.Worksheets
        If modRooms.IsRoomSheet(ws) Then
            hasRoomSheets = True
            Exit For
        End If
    Next ws
    
    If Not hasRoomSheets Then
        MsgBox "No room sheets found to process.", vbInformation, modMain.AppProjectName
        Exit Sub
    End If
    
    ' Show progress and build data
    frmWait.ShowDialog
    modUtil.HideOpMode True
    
    Dim nodesCount As Long
    Dim edgesCount As Long
    
    modPDC.BuildPdcData wb, nodesCount, edgesCount
    
    ' Update state
    clsState.PDCDataBuilt = True
    
    ' Invalidate buttons that depend on PDC data
    clsState.InvalidateControl RIBBON_BTN_BUILD_CHART
    clsState.InvalidateControl RIBBON_BTN_UPDATE_CHART
    
    modUtil.HideOpMode False
    frmWait.Hide
    
    ' User feedback
    MsgBox "PDC Data created:" & vbCrLf & _
           "- " & nodesCount & " Nodes" & vbCrLf & _
           "- " & edgesCount & " Edges", _
           vbInformation, modMain.AppProjectName
    
    Exit Sub
    
ErrHandler:
    frmWait.Hide
    modUtil.HideOpMode False
    modErr.ReportError "modMain.BuildPDCDataFromRooms", Err.Number, Erl, caption:=modMain.AppProjectName
End Sub


' -----------------------------------------------------------------------------------
' Procedure : GeneratePuzzleDependencyChart
' Purpose   : Orchestrates PDC chart generation from PDCData.
'             Validates preconditions, shows progress dialog, calls modPDC.GeneratePuzzleChart,
'             and provides user feedback.
'
' Parameters: (none)
' Returns   : (none)
'
' Notes     :
'   - Requires PDCData sheet to exist (PDCDataBuilt = True)
'   - Creates/replaces "Chart" sheet with visual node/edge representation
' -----------------------------------------------------------------------------------
Public Sub GeneratePuzzleDependencyChart()
    On Error GoTo ErrHandler
    
    ' Verify we have a valid workbook
    If Workbooks.count = 0 Then Exit Sub
    
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    
    ' Verify this is an RDD workbook
    If Not modMain.IsRDDWorkbook(wb) Then
        MsgBox "This is not an RDD workbook.", vbExclamation, modMain.AppProjectName
        Exit Sub
    End If
    
    ' Verify PDCData exists
    If Not clsState.PDCDataBuilt Then
        MsgBox "Please build PDC data first." & vbCrLf & _
               "Use 'Build Data' button before generating the chart.", _
               vbExclamation, modMain.AppProjectName
        Exit Sub
    End If
    
    ' Check if PDCData sheet exists
    On Error Resume Next
    Dim dataSheet As Worksheet
    Set dataSheet = wb.Sheets("PDCData")
    On Error GoTo ErrHandler
    
    If dataSheet Is Nothing Then
        MsgBox "PDCData sheet not found." & vbCrLf & _
               "Please rebuild PDC data first.", _
               vbExclamation, modMain.AppProjectName
        Exit Sub
    End If
    
  ' Check if Chart sheet already exists - ask for confirmation to overwrite
    On Error Resume Next
    Dim existingChart As Worksheet
    Set existingChart = wb.Sheets("Chart")
    On Error GoTo ErrHandler
    
    If Not existingChart Is Nothing Then
        Dim result As VbMsgBoxResult
        result = MsgBox("A chart already exists." & vbCrLf & _
                        "Do you want to replace it?", _
                        vbQuestion + vbYesNo, modMain.AppProjectName)
        If result = vbNo Then Exit Sub
    End If

    ' Show progress and generate chart
    frmWait.ShowDialog
    modUtil.HideOpMode True
    
    modPDC.GeneratePuzzleChart wb
    
    modUtil.HideOpMode False
    frmWait.Hide
    
    ' User feedback
    MsgBox "Puzzle Dependency Chart generated!", vbInformation, modMain.AppProjectName
    
    Exit Sub
    
ErrHandler:
    frmWait.Hide
    modUtil.HideOpMode False
    modErr.ReportError "modMain.GeneratePuzzleDependencyChart", Err.Number, Erl, caption:=modMain.AppProjectName
End Sub


' -----------------------------------------------------------------------------------
' Procedure : SynchonizePuzzleDependencyChart
' Purpose   : Orchestrates PDC chart synchronization with PDCData.
'             Validates preconditions, shows progress dialog, calls modPDC.SyncPuzzleChart,
'             and provides user feedback.
'
' Parameters: (none)
' Returns   : (none)
'
' Notes     :
'   - Requires both PDCData and Chart sheets to exist
'   - Updates existing nodes, creates new ones, redraws all edges
' -----------------------------------------------------------------------------------
Public Sub SynchonizePuzzleDependencyChart()
    On Error GoTo ErrHandler
    
    ' Verify we have a valid workbook
    If Workbooks.count = 0 Then Exit Sub
    
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    
    ' Verify this is an RDD workbook
    If Not modMain.IsRDDWorkbook(wb) Then
        MsgBox "This is not an RDD workbook.", vbExclamation, modMain.AppProjectName
        Exit Sub
    End If
    
    ' Check if PDCData sheet exists
    On Error Resume Next
    Dim dataSheet As Worksheet
    Set dataSheet = wb.Sheets("PDCData")
    On Error GoTo ErrHandler
    
    If dataSheet Is Nothing Then
        MsgBox "PDCData sheet not found." & vbCrLf & _
               "Please build PDC data first.", _
               vbExclamation, modMain.AppProjectName
        Exit Sub
    End If
    
    ' Check if Chart sheet exists
    On Error Resume Next
    Dim chartSheet As Worksheet
    Set chartSheet = wb.Sheets("Chart")
    On Error GoTo ErrHandler
    
    If chartSheet Is Nothing Then
        MsgBox "Chart sheet not found." & vbCrLf & _
               "Please generate the chart first.", _
               vbExclamation, modMain.AppProjectName
        Exit Sub
    End If
    
    ' Show progress and sync chart
    frmWait.ShowDialog
    modUtil.HideOpMode True
    
    modPDC.SyncPuzzleChart wb
    
    modUtil.HideOpMode False
    frmWait.Hide
    
    ' User feedback
    MsgBox "Chart synchronized!", vbInformation, modMain.AppProjectName
    
    Exit Sub
    
ErrHandler:
    frmWait.Hide
    modUtil.HideOpMode False
    modErr.ReportError "modMain.SynchonizePuzzleDependencyChart", Err.Number, Erl, caption:=modMain.AppProjectName
End Sub


' -----------------------------------------------------------------------------------
' Procedure : GotoRoomFromCell
' Purpose   : Navigates to the room sheet referenced by the Room ID in the active cell.
'             Displays message if room not found or cell is empty.
' Parameters: (none)
' Returns   : (none)
' Notes     : Requires room sheets to be discoverable via modRooms.HasRoomID.
'             Context menu callback.
' -----------------------------------------------------------------------------------
Public Sub GotoRoomFromCell()
    On Error GoTo ErrHandler
    
    Dim roomID As String
    Dim currentWorkbook As Workbook: Set currentWorkbook = ActiveWorkbook
    Dim currentCell As Range: Set currentCell = ActiveCell
    
    roomID = Trim$(CStr(currentCell.value))
    If Len(roomID) = 0 Then
        MsgBox "No Room ID in the selected cell.", vbInformation, AppProjectName
        Exit Sub
    End If
    
    Dim roomSheet As Worksheet
    If modRooms.HasRoomID(currentWorkbook, roomID, roomSheet) Then
        Application.GoTo roomSheet.Range("A1"), True
        Exit Sub
    End If
    
    MsgBox "Room '" & roomID & "' not found.", vbInformation, AppProjectName
    Exit Sub

ErrHandler:
    modErr.ReportError "GotoRoomFromCell", Err.Number, Erl, caption:=modMain.AppProjectName
End Sub

' -----------------------------------------------------------------------------------
' Procedure : EditRoomIdentity
' Purpose   : Opens dialog to edit Room ID and Room Alias of the active room sheet,
'             then updates all references throughout the workbook.
' Parameters: (none)
' Returns   : (none)
' Notes     : Uses frmObjectEdit for input. Delegates reference updates to
'             modRooms.UpdateRoomReferences. Shows progress via HideOpMode.
' -----------------------------------------------------------------------------------
Public Sub EditRoomIdentity()
    On Error GoTo ErrHandler
    
    Dim targetSheet As Worksheet
    Dim targetBook As Workbook
    Dim currentRoomID As String
    Dim currentRoomAlias As String
    Dim currentRoomNo As Long
    Dim newRoomID As String
    Dim newRoomAlias As String
    Dim newRoomNo As Long
    Dim oldRoomIDCell As Range
    Dim oldRoomAliasCell As Range
    Dim oldRoomNoCell As Range
    
    Set targetSheet = ActiveSheet
    Set targetBook = targetSheet.Parent
    
    If Not modRooms.IsRoomSheet(targetSheet, currentRoomID) Then
        MsgBox "The active sheet is not a Room sheet.", vbExclamation, AppProjectName
        Exit Sub
    End If
    
    On Error Resume Next
    Set oldRoomIDCell = targetSheet.Range(NAME_CELL_ROOM_ID)
    Set oldRoomAliasCell = targetSheet.Range(NAME_CELL_ROOM_ALIAS)
    Set oldRoomNoCell = targetSheet.Range(NAME_CELL_ROOM_NO)
    On Error GoTo ErrHandler
    
    If oldRoomIDCell Is Nothing Then
        MsgBox "Named range '" & NAME_CELL_ROOM_ID & "' not found on this sheet.", _
            vbExclamation, AppProjectName
        Exit Sub
    End If
    
    If oldRoomAliasCell Is Nothing Then
        MsgBox "Named range '" & NAME_CELL_ROOM_ALIAS & "' not found on this sheet.", _
            vbExclamation, AppProjectName
        Exit Sub
    End If
    
    If oldRoomNoCell Is Nothing Then
        MsgBox "Named range '" & NAME_CELL_ROOM_NO & "' not found on this sheet.", _
            vbExclamation, AppProjectName
        Exit Sub
    End If
    
    currentRoomID = oldRoomIDCell.value
    currentRoomAlias = oldRoomAliasCell.value
    currentRoomNo = CLng(oldRoomNoCell.value)
    
    Dim frmEdit As frmObjectEdit: Set frmEdit = New frmObjectEdit
    
    With frmEdit
        .FormCaption = "Edit Room Identity"
        .Field6Visible = True
        
        .Label1Text = "Current Room ID:"
        .Text1Locked = True
        .Text1Value = currentRoomID
        .Text1RequiresValue = False
        
        .Label2Text = "New Room ID:"
        .Text2Prefix = modConst.ROOM_SHEET_ID_PREFIX
        .Text2Locked = False
        .Text2Value = currentRoomID
        .Text2RequiresValue = True
        
        .Label3Text = "Current Room No:"
        .Text3Locked = True
        .Text3Value = currentRoomNo
        .Text3RequiresValue = False
        
        .Label4Text = "New Room No:"
        .Text4Locked = False
        .Text4Value = currentRoomNo
        .Text4RequiresValue = True
        .Text4NumericOnly = True
        
        .Label5Text = "Current Room Alias:"
        .Text5Locked = True
        .Text5Value = currentRoomAlias
        .Text5RequiresValue = False
        
        .Label6Text = "New Room Alias:"
        .Text6Prefix = modConst.ROOM_SHEET_ALIAS_PREFIX
        .Text6Locked = False
        .Text6Value = currentRoomAlias
        .Text6RequiresValue = True
    
        Do
            .Show vbModal
    
            If .Cancelled Then
                Unload frmEdit: Set frmEdit = Nothing
                Exit Sub
            End If
            
            newRoomID = Trim$(.Text2Value)
            newRoomNo = CLng(.Text4Value)
            newRoomAlias = Trim$(.Text6Value)
            
            'No changes were made.
            If newRoomID = currentRoomID And newRoomAlias = currentRoomAlias And newRoomNo = currentRoomNo Then
                Unload frmEdit: Set frmEdit = Nothing
                Exit Sub
            End If
            
            If newRoomID <> currentRoomID Then
                If modRooms.HasRoomID(targetBook, newRoomID) Then
                    MsgBox "Room ID '" & newRoomID & "' already exists !" & vbCrLf & _
                        "Please choose a different Room ID.", _
                        vbExclamation, AppProjectName
                Else 'New room ID is unique
                    Exit Do
                End If
            End If
            If newRoomAlias <> currentRoomAlias Then
                If modRooms.HasRoomAlias(targetBook, newRoomAlias) Then
                    MsgBox "Room Alias '" & newRoomAlias & "' already exists !" & vbCrLf & _
                        "Please choose a different Room Alias.", _
                        vbExclamation, AppProjectName
                Else 'New room Alias is unique
                    Exit Do
                End If
            End If
            
            If newRoomNo <> currentRoomNo Then
                If Not modRooms.IsValidAGSRoomNo(newRoomNo) Then
                    MsgBox "Room No '" & CStr(newRoomNo) & "' is not a valid AGS room no !" & vbCrLf & _
                        "Please choose a different room no.", _
                        vbExclamation, AppProjectName
                Else
                
                    If modRooms.HasRoomNo(targetBook, newRoomNo) Then
                        MsgBox "Room No '" & CStr(newRoomNo) & "' already exists !" & vbCrLf & _
                            "Please choose a different room no.", _
                            vbExclamation, AppProjectName
                    Else
                        Exit Do
                    End If
                    
                End If
            End If
            
        Loop
           
    End With
    
    Unload frmEdit: Set frmEdit = Nothing
        
    frmWait.ShowDialog
    
    modUtil.HideOpMode True
    
    oldRoomIDCell.value = newRoomID
    oldRoomAliasCell.value = newRoomAlias
    oldRoomNoCell.value = newRoomNo
    
    If newRoomID <> currentRoomID Then
        modTags.TagSheet targetSheet, ROOM_SHEET_ID_TAG_NAME, newRoomID
    End If
    
    Call modRooms.UpdateRoomReferences(targetSheet.Parent, currentRoomID, currentRoomAlias, _
        newRoomID, newRoomAlias)
    
    modUtil.HideOpMode False
    
    frmWait.Hide

    MsgBox "Room Identity updated successfully." & vbCrLf & vbCrLf & _
        "Old Room ID: " & currentRoomID & vbCrLf & _
        "New Room ID: " & newRoomID & vbCrLf & vbCrLf & _
        "Old Room No: " & currentRoomNo & vbCrLf & _
        "New Room No: " & newRoomNo & vbCrLf & vbCrLf & _
        "Old Room Alias: " & currentRoomAlias & vbCrLf & _
        "New Room Alias: " & newRoomAlias, _
        vbInformation, AppProjectName
    
CleanExit:
    Exit Sub
    
ErrHandler:
    modUtil.HideOpMode False
    frmWait.Hide
    modErr.ReportError "modMain.EditRoomIdentity", Err.Number, Erl, caption:=AppProjectName
    Resume CleanExit
End Sub

Public Sub SyncAllLists()
    On Error GoTo ErrHandler

    Dim wb As Workbook
    Set wb = ActiveWorkbook
     
    frmWait.ShowDialog
    modUtil.HideOpMode True
     
    ' Full SYNC of all categories
    modRooms.SynchronizeAllLists wb
          
    clsState.RoomSheetChanged = False
    clsState.InvalidateControl RIBBON_BTN_SYNC_LISTS
    clsState.InvalidateControl RIBBON_BTN_NEED_SYNC_LISTS
     
    modUtil.HideOpMode False
    frmWait.Hide
     
    MsgBox "All lists synchronized successfully!", vbInformation, modMain.AppProjectName
     
CleanExit:
    Exit Sub

ErrHandler:
    modUtil.HideOpMode False
    frmWait.Hide
    MsgBox "Error synchronizing lists: " & Err.Description, vbCritical, modMain.AppProjectName
    Resume CleanExit
End Sub
 
Public Sub ValidateRoomData()

    On Error GoTo ErrHandler
    
    ' Verify we have a valid workbook
    If Workbooks.count = 0 Then Exit Sub
    
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    
    Dim issues As Long
    
    ' Verify this is an RDD workbook
    If Not modMain.IsRDDWorkbook(wb) Then
        MsgBox "This is not an RDD workbook.", vbExclamation, modMain.AppProjectName
        Exit Sub
    End If
    
    ' Check if there are any room sheets to validate
    Dim hasRoomSheets As Boolean
    Dim ws As Worksheet
    
    hasRoomSheets = False
    For Each ws In wb.Worksheets
        If modRooms.IsRoomSheet(ws) Then
            hasRoomSheets = True
            Exit For
        End If
    Next ws
    
    If Not hasRoomSheets Then
        MsgBox "No room sheets found to validate.", vbInformation, modMain.AppProjectName
        Exit Sub
    End If
    
    frmWait.ShowDialog
    modUtil.HideOpMode True
    
    ' Run validation
    issues = modRooms.ValidateRooms(wb)
    clsState.RoomsValidationIssueCount = issues
    
    If issues = 0 Then
        ' Set validation status in state
        clsState.RoomsValidated = True
    End If
    
    ' Invalidate buttons that depend on validation
    clsState.InvalidateControl RIBBON_BTN_BUILD_DATA
        
    modUtil.HideOpMode False
    frmWait.Hide
    
    Exit Sub
    
ErrHandler:
    modErr.ReportError "modMain.ValidateRoomData", Err.Number, Erl, caption:=modMain.AppProjectName
    
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


' -----------------------------------------------------------------------------------
' Function  : DetermineChangeCategory
' Purpose   : Determines which category of monitored range was changed
'
' Parameters:
'   targetSheet  [Worksheet] - The worksheet containing the ranges
'   changedRange [Range]     - The range that was changed
'
' Returns   : [ChangeCategory] - Category of the changed range
'
' Notes     :
'   - Checks single cells first (faster), then ranges
' -----------------------------------------------------------------------------------
Private Function DetermineChangeCategory( _
    ByVal targetSheet As Worksheet, _
    ByVal changedRange As Range) As ChangeCategory
    
    On Error GoTo ErrHandler
    
    ' Check Parallax cell
    If modRanges.IntersectsNamedCell(targetSheet, changedRange, NAME_CELL_PARALLAX) Then
        DetermineChangeCategory = CC_Parallax
        Exit Function
    End If
    
    ' Check Room Metadata cells
    If modRanges.IntersectsNamedCell(targetSheet, changedRange, NAME_CELL_ROOM_ID) Or _
        modRanges.IntersectsNamedCell(targetSheet, changedRange, NAME_CELL_ROOM_NO) Or _
        modRanges.IntersectsNamedCell(targetSheet, changedRange, NAME_CELL_ROOM_ALIAS) Then
        DetermineChangeCategory = CC_RoomMetadata
        Exit Function
    End If
    
    ' Check Scene Metadata cells
    If modRanges.IntersectsNamedCell(targetSheet, changedRange, NAME_CELL_SCENE_ID) Then
        DetermineChangeCategory = CC_SceneMetadata
        Exit Function
    End If
    
    ' Check General Room Settings cells
    If modRanges.IntersectsNamedCell(targetSheet, changedRange, NAME_CELL_GAME_HEIGHT) Or _
        modRanges.IntersectsNamedCell(targetSheet, changedRange, NAME_CELL_GAME_WIDTH) Or _
        modRanges.IntersectsNamedCell(targetSheet, changedRange, NAME_CELL_BG_HEIGHT) Or _
        modRanges.IntersectsNamedCell(targetSheet, changedRange, NAME_CELL_BG_WIDTH) Or _
        modRanges.IntersectsNamedCell(targetSheet, changedRange, NAME_CELL_UI_HEIGHT) Then
        DetermineChangeCategory = CC_GeneralSettings
        Exit Function
    End If
    
    ' === Check Table Ranges ===
    
    ' Actors
    If modRanges.IntersectsNamedRange(targetSheet, changedRange, NAME_RANGE_ACTORS_ACTOR_NAME) Or _
        modRanges.IntersectsNamedCell(targetSheet, changedRange, NAME_RANGE_ACTORS_ACTOR_ID) Then
        DetermineChangeCategory = CC_Actors
        Exit Function
    End If
    
    ' Sounds
    If modRanges.IntersectsNamedRange(targetSheet, changedRange, NAME_RANGE_SOUNDS_DESCRIPTION) Or _
        modRanges.IntersectsNamedRange(targetSheet, changedRange, NAME_RANGE_SOUNDS_SOUND_ID) Then
        DetermineChangeCategory = CC_Sounds
        Exit Function
    End If
    
    ' Special FX
    If modRanges.IntersectsNamedRange(targetSheet, changedRange, NAME_RANGE_SPECIAL_FX_DESCRIPTION) Or _
        modRanges.IntersectsNamedRange(targetSheet, changedRange, NAME_RANGE_SPECIAL_FX_ANIMATION_ID) Then
        DetermineChangeCategory = CC_SpecialFX
        Exit Function
    End If
    
    ' Touchable Objects
    If modRanges.IntersectsNamedRange(targetSheet, changedRange, NAME_RANGE_TOUCHABLE_OBJECTS_HOTSPOT_NAME) Or _
        modRanges.IntersectsNamedRange(targetSheet, changedRange, NAME_RANGE_TOUCHABLE_OBJECTS_HOTSPOT_ID) Then
        DetermineChangeCategory = CC_TouchableObjects
        Exit Function
    End If
    
    ' Pickupable Objects
    If modRanges.IntersectsNamedRange(targetSheet, changedRange, NAME_RANGE_PICKUPABLE_OBJECTS_ITEM_ID) Or _
        modRanges.IntersectsNamedRange(targetSheet, changedRange, NAME_RANGE_PICKUPABLE_OBJECTS_NAME) Then
        DetermineChangeCategory = CC_PickupableObjects
        Exit Function
    End If
    
    ' Multi-State Objects
    If modRanges.IntersectsNamedRange(targetSheet, changedRange, NAME_RANGE_MULTI_STATE_OBJECTS_OBJECT_NAME) Or _
        modRanges.IntersectsNamedRange(targetSheet, changedRange, NAME_RANGE_MULTI_STATE_OBJECTS_STATE_ID) Or _
        modRanges.IntersectsNamedRange(targetSheet, changedRange, NAME_RANGE_MULTI_STATE_OBJECTS_STATE) Then
        DetermineChangeCategory = CC_MultiStateObjects
        Exit Function
    End If
    
    ' Flags
    If modRanges.IntersectsNamedRange(targetSheet, changedRange, NAME_RANGE_FLAGS_FLAG_ID) Or _
        modRanges.IntersectsNamedRange(targetSheet, changedRange, NAME_RANGE_FLAGS_DESCRIPTION) Or _
        modRanges.IntersectsNamedRange(targetSheet, changedRange, NAME_RANGE_FLAGS_BOOL_TYPE) Then
        DetermineChangeCategory = CC_Flags
        Exit Function
    End If
    
    ' Default: no specific category
    DetermineChangeCategory = CC_None
    Exit Function
    
ErrHandler:
    ' On error: return no specific category
    DetermineChangeCategory = CC_None
End Function



