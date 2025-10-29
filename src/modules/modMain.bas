Attribute VB_Name = "modMain"
' ===================================================================================
' Module    : modMain
' Purpose   : Orchestrates use cases; handles UI, validation, selection, confirmation,
'             navigation, logging, error display, and app lifecycle.
'
' Notes     :
'   - Keep this module focused on orchestration and UI flows.
'   - Business logic should live in feature modules/classes.
' ===================================================================================

Option Explicit
Option Private Module

' ===== Private State =================================================================
Private m_objLog As clsLog          ' central Logger
Private m_strAppTempPath As String
Private m_strAppProjectName As String

Private m_objAppEvents As clsAppEvents ' keeps WithEvents sink alive
Private m_wbActiveWorkbook As Workbook ' holds ActiveWorkbook on install

' ===== Public API ====================================================================

' -----------------------------------------------------------------------------------
' Function  : AppProjectName (Get)
' Purpose   : Returns the VBA project name.
' Parameters: (none)
' Returns   : String - Project name
' Notes     : Ensure that SetAppProjectName was executed before the first query.
' -----------------------------------------------------------------------------------
Public Property Get AppProjectName() As String
    AppProjectName = m_strAppProjectName
End Property

' -----------------------------------------------------------------------------------
' Function  : AppTempPath (Get/Let)
' Purpose   : Gets/Sets the path used for temp/log files.
' Parameters: (none)
' Returns   : String (Get)
' Notes     : Ensure trailing "\" when setting.
' -----------------------------------------------------------------------------------
Public Property Get AppTempPath() As String

    AppTempPath = m_strAppTempPath

End Property

Public Property Let AppTempPath(ByVal strNewValue As String)

    ' Ensure trailing backslash
    If Len(strNewValue) > 0 Then
        If Right$(strNewValue, 1) <> "\" Then
            strNewValue = strNewValue & "\"
        End If
    End If
    m_strAppTempPath = strNewValue

End Property

' -----------------------------------------------------------------------------------
' Function  : AppVersion (Get)
' Purpose   : Returns version string from a custom document property.
' Parameters: (none)
' Returns   : String - e.g., "1.2.3"
' Notes     : Uses GetDocumentPropertyValue("RDD_AddInVersion").
' -----------------------------------------------------------------------------------
Public Property Get AppVersion() As String
    AppVersion = GetDocumentPropertyValue(ThisWorkbook, "RDD_AddInVersion", "0.0.0")
End Property

' -----------------------------------------------------------------------------------
' Procedure : AppInstall
' Purpose   : Initializes application-specific settings and resources during first-time add-in installation.
'             Setts default properties, creating required named ranges,
'             registering document tags, or preparing the workbook for use with the add-in.
'             must called from Workbook_AddinInstall().
' Parameters: (none)
' Returns   : (none)
' Notes     : Used so Workbook_Open can reference the previous ActiveWorkbook.
' -----------------------------------------------------------------------------------
Public Sub AppInstall()
'If the add-in is activated when a workbook is opened,
'save the reference to this workbook for Workbook_Open,
'since the add-in is set as the active workbook in Workbook_Open.
If Not ActiveWorkbook Is Nothing Then Set m_wbActiveWorkbook = ActiveWorkbook
End Sub
' -----------------------------------------------------------------------------------
' Function  : AppStart
' Purpose   : Application startup: init logging, wire App events, init state, refresh UI,
'             validating workbook structure.
'             must called from Workbook_open()
' Parameters: (none)
' Returns   : (none)
' Notes     : Requires clsAppEvents and clsState
' -----------------------------------------------------------------------------------
Public Sub AppStart()
    
    ' Ensure temp path exists before logging
    m_strAppTempPath = modUtil.GetTempFolder & "\BYTE RANGER"
    If Dir(m_strAppTempPath, vbDirectory) = "" Then MkDir m_strAppTempPath
    m_strAppTempPath = m_strAppTempPath & "\" & AppProjectName & "\"
    If Dir(m_strAppTempPath, vbDirectory) = "" Then MkDir m_strAppTempPath
        
    '
    SetAppProjectName
    
    ' Logger
    Call OpenLog
    
    ' load options
    modOptions.ReadGeneralOptions

    ' wire application events when running as add-in
    If RDDAddInWkBk.IsAddin Then
        ConnectEventHandler
    End If

    ' init State
    clsState.Init
    clsState.InvalidateRibbon
End Sub

' -----------------------------------------------------------------------------------
' Function  : AppStop
' Purpose   : Application shutdown: saving settings, releasing resources, unhook events, cleanup state, close log.
'             must called from Workbook_BeforeClose()
' Parameters: (none)
' Returns   : (none)
' Notes     : Safe to call multiple times.
' -----------------------------------------------------------------------------------
Public Sub AppStop()

    Call SaveGeneralOptions
    
    ' Detach events
    DisconnectEventHandler

    '  Update ribbon/UI and clear state
    clsState.InvalidateRibbon
    clsState.Cleanup

    ' close Log
    If Not m_objLog Is Nothing Then
        m_objLog.CloseLog
    End If
    Set m_objLog = Nothing
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
' Parameters:
'   wb [Workbook] - Workbook to test.
' Returns   : Boolean - True when tag matches, otherwise False.
' Notes     : Uses APP_DOC_TAG_KEY/APP_DOC_TAG_VAL; defensive on Nothing.
' -----------------------------------------------------------------------------------
Public Function IsRDDWorkbook(ByVal wb As Workbook) As Boolean
    If wb Is Nothing Then Exit Function

    Dim strVal As String
    strVal = modProps.GetDocumentPropertyValue(wb, APP_DOC_TAG_KEY, "")
    IsRDDWorkbook = (StrComp(strVal, APP_DOC_TAG_VAL, vbBinaryCompare) = 0)
End Function

' ===== Ribbon Callback Targets =======================================================

' -----------------------------------------------------------------------------------
' Procedure : ShowLog
' Purpose   : Displays the current log to the user. Ensures a logger instance exists.
'
' Parameters: (none)
' Returns   : (none)
'
' Notes     :
'   - Delegates UI/display to clsLog.ShowLog.
' -----------------------------------------------------------------------------------
Public Sub ShowLog()
    If m_objLog Is Nothing Then
        Call OpenLog
    End If
    m_objLog.ShowLog
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

    On Error GoTo errHandler
    
    Dim strPath As String
    strPath = ReplaceWildcards(modOptions.Opt_ManualPath) & "\"  ' TODO: Add options Formular


    If Dir(strPath & FILENAME_MANUAL) <> "" Then
        ThisWorkbook.FollowHyperlink strPath & FILENAME_MANUAL
    Else
        ' Datei xxxx nicht gefunden
        MsgBox "File " & strPath & FILENAME_MANUAL & _
            " not found!", vbExclamation, AppProjectName
    End If
    

    On Error GoTo 0
    Exit Sub

errHandler:
    Dim lngErr As Long
    lngErr = Err.Number
    MsgBox "Error " & lngErr & " (" & Err.Description & ") in procedure ShowManual, line " & Erl & ".", vbCritical, AppProjectName
    LogError "ShowManual", lngErr, Erl

End Sub

' -----------------------------------------------------------------------------------
' Procedure : ShowOptions
' Purpose   : Displays the options form, passing the active workbook context.
' Parameters: (none)
' Returns   : (none)
' Notes     : Form lifetime is scoped to the procedure.
' -----------------------------------------------------------------------------------
Public Sub ShowOptions()

    On Error GoTo errHandler
    
    Dim udtCur As tOptions
    udtCur = modOptions.GetAllOptions
        
    Dim objActWb As Workbook: Set objActWb = ActiveWorkbook
    
    Dim fOptions As frmOptions
    Set fOptions = New frmOptions: fOptions.Init AppProjectName, udtCur
        
    fOptions.Show
        

    If fOptions.Confirmed Then
        Dim udtNew As tOptions
        udtNew = fOptions.ResultOptions

        Dim strValErr As String
        strValErr = modOptions.ValidateOptions(udtNew)
        If LenB(strValErr) > 0 Then
            MsgBox strValErr, vbExclamation, AppProjectName
            Set fOptions = Nothing
            Exit Sub
        End If

        modOptions.SetAllOptions udtNew
        modOptions.SaveGeneralOptions
        modOptions.SaveWorkbookOptions objActWb
    End If
    
    Set fOptions = Nothing
    
    On Error GoTo 0
    Exit Sub

errHandler:
    Dim lngErr As Long
    lngErr = Err.Number
    MsgBox "Error " & lngErr & " (" & Err.Description & ") in procedure ShowOptions, line " & Erl & ".", vbCritical, AppProjectName
    LogError "ShowOptions", lngErr, Erl

End Sub

' -----------------------------------------------------------------------------------
' Procedure : ShowAbout
' Purpose   : Displays the about dialog.
' Parameters: (none)
' Returns   : (none)
' Notes     : Form lifetime is scoped to the procedure.
' -----------------------------------------------------------------------------------
Public Sub ShowAbout()
    On Error GoTo errHandler
    
    Dim objActWb As Workbook: Set objActWb = ActiveWorkbook
    Dim fAbout As frmAbout
    
    Set fAbout = New frmAbout
    
    fAbout.Show
    
    Set fAbout = Nothing
    
    On Error GoTo 0
    Exit Sub
        
errHandler:
    Dim lngErr As Long
    lngErr = Err.Number
    MsgBox "Error " & lngErr & " (" & Err.Description & ") in procedure ShowAbout, line " & Erl & ".", _
           vbCritical, AppProjectName
    LogError "ShowAbout", lngErr, Erl
End Sub
' -----------------------------------------------------------------------------------
' Function  : AddNewRoom
' Purpose   : Runs the "New Room Sheet" dialog, creates the sheet via modRooms,
'             post-processes visuals, and navigates optionally.
' Parameters:
'   blnGotoNewRoom [Boolean] - If True, jumps to A1 of the created sheet.
' Returns   : String - The created Room ID (empty if cancelled).
' Notes     : Uses frmNewItem and modRooms.
' -----------------------------------------------------------------------------------
Public Function AddNewRoom(Optional ByVal blnGotoNewRoom As Boolean = True) As String
    On Error GoTo errHandler

    Dim objActWks As Worksheet: Set objActWks = ActiveSheet
    Dim objActWb As Workbook: Set objActWb = ActiveWorkbook
    Dim objActCell As Range: Set objActCell = ActiveCell

    Dim objNewWkSh As Worksheet
    Dim lngIdx As Long
    Dim strID As String
    
    Dim fNewItem As frmNewItem: Set fNewItem = New frmNewItem
            
    Application.StatusBar = False
    
    With fNewItem
        .FormCaption = "New Room Sheet"
        .NameLabel = "Room Name"
        .IDLabel = "Room ID"
        .IDVisible = True
        lngIdx = modRooms.GetNextRoomIndex(objActWb)
        strID = modRooms.GetFormattedRoomID(lngIdx)
        .IDText = strID
        .NameText = strID

        .Show                       ' modal
        If Not .Cancelled Then
            
            EnsureWorkbookIsTagged objActWb
     
            Set objNewWkSh = modRooms.AddRoom(objActWb, .NameText, lngIdx)
            If Not objNewWkSh Is Nothing Then
                modUtil.HideOpMode True
                modRooms.ApplyParallaxRangeCover objNewWkSh
                If blnGotoNewRoom Then
                    Application.GoTo objNewWkSh.Range("A1"), True
                Else
                    objActWks.Activate
                    If Not objActCell Is Nothing Then objActCell.Select
                End If
                modUtil.HideOpMode False
                AddNewRoom = strID
            End If
        End If
        Unload fNewItem
    End With
                        
    Set fNewItem = Nothing
                
    On Error GoTo 0
    Exit Function
    
errHandler:
    Dim lngErr As Long
    lngErr = Err.Number
    MsgBox "Error " & lngErr & " (" & Err.Description & ") in procedure AddNewRoom, line " & Erl & ".", vbCritical, AppProjectName
    LogError "AddNewRoom", lngErr, Erl
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
    On Error GoTo errHandler
    
    Dim rngCell As Range: Set rngCell = ActiveCell
    
    Dim strRoomID As String
    
    strRoomID = AddNewRoom(False)
        
    If Len(strRoomID) > 0 Then
        If Not rngCell Is Nothing Then rngCell.Value = strRoomID
    End If
    
    
errHandler:
    Dim lngErr As Long
    lngErr = Err.Number
    MsgBox "Error " & lngErr & " (" & Err.Description & ") in procedure AddNewRoomFromCellCtxMnu, line " & Erl & ".", vbCritical, AppProjectName
    LogError "AddNewRoomFromCellCtxMnu", lngErr, Erl
End Sub

' -----------------------------------------------------------------------------------
' Procedure : RemoveCurrentRoom
' Purpose   : Deletes the active room sheet after confirmation and safety checks.
' Parameters: (none)
' Returns   : (none)
' Notes     : Delegates the deletion to modRooms.RemoveRoom.
' -----------------------------------------------------------------------------------
Public Sub RemoveCurrentRoom()
    On Error GoTo errHandler

    Dim wks As Worksheet
    Set wks = ActiveSheet

    Application.StatusBar = False
    
    If Not modRooms.IsRoomSheet(wks) Then
        MsgBox "Active sheet is not a 'Room' sheet.", vbInformation, AppProjectName
        Exit Sub
    End If
    
    ' Confirm with the user
    If MsgBox("Are you sure you want to delete the sheet '" & wks.Name & "'?" & vbCrLf & _
        "This action cannot be undone.", vbYesNo + vbExclamation, "Confirm Sheet Deletion") <> vbYes Then
        Application.StatusBar = "Deletion cancelled."
        Exit Sub
    End If

    Call modRooms.RemoveRoom(wks)

    Exit Sub
errHandler:
    LogError "RemoveCurrentRoom", Err.Number, Erl
    MsgBox "Error " & Err.Number & " (" & Err.Description & ")", vbCritical, AppProjectName
End Sub

' -----------------------------------------------------------------------------------
' Procedure : GotoRoomFromCell
' Purpose   : Jumps to the room sheet referenced by the active cell value.
' Parameters: (none)
' Returns   : (none)
' Notes     : Requires room sheets to be discoverable via modRooms.HasRoomSheet.
' -----------------------------------------------------------------------------------
Public Sub GotoRoomFromCell()
    On Error GoTo errHandler
    
    Dim strRoomID As String
    Dim wb As Workbook: Set wb = ActiveWorkbook
    Dim rngActCell As Range: Set rngActCell = ActiveCell
    
    strRoomID = Trim$(CStr(rngActCell.Value))
    If Len(strRoomID) = 0 Then
        MsgBox "No Room ID in the selected cell.", vbInformation, AppProjectName
        Exit Sub
    End If
    
    Dim wks As Worksheet
    If modRooms.HasRoomSheet(wb, strRoomID, wks) Then
        Application.GoTo wks.Range("A1"), True
        Exit Sub
    End If
    
    MsgBox "Room '" & strRoomID & "' not found.", vbInformation, AppProjectName
    Exit Sub

errHandler:
    Dim lngErr As Long: lngErr = Err.Number
    MsgBox "Error " & lngErr & " (" & Err.Description & ") in procedure GotoRoomFromCell, line " & Erl & ".", vbCritical, AppProjectName
    LogError "GotoRoomFromCell", lngErr, Erl
End Sub

' -----------------------------------------------------------------------------------
' Procedure : LogError
' Purpose   : Writes an error entry to the central log. Ensures the logger instance
'             is initialized (opens the log on first use).
'
' Parameters:
'   sNameFunc [String]   - Name of the procedure where the error occurred
'   iErrNum   [Long]     - (Optional) VBA Err.Number
'   iErrLine  [Integer]  - (Optional) Line number (Erl)
'
' Returns   : (none)
'
' Notes     :
'   - Safe to call in any error Handler block.
'   - Calls OpenLog on demand; forwards to clsLog.WriteError.
' -----------------------------------------------------------------------------------
Public Sub LogError(ByVal sNameFunc As String, Optional iErrNum As Long = -9999, Optional iErrLine As Integer = -1)
    On Error Resume Next
    If m_objLog Is Nothing Then
        Call OpenLog
    End If
    
    m_objLog.WriteError sNameFunc, iErrNum, iErrLine
    On Error GoTo 0
End Sub

' ===== Private Methods ===============================================================

' -----------------------------------------------------------------------------------
' Function  : ConnectEventHandler
' Purpose   : Enables application-level event handling by assigning the Excel
'             Application object to the clsAppEvents instance.
'
' Parameters:
'   (none)
'
' Returns   :
'
' Notes     :
'   - Requires a class module `clsAppEvents` exposing an `App` property (WithEvents).
'   - Uses `On Error Resume Next`; shows a critical message box on failure
'     ("Unable to enable application events.", title "RDD Add-In").
'   - Keeps a private instance alive in this module.
' -----------------------------------------------------------------------------------
Private Function ConnectEventHandler() As Boolean

    On Error GoTo errHandler
    
    If m_objAppEvents Is Nothing Then Set m_objAppEvents = New clsAppEvents
    Set m_objAppEvents.App = Application
    ConnectEventHandler = True
    Exit Function
    
errHandler:
    On Error Resume Next
    LogError "ConnectEventHandler", Err.Number, Erl
    
    MsgBox "Unable to enable application events." & vbCrLf & _
        "Error " & Err.Number & ": " & Err.Description, _
        vbOKOnly Or vbCritical, AppProjectName
           
    ' Ensure instance is detached/cleared on failure
    If Not m_objAppEvents Is Nothing Then Set m_objAppEvents.App = Nothing
    Set m_objAppEvents = Nothing
End Function

' -----------------------------------------------------------------------------------
' Function  : DisconnectEventHandler
' Purpose   : Disables application-level event handling by releasing the reference
'             to the Excel Application object.
'
' Parameters:
'   (none)
'
' Returns   :
'
' Notes     :
'   - Safe to call multiple times; sets `AppEvents.App = Nothing`.
'   - Clears instance reference so events stop firing.
' -----------------------------------------------------------------------------------
Private Sub DisconnectEventHandler()
    On Error Resume Next
    If Not m_objAppEvents Is Nothing Then Set m_objAppEvents.App = Nothing
    Set m_objAppEvents = Nothing
    On Error GoTo 0
End Sub

' -----------------------------------------------------------------------------------
' Procedure : OpenLog
' Purpose   : Creates the central logger instance and opens the log file with a
'             standard header (project name + version).
'
' Parameters:
'   (none)
'
' Returns   : (none)
'
' Notes     :
'   - Uses AppTempPath as base directory (must exist).
'   - Sets file name to "<Project>_Log".
'   - Private helper intended to be called by startup or ShowLog.
' -----------------------------------------------------------------------------------
Private Sub OpenLog()
    Set m_objLog = New clsLog
    m_objLog.PathLogFile = AppTempPath
    m_objLog.LogFilename = AppProjectName & "_Log"
    m_objLog.OpenLog AppProjectName & " " & AppVersion
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
    m_strAppProjectName = Err.Source
    On Error GoTo 0
End Sub
