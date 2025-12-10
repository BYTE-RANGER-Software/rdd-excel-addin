VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSearchResults 
   Caption         =   "Search Results"
   ClientHeight    =   5467
   ClientLeft      =   99
   ClientTop       =   429
   ClientWidth     =   8206.001
   OleObjectBlob   =   "frmSearchResults.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmSearchResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ====================================================================================
' UserForm : frmSearchResults
' Purpose  : Displays search results with navigation capability.
'            Double-click on a result to navigate to the source cell.
'
' Usage:
'   Dim frm As frmSearchResults
'   Set frm = New frmSearchResults
'   frm.Initialize "i_key", results, ActiveWorkbook
'   frm.Show vbModeless
'
' Dependencies:
'   - modSearch : NavigateToSearchResult
'
' Notes:
'   - Results are displayed in a ListBox with columns: Sheet, Cell, Context, Value
'   - Double-click navigates to the cell
'   - Form is modeless to allow interaction with the workbook
'   - Attribute VB_PredeclaredId = False
' ====================================================================================
Option Explicit

Private m_searchTerm As String
Private m_results As Collection
Private m_targetBook As Workbook

' ====================================================================================
' Public API
' ====================================================================================

' -----------------------------------------------------------------------------------
' Procedure : Initialize
' Purpose   : Initializes the form with search results.
'
' Parameters:
'   searchTerm  [String]     - The term that was searched for
'   results     [Collection] - Collection of result arrays
'   targetBook  [Workbook]   - Workbook the results belong to
' -----------------------------------------------------------------------------------
Public Sub Initialize(ByVal searchTerm As String, _
    ByVal results As Collection, _
    ByVal targetBook As Workbook)
    
    m_searchTerm = searchTerm
    Set m_results = results
    Set m_targetBook = targetBook
    
    ' Update caption
    Me.caption = "Search Results for '" & searchTerm & "' (" & results.count & " found)"
    
    ' Populate ListBox
    PopulateResults
End Sub

' ====================================================================================
' Form Events
' ====================================================================================

Private Sub UserForm_Initialize()
    ' Set up ListBox columns
    With lstResults
        .ColumnCount = 4
        .ColumnWidths = "100;60;120;150"
        .ColumnHeads = False
        .Font.name = "Courier New"
        .Font.Size = 8
    End With
    
    ' Add header row manually (ListBox doesn't support true headers)
    lblHeader.caption = "Sheet" & Space(20) & "Cell" & Space(7) & "Context" & Space(25) & "Value"
    lblHeader.Font.name = "Courier New"
    lblHeader.Font.Size = 8
    lblHeader.Font.Bold = True
    
    cmdNavigate.enabled = False
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Clean up references
    Set m_results = Nothing
    Set m_targetBook = Nothing
End Sub

' ====================================================================================
' Control Events
' ====================================================================================

Private Sub lstResults_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    NavigateToSelected
End Sub

Private Sub cmdNavigate_Click()
    NavigateToSelected
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

' ====================================================================================
' Private Helpers
' ====================================================================================

' -----------------------------------------------------------------------------------
' Procedure : PopulateResults
' Purpose   : Fills the ListBox with search results.
' -----------------------------------------------------------------------------------
Private Sub PopulateResults()
    Dim resultData As Variant
    Dim i As Long
    
    lstResults.Clear
    
    If m_results Is Nothing Then Exit Sub
    If m_results.count = 0 Then
        lstResults.AddItem "No results found."
        cmdNavigate.enabled = False
        Exit Sub
    End If
    
    cmdNavigate.enabled = True
    
    For i = 1 To m_results.count
        resultData = m_results(i)
        
        With lstResults
            .AddItem resultData(0)           ' Sheet name
            .list(.ListCount - 1, 1) = resultData(1)  ' Cell address
            .list(.ListCount - 1, 2) = resultData(2)  ' Context
            .list(.ListCount - 1, 3) = resultData(3)  ' Value
        End With
    Next i
    
    ' Select first item
    If lstResults.ListCount > 0 Then
        lstResults.ListIndex = 0
    End If
End Sub

' -----------------------------------------------------------------------------------
' Procedure : NavigateToSelected
' Purpose   : Navigates to the cell of the selected result.
' -----------------------------------------------------------------------------------
Private Sub NavigateToSelected()
    On Error GoTo ErrHandler
    
    If lstResults.ListIndex < 0 Then
        MsgBox "Please select a result first.", vbInformation, "Search Results"
        Exit Sub
    End If
    
    Dim sheetName As String
    Dim cellAddr As String
    
    sheetName = lstResults.list(lstResults.ListIndex, 0)
    cellAddr = lstResults.list(lstResults.ListIndex, 1)
    
    ' Validate we have valid data
    If Len(sheetName) = 0 Or Len(cellAddr) = 0 Then Exit Sub
    If sheetName = "No results found." Then Exit Sub
    
    ' Navigate using modSearch
    modSearch.NavigateToSearchResult sheetName, cellAddr, m_targetBook
    
    Exit Sub
    
ErrHandler:
    MsgBox "Could not navigate to the selected cell." & vbCrLf & _
           "Error: " & Err.Description, vbExclamation, "Navigation Error"
End Sub


