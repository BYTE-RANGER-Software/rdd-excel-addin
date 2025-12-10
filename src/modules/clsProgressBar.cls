VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' Class: clsProgressBar
' Purpose: Reusable progress bar for Excel status bar
' Author: [Your Name]
' Date: 2024-12-10
'==============================================================================
Option Explicit

' Private member variables
Private m_TotalSteps As Long
Private m_CurrentStep As Long
Private m_BarLength As Long
Private m_Prefix As String
Private m_ShowPercentage As Boolean
Private m_ShowBar As Boolean
Private m_OriginalStatusBar As Variant

' Property: TotalSteps - Set total number of steps
Public Property Let totalSteps(ByVal value As Long)
    m_TotalSteps = value
End Property

Public Property Get totalSteps() As Long
    totalSteps = m_TotalSteps
End Property

' Property: CurrentStep - Get current step
Public Property Get CurrentStep() As Long
    CurrentStep = m_CurrentStep
End Property

' Property: BarLength - Set length of progress bar in characters
Public Property Let BarLength(ByVal value As Long)
    m_BarLength = value
End Property

' Property: Prefix - Set prefix text before progress bar
Public Property Let Prefix(ByVal value As String)
    m_Prefix = value
End Property

' Property: ShowPercentage - Show percentage display
Public Property Let ShowPercentage(ByVal value As Boolean)
    m_ShowPercentage = value
End Property

' Property: ShowBar - Show graphical bar
Public Property Let ShowBar(ByVal value As Boolean)
    m_ShowBar = value
End Property

'==============================================================================
' Initialize class with default values
'==============================================================================
Private Sub Class_Initialize()
    m_TotalSteps = 100
    m_CurrentStep = 0
    m_BarLength = 40
    m_Prefix = "Fortschritt"
    m_ShowPercentage = True
    m_ShowBar = True
    
    ' Save original status bar state
    m_OriginalStatusBar = Application.DisplayStatusBar
    Application.DisplayStatusBar = True
End Sub

'==============================================================================
' Start: Initialize progress bar with total steps
'==============================================================================
Public Sub Start(ByVal totalSteps As Long, Optional ByVal prefixText As String = "")
    m_TotalSteps = totalSteps
    m_CurrentStep = 0
    
    If prefixText <> "" Then
        m_Prefix = prefixText
    End If
    
    Application.ScreenUpdating = False
    UpdateDisplay
End Sub

'==============================================================================
' Increment: Advance progress by one step
'==============================================================================
Public Sub Increment(Optional ByVal customMessage As String = "")
    m_CurrentStep = m_CurrentStep + 1
    If m_CurrentStep > m_TotalSteps Then m_CurrentStep = m_TotalSteps
    UpdateDisplay customMessage
End Sub

'==============================================================================
' SetProgress: Set progress to specific step
'==============================================================================
Public Sub SetProgress(ByVal stepNumber As Long, Optional ByVal customMessage As String = "")
    m_CurrentStep = stepNumber
    If m_CurrentStep > m_TotalSteps Then m_CurrentStep = m_TotalSteps
    If m_CurrentStep < 0 Then m_CurrentStep = 0
    UpdateDisplay customMessage
End Sub

'==============================================================================
' UpdateDisplay: Update the status bar display
'==============================================================================
Private Sub UpdateDisplay(Optional ByVal customMessage As String = "")
    Dim percentComplete As Double
    Dim progressBar As String
    Dim statusText As String
    
    If m_TotalSteps = 0 Then Exit Sub
    
    percentComplete = m_CurrentStep / m_TotalSteps
    
    ' Build status text
    statusText = m_Prefix & ": "
    
    ' Add custom message if provided
    If customMessage <> "" Then
        statusText = statusText & customMessage & " "
    End If
    
    ' Add graphical bar if enabled
    If m_ShowBar Then
        progressBar = String(Int(percentComplete * m_BarLength), "¦") & _
                     String(m_BarLength - Int(percentComplete * m_BarLength), "¦")
        statusText = statusText & "[" & progressBar & "] "
    End If
    
    ' Add percentage if enabled
    If m_ShowPercentage Then
        statusText = statusText & Format(percentComplete, "0%") & _
                    " (" & m_CurrentStep & "/" & m_TotalSteps & ")"
    End If
    
    Application.StatusBar = statusText
    DoEvents ' Allow UI to update
End Sub

'==============================================================================
' Complete: Finish progress and reset status bar
'==============================================================================
Public Sub Complete(Optional ByVal finalMessage As String = "Abgeschlossen!")
    m_CurrentStep = m_TotalSteps
    
    If finalMessage <> "" Then
        Application.StatusBar = finalMessage
        Application.Wait Now + TimeValue("00:00:01") ' Show for 1 second
    End If
    
    Reset
End Sub

'==============================================================================
' Reset: Reset status bar to original state
'==============================================================================
Public Sub Reset()
    Application.StatusBar = False
    Application.ScreenUpdating = True
End Sub

'==============================================================================
' Class_Terminate: Ensure status bar is reset when object is destroyed
'==============================================================================
Private Sub Class_Terminate()
    Reset
End Sub

