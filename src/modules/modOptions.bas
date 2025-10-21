Attribute VB_Name = "modOptions"
Option Explicit
Option Private Module

Private Type g_tGeneralOptions
    sManualPath As String
End Type

Private m_typGeneralOptions As g_tGeneralOptions

Private Const REG_APP_NAME As String = "RDD-AddIn"
Private Const REG_SECTION_GENERAL As String = "General"

Private Const REG_MANUAL_PATH As String = "ManualPath"

Private m_blnOptionsChanged As Boolean ' Have the Opt_... settings been changed?

Public Property Get OptionsChanged() As Boolean

    OptionsChanged = m_blnOptionsChanged

End Property

Public Property Let OptionsChanged(ByVal blnNewValue As Boolean)

    m_blnOptionsChanged = blnNewValue

End Property

Public Property Get Opt_ManualPath() As String

    Opt_ManualPath = m_typGeneralOptions.sManualPath

End Property

Public Property Let Opt_ManualPath(ByVal strNewValue As String)

    If strNewValue <> m_typGeneralOptions.sManualPath Then
        m_typGeneralOptions.sManualPath = strNewValue
        OptionsChanged = True
    End If

End Property

Private Function GetDefaultManualPath() As String
    GetDefaultManualPath = modConst.WILDCARD_MY_DOCUMENTS & "\" & AppProjectName & "\Doku"
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub  : ReadGeneralOptions - Reads the general settings for the tool from the registry

'* Created    : 14.04.2023
'* Author     : thiemann
'* Contacts   :
'* Copyright  : BYTE RANGER Software
'* Argument(s):                                           Description
'*
'*
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub ReadGeneralOptions()
    ' read in all parameters
        
    m_typGeneralOptions.sManualPath = GetSetting(REG_APP_NAME, REG_SECTION_GENERAL, REG_MANUAL_PATH, GetDefaultManualPath())
 
    OptionsChanged = False
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub  : ReadWorkbookOptions - Reads the settings for the tool from the user properties of the current workbook

'* Created    : 14.04.2023
'* Author     : thiemann
'* Contacts   :
'* Copyright  : BYTE RANGER Software
'* Argument(s):                                           Description
'*
'* objWkBk As Workbook: The workbook from which the user properties are to be read.
'*
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub ReadWorkbookOptions(ByVal objWkBk As Workbook)
    ' read in all parameters
        
    'Opt_... = modProps.GetDocumentPropertyValue(objWkBk, PROP_..., ...)
    'Opt_... = modProps.GetDocumentPropertyValue(objWkBk, PROP_..., ...)
    OptionsChanged = False
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub  : SaveGeneralOptions - Saves the general settings for the tool to the registry
'* Created    : 14.04.2023
'* Author     : thiemann
'* Contacts   :
'* Copyright  : BYTE RANGER Software
'* Argument(s):                                           Description
'*
'* blnIgnoreChangeFlag As Boolean: specifies whether OptionsChanged should be ignored
'*
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub SaveGeneralOptions(Optional blnIgnoreChangeFlag As Boolean = False)
    
    If OptionsChanged Or blnIgnoreChangeFlag Then
        
        SaveSetting REG_APP_NAME, REG_SECTION_GENERAL, REG_MANUAL_PATH, m_typGeneralOptions.sManualPath
        
    End If
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub  : SaveGeneralOptions - Saves the general settings for the tool to the registry
'* Created    : 14.04.2023
'* Author     : thiemann
'* Contacts   :
'* Copyright  : BYTE RANGER Software
'* Argument(s):                                           Description
'*
'* objWkBk As Workbook: The workbook to which the user properties are be written.
'*
'* blnIgnoreChangeFlag As Boolean: specifies whether OptionsChanged should be ignored
'*
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub SaveWorkbookOptions(ByVal objWkBk As Workbook, Optional blnIgnoreChangeFlag As Boolean = False)
    
    If OptionsChanged Or blnIgnoreChangeFlag Then
        
        'modProps.SetDocumentProperty objWkBk, PROP_..., Opt_...
        'modProps.SetDocumentProperty objWkBk, PROP_..., Opt_...

        OptionsChanged = False
        
    End If
End Sub
