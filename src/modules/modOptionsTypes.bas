Attribute VB_Name = "modOptionsTypes"
' -----------------------------------------------------------------------------------
' Module    : modOptionsTypes
' Purpose   : Defines DTOs and data types used by modOptions for configuration and
'             persistence of general and workbook-specific settings.
'
' Notes     :
'   - Keep this module free of behavior; types only.
'   - Extend tOptions as needed for additional settings.
' -----------------------------------------------------------------------------------
Option Explicit

' ===== Types and State ==============================================================
Public Type tOptions
    ' General part
    manualPath As String
    ' Workbook part (extend as needed)
End Type
