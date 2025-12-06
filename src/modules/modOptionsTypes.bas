Attribute VB_Name = "modOptionsTypes"
' -----------------------------------------------------------------------------------
' Module    : modOptionsTypes
' Purpose   : Defines DTOs and data types used by modOptions for configuration and
'             persistence of general and workbook-specific settings.
'
' Notes     :
'   - Keep this module free of behavior; types only.
'   - Extend tOptions as needed for additional settings.
'
' Sections  :
'   - tOptions          : Main structure for all options
'   - tGeneralOptions   : General Add-In Settings (Registry)
'   - tWorkbookOptions  : Workbook-specific settings (Document Properties)
' -----------------------------------------------------------------------------------
Option Explicit

' ===== Types ========================================================================

' -----------------------------------------------------------------------------------
' Type      : tGeneralOptions
' Purpose   : General add-in settings, stored in the Windows registry.
'             These settings apply to all workbooks.
' -------------------------------------------------------- ---------------------------
Public Type tGeneralOptions
    ' --- Paths ---
    manualPath              As String   ' Path to the manual directory
    
    ' --- Logging ---
    logRetentionDays        As Long     ' Retention period for log files in days (0 = unlimited)
End Type

' -----------------------------------------------------------------------------------
' Type      : tWorkbookOptions
' Purpose   : Workbook-specific settings, saved as Custom
'             Document Properties. These settings apply to each RDD workbook.
' -----------------------------------------------------------------------------------
Public Type tWorkbookOptions
    ' --- Room Sheet Defaults ---
    defaultGameWidth        As Long     ' Default game width (e.g., 320, 640, 1920)
    defaultGameHeight       As Long     ' Default game height (e.g., 200, 480, 1080)
    defaultBGWidth          As Long     ' Default background width
    defaultBGHeight         As Long     ' Default background height
    defaultUIHeight         As Long     ' Default UI height (e.g., 40, 0)
    defaultPerspective      As String   ' Default perspective (“Top-Down,” “Side-Scroller,” etc.)
    defaultParallax         As String   ' Default parallax (“None”, “Horizontal”, etc.)
    defaultSceneMode        As String   ' Default scene mode
    
    ' --- Behavior ---
    autoSyncLists           As Boolean  ' Automatic list synchronization when changes are made
    showValidationWarnings  As Boolean  ' Show validation warnings
    protectRoomSheets       As Boolean  ' Protect room sheets after creation
End Type

' -----------------------------------------------------------------------------------
' Type: tOptions
' Purpose: Main structure that summarizes all options.
' Used by modOptions for load/save operations.
' -------------------------------------------------------- ---------------------------
Public Type tOptions
    General     As tGeneralOptions      ' General settings (registry)
    Workbook    As tWorkbookOptions     ' Workbook settings (document properties)
End Type
