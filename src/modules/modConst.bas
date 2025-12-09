Attribute VB_Name = "modConst"
' ====================================================================================
' Module    : modConst
' Purpose   : Centralized constants for the RDD Add-In application.
'             Contains sheet names, named range identifiers, cell references,
'             file paths, and other application-wide constants.
'
' Public API:
'   - All constants are Public and accessed directly by name
'
' Categories:
'   - File paths and names (FILENAME_MANUAL, WILDCARD_*)
'   - Sheet names (SHEET_*)
'   - Named cell identifiers (NAME_CELL_*)
'   - Named range identifiers (NAME_RANGE_*)
'   - UI text constants
'
' Dependencies:
'   - None (pure constants module)
'
' Notes     :
'   - Use this module to avoid magic strings throughout the codebase
'   - All constants use meaningful prefixes (NAME_CELL_, NAME_RANGE_, etc.)
'   - Keep alphabetically sorted within each category for maintainability
'
' ====================================================================================
Option Explicit
Option Private Module

' Ribbon
Public Const RIBBON_BTN_BUILD_DATA As String = "RB75dd2c44_btnBuildData"
Public Const RIBBON_BTN_BUILD_CHART As String = "RB75dd2c44_btnBuildChart"
Public Const RIBBON_BTN_UPDATE_CHART As String = "RB75dd2c44_btnUpdateChart"
Public Const RIBBON_BTN_SYNC_LISTS As String = "RB75dd2c44_btnSyncLists"
Public Const RIBBON_BTN_NEED_SYNC_LISTS As String = "RB75dd2c44_btnNeedSyncLists"

Public Const RIBBON_CTX_MNU_BTN_1 As String = "RB75dd2c44_btnDynCtxMnu1"
Public Const RIBBON_CTX_MNU_BTN_2 As String = "RB75dd2c44_btnDynCtxMnu2"
    
Public Const FILENAME_MANUAL As String = "RDD-AddIn Manual.pdf"

Public Const WILDCARD_APP_PATH As String = "%AppPath%"
Public Const WILDCARD_MY_DOCUMENTS As String = "%MyDocuments%"

' Sheets
Public Const SHEET_ROOM_TEMPLATE As String = "Room_Template"
Public Const SHEET_DISPATCHER As String = "Dispatcher" ' contains data for the drop-down lists and macro code for automation

' Named Cells
Public Const NAME_CELL_GAME_HEIGHT As String = "Cell_GameHeight"
Public Const NAME_CELL_GAME_WIDTH As String = "Cell_GameWidth"
Public Const NAME_CELL_BG_HEIGHT As String = "Cell_BGHeight"
Public Const NAME_CELL_BG_WIDTH As String = "Cell_BGWidth"
Public Const NAME_CELL_PARALLAX As String = "Cell_Parallax"
Public Const NAME_CELL_PERSPECTIVE As String = "Cell_Perspective"
Public Const NAME_CELL_ROOM_ALIAS As String = "Cell_RoomAlias"
Public Const NAME_CELL_ROOM_ID As String = "Cell_RoomID"
Public Const NAME_CELL_ROOM_NO As String = "Cell_RoomNo"
Public Const NAME_CELL_SCENE_ID As String = "Cell_SceneID"
Public Const NAME_CELL_SCENE_MODE As String = "Cell_SceneMode"
Public Const NAME_CELL_UI_HEIGHT As String = "Cell_UIHeight"
Public Const NAME_CELL_VIEWPORT_H As String = "Cell_ViewportH"
Public Const NAME_CELL_VIEWPORT_W As String = "Cell_ViewportW"

' Named Ranges
Public Const NAME_RANGE_ACTORS_ACTOR_ID As String = "Range_Actors_ActorID"
Public Const NAME_RANGE_ACTORS_ACTOR_NAME As String = "Range_Actors_ActorName"
Public Const NAME_RANGE_ACTORS_CONDITION As String = "Range_Actors_Condition"
Public Const NAME_RANGE_ACTORS_NOTES As String = "Range_Actors_Notes"
Public Const NAME_RANGE_ADD_PARALLAX_SET_WITH_HEADER As String = "Range_AddParallaxSetWithHeader"
Public Const NAME_RANGE_CHK_LIST_STATUS As String = "Range_ChkList_Status"
Public Const NAME_RANGE_DOORS_TO_NOTES As String = "Range_DoorsTo_Notes"
Public Const NAME_RANGE_DOORS_TO_ROOM_ALIAS As String = "Range_DoorsTo_RoomAlias"
Public Const NAME_RANGE_DOORS_TO_ROOM_ID As String = "Range_DoorsTo_RoomID"
Public Const NAME_RANGE_FLAGS_BOOL_TYPE As String = "Range_Flags_BoolType"
Public Const NAME_RANGE_FLAGS_DESCRIPTION As String = "Range_Flags_Description"
Public Const NAME_RANGE_FLAGS_FLAG_ID As String = "Range_Flags_FlagID"
Public Const NAME_RANGE_FLAGS_INITIAL_STATE As String = "Range_Flags_InitialState"
Public Const NAME_RANGE_FLAGS_SCOPE As String = "Range_Flags_Scope"
Public Const NAME_RANGE_IMAGE_AREA As String = "Range_ImageArea"
Public Const NAME_RANGE_MULTI_STATE_OBJECTS_NOTES As String = "Range_MultiStateObjects_Notes"
Public Const NAME_RANGE_MULTI_STATE_OBJECTS_OBJECT_NAME As String = "Range_MultiStateObjects_ObjectName"
Public Const NAME_RANGE_MULTI_STATE_OBJECTS_STATE As String = "Range_MultiStateObjects_State"
Public Const NAME_RANGE_MULTI_STATE_OBJECTS_STATE_ID As String = "Range_MultiStateObjects_StateID"
Public Const NAME_RANGE_PARALLAX_ASSET_KEY As String = "Range_ParallaxAssetKey"
Public Const NAME_RANGE_PARALLAX_LAYER As String = "Range_ParallaxLayer"
Public Const NAME_RANGE_PARALLAX_NOTES As String = "Range_ParallaxNotes"
Public Const NAME_RANGE_PICKUPABLE_OBJECTS_ITEM_ID As String = "Range_PickupableObjects_ItemID"
Public Const NAME_RANGE_PICKUPABLE_OBJECTS_NAME As String = "Range_PickupableObjects_Name"
Public Const NAME_RANGE_PICKUPABLE_OBJECTS_NOTES As String = "Range_PickupableObjects_Notes"
Public Const NAME_RANGE_PUZZLES_ACTION As String = "Range_Puzzles_Action"
Public Const NAME_RANGE_PUZZLES_DEPENDS_ON As String = "Range_Puzzles_DependsOn"
Public Const NAME_RANGE_PUZZLES_DIFFICULTY As String = "Range_Puzzles_Difficulty"
Public Const NAME_RANGE_PUZZLES_GRANTS As String = "Range_Puzzles_Grants"
Public Const NAME_RANGE_PUZZLES_IQPOINTS As String = "Range_Puzzles_IQPoints"
Public Const NAME_RANGE_PUZZLES_NOTES As String = "Range_Puzzles_Notes"
Public Const NAME_RANGE_PUZZLES_OWNER As String = "Range_Puzzles_Owner"
Public Const NAME_RANGE_PUZZLES_PUZZLE_ID As String = "Range_Puzzles_PuzzleID"
Public Const NAME_RANGE_PUZZLES_REQUIRES As String = "Range_Puzzles_Requires"
Public Const NAME_RANGE_PUZZLES_STATUS As String = "Range_Puzzles_Status"
Public Const NAME_RANGE_PUZZLES_TARGET As String = "Range_Puzzles_Target"
Public Const NAME_RANGE_PUZZLES_TITLE As String = "Range_Puzzles_Title"
Public Const NAME_RANGE_ROOM_DESCRIPTION As String = "Range_RoomDescription"
Public Const NAME_RANGE_SOUNDS_DESCRIPTION As String = "Range_Sounds_Description"
Public Const NAME_RANGE_SOUNDS_NOTES As String = "Range_Sounds_Notes"
Public Const NAME_RANGE_SOUNDS_SOUND_ID As String = "Range_Sounds_SoundID"
Public Const NAME_RANGE_SOUNDS_TRIGGER As String = "Range_Sounds_Trigger"
Public Const NAME_RANGE_SOUNDS_TYPE As String = "Range_Sounds_Type"
Public Const NAME_RANGE_SPECIAL_FX_ANIMATION_ID As String = "Range_SpecialFX_AnimationID"
Public Const NAME_RANGE_SPECIAL_FX_DESCRIPTION As String = "Range_SpecialFX_Description"
Public Const NAME_RANGE_SPECIAL_FX_NOTES As String = "Range_SpecialFX_Notes"
Public Const NAME_RANGE_SPECIAL_FX_TRIGGER As String = "Range_SpecialFX_Trigger"
Public Const NAME_RANGE_SPECIAL_FX_TYPE As String = "Range_SpecialFX_Type"
Public Const NAME_RANGE_TOUCHABLE_OBJECTS_HOTSPOT_ID As String = "Range_TouchableObjects_HotspotID"
Public Const NAME_RANGE_TOUCHABLE_OBJECTS_HOTSPOT_NAME As String = "Range_TouchableObjects_Name"
Public Const NAME_RANGE_TOUCHABLE_OBJECTS_NOTES As String = "Range_TouchableObjects_Notes"
Public Const NAME_RANGE_WHAT_HAPPENS_HERE As String = "Range_WhatHappensHere"

' Data Tables (dynamic ListObject)
Public Const NAME_DATA_TABLE As String = "DropDownLists"

' Room Sheet related Constants
Public Const ROOM_SHEET_DEFAULT_PREFIX As String = "Room"
Public Const ROOM_SHEET_ALIAS_PREFIX As String = "r_"
Public Const ROOM_SHEET_ID_PREFIX As String = "R"
Public Const ROOM_SHEET_ID_TAG_NAME As String = "RoomID"

Public Const ROOM_HDR_NM_ADD_PARALLAX_SET As String = "ADDITIONAL PARALLAX SETTINGS"
Public Const ROOM_ADD_PARALLAX_SET_HIDE_TOKEN As String = "None"
Public Const ROOM_ADD_PARALLAX_SET_COVER_NAME As String = "COVER_ADD_PARALLAX_SET"

'Data Table Header Names
Public Const LISTS_HEADER_STATUS As String = "Status"
Public Const LISTS_HEADER_PERSPECTIVE As String = "Perspective"
Public Const LISTS_HEADER_SCENE_MODE As String = "Scene Mode"
Public Const LISTS_HEADER_PARALLAX As String = "Parallax"
Public Const LISTS_HEADER_LAYERS As String = "Layers"
Public Const LISTS_HEADER_WIDTH As String = "Width"
Public Const LISTS_HEADER_HEIGHT As String = "Height"
Public Const LISTS_HEADER_UI_HEIGHT As String = "UI Height"
Public Const LISTS_HEADER_ROOM_ALIAS As String = "Room Alias"
Public Const LISTS_HEADER_ROOM_ID As String = "Room ID"
Public Const LISTS_HEADER_ROOM_NO As String = "Room No"
Public Const LISTS_HEADER_SCENE_ID As String = "Scene ID"
Public Const LISTS_HEADER_ACTOR_ID As String = "Actor ID"
Public Const LISTS_HEADER_ACTOR_NAME As String = "Actor Name"
Public Const LISTS_HEADER_SOUND_ID As String = "Sound ID"
Public Const LISTS_HEADER_SOUND_NAME As String = "Sound Name"
Public Const LISTS_HEADER_SOUND_TYPE As String = "Sound Type"
Public Const LISTS_HEADER_ANIMATION_ID As String = "Animation ID"
Public Const LISTS_HEADER_ANIMATION_NAME As String = "Animation Name"
Public Const LISTS_HEADER_ANIMATION_TYPE As String = "Animation Type"
Public Const LISTS_HEADER_ITEM_ID As String = "Item ID"
Public Const LISTS_HEADER_ITEM_NAME As String = "Item Name"
Public Const LISTS_HEADER_STATE_OBJECT_NAME As String = "State Object Name"
Public Const LISTS_HEADER_STATE_OBJECT_STATE As String = "State Object State"
Public Const LISTS_HEADER_STATE_OBJECT_ID As String = "State Object ID"
Public Const LISTS_HEADER_HOTSPOT_ID As String = "Hotspot ID"
Public Const LISTS_HEADER_HOTSPOT_NAME As String = "Hotspot Name"
Public Const LISTS_HEADER_FLAG_ID As String = "Flag ID"
Public Const LISTS_HEADER_FLAG_SCOPE As String = "Flag Scope"
Public Const LISTS_HEADER_FLAG_TYPE As String = "Flag Type"
Public Const LISTS_HEADER_FLAG_NAME As String = "Flag Name"
Public Const LISTS_HEADER_PUZZLE_ID As String = "Puzzle ID"
Public Const LISTS_HEADER_PUZZLE_ACTION As String = "Puzzle Action"
Public Const LISTS_HEADER_PUZZLE_DIFFICULTY As String = "Puzzle Difficulty"
Public Const LISTS_HEADER_PUZZLE_OWNER As String = "Puzzle Owner"
Public Const LISTS_HEADER_PUZZLE_STATUS As String = "Puzzle Status"
Public Const LISTS_HEADER_PUZZLE_POINTS As String = "Puzzle Points"

' Dispatcher related constants
Public Const BTN_INSERT_ROOM_PICTURE As String = "btnInsertImage"
Public Const MACRO_BTN_INSERT_PICTURE As String = "Dispatcher.InsertPicture_Button"

' App constants
Public Const APP_DOC_TAG_KEY As String = "RDD_ADDIN_DOC" ' marker for workbooks of the RDD add-in
Public Const APP_DOC_TAG_VAL As String = "v1.0"

' Error base Number
Public Const ERR_BASE As Long = 60000
Public Const ERR_MISSING_DISPATCHER         As Long = ERR_BASE + 1
Public Const ERR_MISSING_DATA_TABLE         As Long = ERR_BASE + 2

' errors used in clsFormDrop
Public Const ERR_FORM_DROP_BASE             As Long = ERR_BASE + 100

Public Const ERR_NOT_INITIALIZED            As Long = ERR_FORM_DROP_BASE + 1      ' "Call .Init first."
Public Const ERR_USE_INSTANCE               As Long = ERR_FORM_DROP_BASE + 2      ' "Use an instance to call this member."
Public Const ERR_LISTS_NOT_SET              As Long = ERR_FORM_DROP_BASE + 3      ' "No lists have been initialized..."
Public Const ERR_NO_ARRAY                   As Long = ERR_FORM_DROP_BASE + 4      ' "must be an array.."
Public Const ERR_OBJ_IS_NOTHING             As Long = ERR_FORM_DROP_BASE + 5      ' "is Nothing."
Public Const ERR_DIF_LENGTH                 As Long = ERR_FORM_DROP_BASE + 6      ' ".. must have the same length."
Public Const ERR_INSTANCE_EXISTS            As Long = ERR_FORM_DROP_BASE + 7

' errors used in modRooms
Public Const ERR_ROOMS_BASE                 As Long = ERR_BASE + 200

Public Const ERR_ARG_NULL_TARGETSHEET       As Long = ERR_ROOMS_BASE + 1    ' "Argument 'targetSheet' must not be Nothing."
Public Const ERR_NOT_A_ROOM_SHEET           As Long = ERR_ROOMS_BASE + 2    ' "The provided sheet is not a Room sheet."
Public Const ERR_ROOM_SHEET_REFERENCED      As Long = ERR_ROOMS_BASE + 3    ' "Room sheet cannot be deleted because it is referenced..."
Public Const ERR_INVALID_ROOM_NAME          As Long = ERR_ROOMS_BASE + 4

'meta data Constants for form dropDowns
Public Const FD_ANCHOR_NAME_PATTERN As String = "*!DD_Anchor_*"
Public Const FD_SLAVE_ANCHOR_NAME_PATTERN As String = FD_ANCHOR_NAME_PATTERN & ".[0-9]*."

Public Const FD_META_PREFIX        As String = "FD:"
Public Const FD_META_PAIR_SEP      As String = ";"          ' key=value;key=value
Public Const FD_META_LIST_SEP      As String = "|"          ' a|b|c

Public Const FD_META_KEY_CAT       As String = "cat"        ' worksheet name for category range
Public Const FD_META_KEY_SUBS      As String = "subs"       ' worksheet names for sub ranges (| separated)

Public Const FD_META_KEY_CAT_TBL   As String = "cattbl"     ' table name for category
Public Const FD_META_KEY_CAT_COL   As String = "catcol"     ' header name for category column
Public Const FD_META_KEY_SUBS_TBL  As String = "substbl"    ' table name for subs
Public Const FD_META_KEY_SUBS_COLS As String = "subscols"   ' header names for sub columns (| separated)
Public Const FD_META_KEY_CAT_SHEET As String = "catsheet"   ' optional: sheet name where cattbl is located
Public Const FD_META_KEY_SUBS_SHEET As String = "subsheet"  ' optional: sheet name where substbl is located

