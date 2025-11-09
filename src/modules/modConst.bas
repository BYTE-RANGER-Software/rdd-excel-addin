Attribute VB_Name = "modConst"
Option Explicit
Option Private Module

Public Const FILENAME_MANUAL As String = "RDD-AddIn Manual.pdf"

Public Const WILDCARD_APP_PATH As String = "%AppPath%"
Public Const WILDCARD_MY_DOCUMENTS As String = "%MyDocuments%"

' Sheets
Public Const SHEET_ROOM_TEMPLATE As String = "Room_Template"
Public Const SHEET_DISPATCHER As String = "Dispatcher" ' contains data for the drop-down lists and macro code for automation

' Named Cells
Public Const NAME_CELL_BGHEIGHT As String = "Cell_BGHeight"
Public Const NAME_CELL_BGWIDTH As String = "Cell_BGWidth"
Public Const NAME_CELL_PARALLAX As String = "Cell_Parallax"
Public Const NAME_CELL_PERSPECTIVE As String = "Cell_Perspective"
Public Const NAME_CELL_ROOM_ALIAS As String = "Cell_RoomAlias"
Public Const NAME_CELL_ROOM_ID As String = "Cell_RoomID"
Public Const NAME_CELL_ROOM_NO As String = "Cell_RoomNo"
Public Const NAME_CELL_SCENE_ID As String = "Cell_SceneID"
Public Const NAME_CELL_SCENE_MODE As String = "Cell_SceneMode"
Public Const NAME_CELL_UIHEIGHT As String = "Cell_UIHeight"
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
Public Const NAME_RANGE_TOUCHABLE_OBJECTS_LABEL As String = "Range_TouchableObjects_Label"
Public Const NAME_RANGE_TOUCHABLE_OBJECTS_NOTES As String = "Range_TouchableObjects_Notes"
Public Const NAME_RANGE_WHAT_HAPPENS_HERE As String = "Range_WhatHappensHere"

' Data Table Named Columns
' Deprecated, do not use in further programming => TODO: replace by dynamic ListObject methods
Public Const NAME_LIST_ROOM_IDS As String = "List_RoomIDs"
Public Const NAME_LIST_SCENE_IDS As String = "List_SceneIDs"
Public Const NAME_LIST_OBJECTS As String = "List_Objects"
Public Const NAME_LIST_ACTORS As String = "List_Actors"
Public Const NAME_LIST_PUZZLE_TYPES As String = "List_Puzzle_Types"

' Data Tables (dynamic ListObject)
Public Const NAME_DATA_TABLE As String = "DropDownLists"

' Room Sheet related Constants
Public Const ROOM_SHEET_PREFIX As String = "Room"
Public Const ROOM_SHEET_ID_TAG_NAME As String = "RoomID"
Public Const ROOM_SHEET_ID_TAG_VAL_PRE As String = "R"

Public Const ROOM_HDR_NM_ADD_PARALLAX_SET As String = "ADDITIONAL PARALLAX SETTINGS"
Public Const ROOM_ADD_PARALLAX_SET_HIDE_TOKEN As String = "None"
Public Const ROOM_ADD_PARALLAX_SET_COVER_NAME As String = "COVER_ADD_PARALLAX_SET"

'Data Table Header Names
Public Const LISTS_HEADER_ROOM_ID As String = "Room IDs"
Public Const LISTS_HEADER_ROOM_ALIAS As String = "Room Alias"
Public Const LISTS_HEADER_SCENE_ID As String = "Scene IDs"
Public Const LISTS_HEADER_OBJECTS As String = "Objects"

'Data Table Column No
' Deprecated ! Don't use in further coding => TODO: replace by dynamic ListObject methods
Public Const LISTS_COL_ROOM_ID As Long = 3
Public Const LISTS_COL_ROOM_ALIAS As Long = 4
Public Const LISTS_COL_SCENE_ID As Long = 5
Public Const LISTS_COL_OBJECTS As Long = 9

' Dispatcher related constants
Public Const BTN_INSERT_ROOM_PICTURE As String = "btnInsertImage"
Public Const MACRO_BTN_INSERT_PICTURE As String = "Dispatcher.InsertPicture_Button"

' App constants
Public Const APP_DOC_TAG_KEY As String = "RDD_ADDIN_DOC" ' marker for workbooks of the RDD add-in
Public Const APP_DOC_TAG_VAL As String = "v1.0"




