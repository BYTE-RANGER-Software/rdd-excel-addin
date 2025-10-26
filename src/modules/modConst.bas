Attribute VB_Name = "modConst"
Option Explicit
Option Private Module

Public Const FILENAME_MANUAL As String = "RDD-AddIn Manual.pdf"

Public Const WILDCARD_APP_PATH As String = "%AppPath%"
Public Const WILDCARD_MY_DOCUMENTS As String = "%MyDocuments%"


Public Const SHEET_ROOM_TEMPLATE As String = "Room_Template"
Public Const SHEET_DISPATCHER As String = "Dispatcher" ' contains data for the drop-down lists and macro code for automation

Public Const NAME_CELL_ROOM_ID As String = "Cell_RoomID"
Public Const NAME_CELL_SCENE_ID As String = "Cell_SceneID"
Public Const NAME_CELL_PARALLAX As String = "Cell_Parallax"
Public Const NAME_RANGE_PICKUPABLE_OBJ As String = "Range_PickupableObjects"
Public Const NAME_RANGE_MULTISTATE_OBJ As String = "Range_MultipleStateObjects"
Public Const NAME_RANGE_TOUCHABLE_OBJ As String = "Range_TouchableObjects"

Public Const NAME_LIST_ROOM_IDS As String = "List_RoomIDs"
Public Const NAME_LIST_SCENE_IDS As String = "List_SceneIDs"
Public Const NAME_LIST_OBJECTS As String = "List_Objects"
Public Const NAME_LIST_ACTORS As String = "List_Actors"
Public Const NAME_LIST_PUZZLE_TYPES As String = "List_Puzzle_Types"
Public Const NAME_RANGE_ADD_PARALLAX As String = "Range_AddParallaxSetWithHeader"
Public Const NAME_RANGE_DOORS_TO As String = "Range_DoorsTo"
Public Const NAME_RANGE_PUZZLES_PUZZLE_ID As String = "Range_Puzzles_PuzzleID"
Public Const NAME_RANGE_PUZZLES_TYPE As String = "Range_Puzzles_Type"

Public Const ROOM_SHEET_PREFIX As String = "Room"
Public Const ROOM_SHEET_ID_TAG_NAME As String = "RoomID"
Public Const ROOM_SHEET_ID_TAG_VAL_PRE As String = "R"

Public Const ROOM_HDR_NM_ADD_PARALLAX_SET As String = "ADDITIONAL PARALLAX SETTINGS"
Public Const ROOM_ADD_PARALLAX_SET_HIDE_TOKEN As String = "None"
Public Const ROOM_ADD_PARALLAX_SET_COVER_NAME As String = "COVER_ADD_PARALLAX_SET"

Public Const LISTS_COL_ROOM_ID As Long = 3
Public Const LISTS_HEADER_ROOM_ID As String = "Room IDs"
Public Const LISTS_COL_SCENE_ID As Long = 4
Public Const LISTS_HEADER_SCENE_ID As String = "Scene IDs"
Public Const LISTS_COL_OBJECTS As Long = 8
Public Const LISTS_HEADER_OBJECTS As String = "Objects"

Public Const BTN_INSERT_ROOM_PICTURE As String = "btnInsertImage"
Public Const MACRO_BTN_INSERT_PICTURE As String = "Dispatcher.InsertPicture_Button"

Public Const APP_DOC_TAG_KEY As String = "RDD_ADDIN_DOC"    ' marker for workbooks of the RDD add-in
Public Const APP_DOC_TAG_VAL As String = "v1.0"




