# RDD-AddIn Manual

**Room Design Document Add-In for Excel**  
*Version 0.10 – For Adventure Game Studio (AGS) Game Development*

---

## Table of Contents

1. [Introduction](#1-introduction)
2. [Installation](#2-installation)
3. [User Interface](#3-user-interface)
   - 3.1 [The Ribbon Menu](#31-the-ribbon-menu)
   - 3.2 [Context Menus](#32-context-menus)
4. [Room Templates](#4-room-templates)
   - 4.1 [Room Sheet Structure](#41-room-sheet-structure)
   - 4.2 [Creating and Editing Rooms](#42-creating-and-editing-rooms)
5. [Puzzle Dependency Chart (PDC)](#5-puzzle-dependency-chart-pdc)
   - 5.1 [Concept and Methodology](#51-concept-and-methodology)
   - 5.2 [PDC Workflow](#52-pdc-workflow)
   - 5.3 [Chart Navigation](#53-chart-navigation)
6. [List Synchronization](#6-list-synchronization)
7. [Search and Navigation](#7-search-and-navigation)
8. [Export Functions](#8-export-functions)
9. [Options and Settings](#9-options-and-settings)
10. [Tips and Best Practices](#10-tips-and-best-practices)

---

## 1. Introduction

The **RDD-AddIn** (Room Design Document Add-In) is a comprehensive Excel extension for developing Adventure Games with Adventure Game Studio (AGS). It provides a structured method for documenting rooms, puzzles, items, actors, and other game elements.

The add-in is based on proven game design methods, particularly the **Puzzle Dependency Chart** methodology by Ron Gilbert, which was used in the development of classics like Monkey Island.

### Main Features

| Feature | Description |
|---------|-------------|
| Room Management | Create, edit, and manage Room Design Documents |
| Puzzle Dependency Chart | Visualization of puzzle dependencies based on Ron Gilbert's method |
| Dropdown Synchronization | Automatic list management from room data |
| Context Menus | Quick access to frequently used functions |
| Find Usage | Search for usages of Items, Actors, Flags |
| Export | PDF and CSV export of documentation |

---

## 2. Installation

### System Requirements

- Microsoft Excel 2010 or higher (Windows)
- Macros must be enabled
- Scripting Runtime Library (scrrun.dll) – included by default

### Installation Steps

**Step 1:** Copy the file `RDD.xlam` to the Excel Add-Ins folder:

```shell
%APPDATA%\Microsoft\AddIns\
```

**Step 2:** Open Excel and navigate to:  
*File → Options → Add-Ins → Manage Excel Add-Ins → Go...*

**Step 3:** Check the box next to "RDD" and click OK.

**Step 4:** The new tab "RDD" now appears in the Ribbon menu.

> 💡 **Info:** When you start the program for the first time, a working folder is created under `%AppData%\BYTE RANGER\RDDAddIn`, which contains log files and the manual, as well as a temporary folder under `%Temp%\BYTE RANGER\RDDAddIn`.

---

## 3. User Interface

### 3.1 The Ribbon Menu

![Ribbon](images/Ribbon.png)
After installation, a new tab **RDD** appears in the Excel Ribbon with the following groups:

#### Group: Rooms

| Button | Function |
|--------|----------|
| **Add Room** | Creates a new Room Sheet based on the template |
| **Edit** | Opens dialog to edit Room ID, Scene ID, Alias |
| **Delete** | Deletes the current Room Sheet (with reference check) |
| **Sync Lists** | Synchronizes all dropdown lists from room data |
| **Validate** | Checks data for duplicates, missing references, cycles |

#### Group: Dependency Chart

| Button | Function |
|--------|----------|
| **Build Data** | Extracts puzzle data and creates PDCData sheet |
| **Generate Chart** | Creates visual Puzzle Dependency Chart |
| **Update Chart** | Updates existing chart with new data |

#### Group: Export

| Button | Function |
|--------|----------|
| **PDF Export** | Exports Room Sheets as printable PDF |
| **CSV Export** | Exports PDC data as CSV (nodes.csv, edges.csv) |

#### Group: Info

| Button | Function |
|--------|----------|
| **Options** | Opens settings dialog |
| **Log** | Displays log files |
| **Manual** | Opens this manual |
| **Version** | Shows About dialog with version information |

### 3.2 Context Menus

The add-in extends the Excel context menu (right-click) with context-sensitive options. Different menu options are displayed depending on the active cell position:

| Cell Type | Menu Option 1 | Menu Option 2 |
|-----------|---------------|---------------|
| **Room ID/Alias** | Add New Room | Goto Room |
| **Puzzle ID** | Goto Node in Chart | Show Dependencies |
| **Item ID** | Find Usage | – |
| **Actor ID** | Find Usage | – |
| **Hotspot ID** | Find Usage | – |
| **Flag ID** | Find Usage | – |
| **Dependencies** | Goto Referenced | – |

---

## 4. Room Templates

### 4.1 Room Sheet Structure

Each Room Sheet follows a standardized structure with multiple sections:

| Section | Rows | Content |
|---------|------|---------|
| **ROOM HEADER** | 1 | Room ID, Scene ID, Room No, Room Alias |
| **CHECKLIST** | 3-12 | Status tracking for assets (Backgrounds, Events, Speech, etc.) |
| **PICTURE AREA** | 3-12 | Space for screenshot or concept art |
| **SCENE DESCRIPTION** | 15-23 | Narrative description of the scene |
| **WHAT HAPPENS HERE?** | 24-38 | Story events and gameplay occurrences |
| **GENERAL SETTINGS** | 24-38 | Perspective, Parallax, Dimensions, Viewport |
| **DOORS TO...** | 40-53 | Connections to other rooms |
| **ACTORS** | 40-53 | Characters with conditions |
| **SOUNDS** | 55-68 | Sound effects and music |
| **SPECIAL FX** | 55-68 | Animations and effects |
| **PICKUPABLE OBJECTS** | 70-83 | Items to collect |
| **MULTI-STATE OBJECTS** | 70-83 | Objects with multiple states |
| **TOUCHABLE OBJECTS** | 85-98 | Hotspots and interactive areas |
| **FLAGS / KNOWLEDGE** | 85-98 | Variables and knowledge flags |
| **PUZZLES** | 100-115 | Complete puzzle documentation |

#### PUZZLES Columns

| Column | Description |
|--------|-------------|
| Puzzle ID | Unique ID (e.g., P001, P002) |
| Title | Brief description of the puzzle |
| Target | Target object of the action |
| Action/Verb | Use, Talk, Give, Look, etc. |
| DependsOn | Required puzzles (comma-separated) |
| Requires | Required Items/Flags |
| Grants | Items/Flags granted after solving |
| Difficulty | Difficulty level |
| Owner | Responsible designer |
| Status | todo, in progress, done, n/a |
| Points | IQ points |
| Notes | Additional notes |

### 4.2 Creating and Editing Rooms

#### Creating a New Room

1. Click **"Add Room"** in the Ribbon
2. Optional: Enter the Scene ID in the dialog (e.g., "Hindu Temple")
3. Enter a Room Alias (e.g., "Entrance")
4. Enter a AGS Room Number (e.g., "1")
5. A new sheet is created based on the template

> ⚠️ **Note:** Room IDs must be unique and are automatically generated according to the following pattern: “R###.”  
The alias is automatically prefixed with “r_”..

#### Editing Room Identity

1. Navigate to the desired Room Sheet
2. Click **"Edit"** in the Ribbon
3. Modify Room ID, Scene ID, or Alias
4. All references are automatically updated

#### Deleting a Room

1. Navigate to the Room Sheet to be deleted
2. Click **"Delete"** in the Ribbon
3. Confirm the deletion
4. The system checks for references in other rooms first

---

## 5. Puzzle Dependency Chart (PDC)

### 5.1 Concept and Methodology

The **Puzzle Dependency Chart** (PDC) is a visual method for representing puzzle dependencies in adventure games. This technique was developed by Ron Gilbert and used in classic LucasArts adventures like "The Secret of Monkey Island."

#### Node Types

| Node Type | ID Prefix | Description | Color |
|-----------|-----------|-------------|-------|
| Puzzle | P001, P002... | A solvable puzzle/task | Blue |
| Item | i_key, i_map... | An inventory item | Green |
| Flag (Global) | g_doorOpen... | Global knowledge variable | Purple |
| Flag (Room) | r_visited... | Room-specific variable | Orange |

#### Edge Types (Connections)

| Edge Type | Puzzle Column | Meaning |
|-----------|---------------|---------|
| depends | DependsOn | Puzzle X must be solved before Puzzle Y |
| requires | Requires | Puzzle requires Item/Flag to solve |
| grants | Grants | Puzzle grants Item/Flag after solving |

### 5.2 PDC Workflow

```txt
┌─────────────┐    ┌─────────────┐    ┌─────────────┐    ┌─────────────┐    ┌─────────────┐
│ Fill Room   │───>│   Run       │───>│ Build PDC   │───>│  Generate   │───>│ Navigate &  │
│ Sheets with │    │ Validation  │    │    Data     │    │   Chart     │    │  Analyze    │
│   Puzzles   │    │             │    │             │    │             │    │             │
└─────────────┘    └─────────────┘    └─────────────┘    └─────────────┘    └─────────────┘
```

#### Step 1: Document Puzzles

- Enter Puzzle ID
- Define DependsOn
- Set Requires
- Assign Grants

#### Step 2: Validation

- Check duplicates
- Validate IDs
- Check references
- Detect cycles

#### Step 3: Build Data

- Extract nodes
- Create edges
- Assign types
- "PDCData" sheet is created

#### Step 4: Generate Chart

- Create shapes
- Draw connectors
- Apply layout
- "Chart" sheet is created

#### Step 5: Navigation

- Ctrl+Click on node → Jump to source
- Analyze dependencies

### 5.3 Chart Navigation

The generated chart is fully interactive:

**Ctrl+Click on a Node:**  
Jumps directly to the puzzle definition in the corresponding Room Sheet. This uses the Windows API (GetAsyncKeyState) to detect the Ctrl key during click.

**Context Menu on PDCData:**  
Right-clicking on a puzzle cell in the PDCData sheet shows additional options:

- "Goto Node in Chart" – Navigate to the node in the chart
- "Show Dependencies" – Display all dependencies

---

## 6. List Synchronization

The add-in automatically manages dropdown lists in the Dispatcher sheet. These lists are used for validation and auto-complete in Room Sheets.

### Managed Lists

| List | Source | Usage |
|------|--------|-------|
| Room ID | All Room Sheets | DOORS TO... navigation |
| Room Alias | All Room Sheets | DOORS TO... navigation |
| Scene ID | All Room Sheets | Referencing |
| Actor ID | ACTORS sections | Puzzle Owner/Target |
| Actor Name | ACTORS sections | Display |
| Sound ID | SOUNDS sections | Display |
| Sound Description | SOUNDS sections | Display |
| Item ID | PICKUPABLE OBJECTS | Puzzle Requires/Grants |
| Item Name | PICKUPABLE OBJECTS | Display |
| Flag ID | FLAGS sections | Puzzle Requires/Grants |
| Hotspot ID | TOUCHABLE OBJECTS | Puzzle Target |
| Puzzle ID | PUZZLES sections | DependsOn |

### Automatic Synchronization

With the "Auto Sync Lists" option enabled, lists are automatically updated when:

- A Room Sheet is modified
- A new room is created
- A room is deleted

### Manual Synchronization

The **"Synchronize Lists"** button in the Ribbon forces a complete synchronization.

The button shows two states:

- ![🟢 **Green:**](images/SyncGreen.png) Lists are synchronized
- ![🟠 **Orange:**](images/SyncOrange.png) Changes detected, sync recommended

---

## 7. Search and Navigation

The "Find Usage" function allows finding all usages of Items, Actors, Hotspots, and Flags across all Room Sheets.

### Using Find Usage

1. Position the cursor on an ID cell (e.g., Item ID)
2. Right-click → Select "Find Usage"
3. The search results window opens
4. Double-click on a result to navigate to the corresponding cell

### Searched Areas

| Element | Searched Columns |
|---------|------------------|
| Items | Puzzles_Requires, Puzzles_Grants, PickupableObjects_ItemID |
| Actors | Actors_Condition, Puzzles_Owner, Puzzles_Target |
| Hotspots | TouchableObjects_HotspotID, Puzzles_Target |
| Flags | Flags_FlagID, Puzzles_Requires, Puzzles_Grants |

> 💡 **Tip:** The search supports comma-separated values in cells. If a cell contains `i_key, i_map`, searching for `i_key` will show this cell as a match.

---

## 8. Export Functions

### PDF Export

The PDF export creates a printable document with all Room Sheets:

- Each Room Sheet is exported as a separate page
- Formatting and images are preserved
- Optimized for A4 landscape format

**Access:** Ribbon → Export → "PDF Export"  
**Location:** Dialog for selecting the target folder

### CSV Export

The CSV export creates separate files for PDC data:

- `nodes.csv` – All puzzle nodes
- `edges.csv` – All dependencies

These files can be used in other tools (e.g., Graphviz, yEd) for further visualization.

---

## 9. Options and Settings

The Options window (Ribbon → Info → "Options") provides two areas:

### General Settings (Registry)

| Setting | Description | Default |
|---------|-------------|---------|
| Manual Path | Path to manual directory | `%AppData%\BYTE RANGER\RDDAddIn\` |
| Log Retention Days | Days until old logs are deleted | 30 |

### Workbook Settings (Document Properties)

| Setting | Description | Default |
|---------|-------------|---------|
| Default Game Width | Default game width in pixels | 320 |
| Default Game Height | Default game height in pixels | 200 |
| Default BG Width | Default background width | 320 |
| Default BG Height | Default background height | 200 |
| Default UI Height | Default UI height | 40 |
| Default Perspective | Default perspective | (empty) |
| Default Parallax | Default parallax mode | None |
| Default Scene Mode | Default scene mode | (empty) |
| Auto Sync Lists | Automatic list synchronization | True |
| Show Validation Warnings | Show validation warnings | True |

---

## 10. Tips and Best Practices

### Naming Conventions

| Element | Convention | Example | Handling |
|---------|------------|---------|----------|
| Room ID | R + three-digit number | R001, R002, R100 | ✅ Automatic |
| Room Alias | r_ + descriptive name | r_entrance, r_cellar | ✅ Automatic |
| Puzzle ID | P + three-digit number | P001, P002 | ✅ Automatic |
| Item ID | i_ + name | i_key, i_goldcoin | ✅ Automatic |
| Flag (Global) | g_ + name | g_doorUnlocked | ✅ Automatic |
| Flag (Room) | r_ + name | r_visited | ✅ Automatic |
| Actor ID | c + name (Character) | cEgo, cBartender | ✅ Automatic |
| Hotspot ID | h + name | hDoor, hWindow | ✅ Automatic |
| State Object | o + name | oDoor, oLever | ✅ Automatic |

> 👉 Note: “✅ Automatic” means that the add-in assigns the ID or name directly according to the convention when creating the item. With “⚠️ Manual,” the user must ensure that the correct spelling is used.

### Workflow Recommendations

- **Validate regularly:** Always run validation after major changes.
- **Keep lists synchronized:** With Auto-Sync disabled, synchronize manually regularly.
- **Create backups:** Make a copy of the workbook before major changes.
- **Build PDC iteratively:** Start with main puzzles and refine later.
- **Consistent IDs:** Use the same naming conventions throughout.

### Troubleshooting

| Problem | Solution |
|---------|----------|
| Ribbon doesn't appear | Re-enable add-in under Excel Options → Add-Ins |
| Buttons grayed out | Ensure an RDD workbook is open |
| Validation errors | Check the log file under Info → Log |
| Chart not updated | Run "Update Chart" after data changes |
| Lists not synchronized | Manually run "Sync Lists" |
| Context menu doesn't appear | Position cursor on a valid ID cell |

---

*RDD-AddIn – Room Design Document Add-In for Adventure Game Studio*  
*Documentation Version 0.10*
