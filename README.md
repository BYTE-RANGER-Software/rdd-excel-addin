![License](https://img.shields.io/badge/license-MIT-blue)
[![Project Maintenance](https://img.shields.io/maintenance/yes/2025.svg)](https://github.com/byte-ranger-software/rdd-excel-addin 'GitHub Repository')
![Status](https://img.shields.io/badge/status-beta-yellow)

# RDD Excel Add‑in (Room Design Document + Puzzle Dependency Chart)

> ⚠️ **This project is currently in beta state and not yet ready for production use.**  
> Bugs or incomplete functionality are expected.

**Purpose** – An Excel add‑in that helps you author Room Design Documents and a Puzzle Dependency Chart in one place, then keep them consistent.  
Designed for Adventure Games made with **Adventure Game Studio** (AGS), with a 9‑verb UI and data‑oriented workflow.

---

## Features

### Rooms

- **Add Room** – Create a new Room sheet from a template, apply IDs and validations
- **Edit Room** – Modify Room ID, Scene ID, or Alias with automatic reference updates
- **Remove Room** – Delete Room sheets with dependency checking

### Dropdown Lists

Two safe update modes:

- **Update Dropdown Lists** – Append missing items only, existing entries remain unchanged
- **Synchronize Dropdown Lists** – Rebuild from Rooms, remove items that no longer exist

### Puzzle Dependency Chart (PDC)

Based on Ron Gilbert's methodology from classic LucasArts adventures:

- **Validate** – Check duplicate IDs, missing references, and hints for cycles
- **Build Data** – Generate nodes and edges from all Rooms into the **PDCData** sheet
- **Generate Chart** – Create visual dependency chart with clickable nodes
- **Update Chart** – Refresh the chart from the **PDCData** sheet

### Search & Navigation

- **Find Usage** – Search for Items, Actors, Hotspots, Flags across all Rooms
- **Goto Node in Chart** – Navigate from puzzle data to chart visualization
- **Goto Referenced** – Jump to referenced items from Dependencies column

### Export

- **PDF Export** – Export Room sheets as printable documentation
- **CSV Export** – Export nodes and edges for use in external tools (Graphviz, yEd)

---

## Repository Layout

```shell
/
├─ src/
│  ├─ RDD_AddIn.xlsm      # Editable source workbook (ribbon, modules)
│  ├─ modules/            # VBA module exports (.bas/.cls/.frm)
│  ├─ ribbon/             # customUI14.xml, callbacks map
│  └─ ribbon_icons/       # PNG icons (16, 24, 32, 48, 64, 128)
├─ bin/
│  └─ RDD.xlam            # Add-in build (ready to install)
├─ docs/
│  ├─ manual.md           # User manual (EN)
│  ├─ manual_de.md        # User manual (DE)
│  └─ RDD-AddIn_Manual.pdf
└─ README.md
```

---

## Installation

1. Copy `RDD.xlam` to Excel Add-Ins folder:

   ```bat
   %APPDATA%\Microsoft\AddIns\
   ```

2. Open Excel → File → Options → Add-Ins → Manage Excel Add-Ins → Go...

3. Check "RDD_AddIn" and click OK

4. The **RDD-AddIn** tab appears in the Ribbon

---

## Architecture

The add-in follows a layered architecture for maintainability and extensibility:

```txt
┌─────────────────────────────────────────────────────────────────────────────┐
│                              UI Layer                                       │
│  ┌─────────────┐  ┌─────────────┐  ┌─────────────┐  ┌─────────────┐         │
│  │  modRibbon  │  │modCellCtxMnu│  │   frm*      │  │ frmOptions  │         │
│  │   Ribbon    │  │Context Menu │  │  UserForms  │  │Options Dlg  │         │
│  └─────────────┘  └─────────────┘  └─────────────┘  └─────────────┘         │
└─────────────────────────────────────────────────────────────────────────────┘
                                     ↕
┌─────────────────────────────────────────────────────────────────────────────┐
│                           Controller Layer                                  │
│  ┌─────────────┐  ┌─────────────┐  ┌─────────────┐  ┌─────────────┐         │
│  │  modMain    │  │clsAppEvents │  │  clsState   │  │ Dispatcher  │         │
│  │  Central    │  │Event Handler│  │State Manager│  │  FormDrop   │         │
│  └─────────────┘  └─────────────┘  └─────────────┘  └─────────────┘         │
└─────────────────────────────────────────────────────────────────────────────┘
                                     ↕
┌─────────────────────────────────────────────────────────────────────────────┐
│                         Business Logic Layer                                │
│  ┌─────────────┐  ┌─────────────┐  ┌─────────────┐  ┌─────────────┐         │
│  │  modRooms   │  │   modPDC    │  │  modSearch  │  │  modLists   │         │
│  │Room Operat. │  │  PDC Logic  │  │ Find Usage  │  │Dropdown Sync│         │
│  └─────────────┘  └─────────────┘  └─────────────┘  └─────────────┘         │
└─────────────────────────────────────────────────────────────────────────────┘
                                     ↕
┌─────────────────────────────────────────────────────────────────────────────┐
│                             Data Layer                                      │
│  ┌─────────────┐  ┌─────────────┐  ┌─────────────┐  ┌─────────────┐         │
│  │ modOptions  │  │  modProps   │  │ modRanges   │  │  modExport  │         │
│  │  Settings   │  │Doc Properti.│  │Named Ranges │  │  PDF/CSV    │         │
│  └─────────────┘  └─────────────┘  └─────────────┘  └─────────────┘         │
└─────────────────────────────────────────────────────────────────────────────┘
```

### Module Overview

| Module | Layer | Purpose |
|--------|-------|---------|
| **modMain** | Controller | Central controller, lifecycle management, feature orchestration |
| **modRibbon** | UI | Ribbon callbacks, UI routing |
| **modCellCtxMnu** | UI | Context menu logic and dynamic menu updates |
| **modRooms** | Business Logic | Room sheet operations, creation, deletion, validation |
| **modPDC** | Business Logic | Puzzle Dependency Chart generation and synchronization |
| **modSearch** | Business Logic | Find Usage implementation across all rooms |
| **modLists** | Business Logic | Dropdown list synchronization (append/rebuild modes) |
| **modOptions** | Data | Settings management (Registry + Document Properties) |
| **modProps** | Data | Document Properties read/write operations |
| **modRanges** | Data | Named Range operations and intersection checks |
| **modExport** | Data | PDF and CSV export functionality |
| **modConst** | Shared | Constants for named ranges, sheet names, error codes |
| **modUtil** | Shared | Utility functions |
| **modErr** | Shared | Error handling and logging |

### Class Modules

| Class | Purpose |
|-------|---------|
| **clsAppEvents** | Application-level event handler (WithEvents sink) |
| **clsState** | Global state manager (Singleton pattern) |
| **clsFormDrop** | Form-based dropdown control with cascading support |
| **clsFormDropManager** | Lifecycle management for FormDrop instances |
| **clsLog** | Logging implementation |
| **clsProgressBar** | Progress indicator for long operations |

### UserForms

| Form | Purpose |
|------|---------|
| **frmObjectEdit** | Room identity editor (ID, Scene, Alias) |
| **frmOptions** | Settings dialog (General + Workbook options) |
| **frmSearchResults** | Find Usage results display |
| **frmAbout** | About dialog with version info |
| **frmWait** | Wait indicator for long operations |

---

## Data Model

### Sheets

| Sheet | Purpose |
|-------|---------|
| **Room_Template** | Template for new Room sheets (hidden) |
| **Dispatcher** | Contains dropdown list data and FormDrop anchors |
| **PDCData** | Generated puzzle nodes and edges |
| **Chart** | Visual Puzzle Dependency Chart |

### PDC Node Types

| Type | ID Prefix | Description |
|------|-----------|-------------|
| puzzle | P001, P002... | Solvable puzzle/task |
| item | i_key, i_map... | Inventory item |
| flag (global) | g_doorOpen... | Global knowledge variable |
| flag (room) | r_visited... | Room-specific variable |

### PDC Edge Types

| Type | Source Column | Meaning |
|------|---------------|---------|
| depends | DependsOn | Puzzle X must be solved before Puzzle Y |
| requires | Requires | Puzzle requires Item/Flag to solve |
| grants | Grants | Puzzle grants Item/Flag after solving |

---

## Settings Storage

### Registry (General Settings)

```text
HKCU\Software\VB and VBA Program Settings\RDD-AddIn\
├─ General\
│  └─ ManualPath
└─ Logging\
   └─ LogRetentionDays
```

### Document Properties (Workbook Settings)

Stored as Custom Document Properties with `RDD_` prefix:

- `RDD_DefaultGameWidth`, `RDD_DefaultGameHeight`
- `RDD_DefaultBGWidth`, `RDD_DefaultBGHeight`
- `RDD_DefaultUIHeight`
- `RDD_DefaultPerspective`, `RDD_DefaultParallax`, `RDD_DefaultSceneMode`
- `RDD_AutoSyncLists`, `RDD_ShowValidationWarnings`, `RDD_ProtectRoomSheets`

---

## Dependencies

- Microsoft Excel 2019+ (Windows)
- Microsoft Scripting Runtime (scrrun.dll) – for Dictionary objects
- Windows API – for Ctrl+Click detection in chart navigation

---

## Building

1. Open `src/RDD_AddIn.xlsm` in Excel
2. Make changes to VBA modules or Ribbon XML
3. Save as `.xlam`

---

## Contributing

1. Fork the repository
2. Create a feature branch
3. Export changed VBA modules to `src/modules/`
4. Submit a pull request

---

## License

MIT License – see [LICENSE](LICENSE) for details.

---

## Acknowledgments

- **Ron Gilbert** – Puzzle Dependency Chart methodology
- **Adventure Game Studio** – Target game engine
- The classic LucasArts adventure games that inspired this workflow
