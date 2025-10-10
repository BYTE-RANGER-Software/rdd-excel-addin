# RDD Add-In Manual

## Description

The add‑in adds a tab **RDD** with four groups.

## Ribbon structure

### Rooms

- *Add Room*, create a new Room from the template, apply IDs and validations.
- *Remove Room*, delete the current Room, check for dependent entries.
- *Update Dropdown Lists*, append missing items from Rooms only, keep existing entries unchanged.
- *Synchronize Dropdown Lists*, completely rewrite Lists from Rooms, remove entries that no longer exist.  

### Puzzles

- *Validate*, check for duplicate IDs, missing references, potential cycles.
- *Refresh Puzzle List*, rebuild the internal puzzle list from all Rooms for validations and lookups.  

### Dependency Chart

- *Generate Data*, collect dependencies from all Rooms and fill the **Data** sheet.
- *Create Chart*, build the Puzzle Dependency Chart and apply layout and styles.
- *Update Chart*, redraw the chart from the **Data** sheet, keep the current view refreshed.  

### Export

- *Export PDF*, export RDD views and the Dependency Chart to PDF.
- *Export CSV*, export Rooms, Puzzles, and Edges as CSV files.  

---

## Data model

The add‑in organizes data into a few well‑known sheets, this keeps RDD and the dependency graph consistent.

### Lists

Reusable dropdown sources, for example Status, Owner, Difficulty, Puzzle Type, Item lists, Verb lists.

### Room

One sheet per room. A named cell `RoomID`, plus a **Puzzles** table with consistent columns, for example,

- `RoomID`, `PuzzleID` (format `R001_P003`), `Title`, `Goal`, `Type`, `DependsOn` (comma‑separated),
- `RequiresItem` (comma‑separated), `GrantsItem` (comma‑separated),
- `Difficulty`, `Owner`, `Status`, `Notes`.

### Data

Generated edge list for the chart,

- `ID` (running), `From` (source PuzzleID), `To` (target PuzzleID), `Type`, `Condition`, `Notes`.

### Chart

The Puzzle Dependency Chart built from **Data**.

### Validation

Reports for duplicate IDs, missing references, optional cycle hints.

> The combined RDD + PDC workflow and sheet roles were discussed and validated during planning.  
