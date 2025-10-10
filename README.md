![License](https://img.shields.io/badge/license-MIT-blue)
[![Project Maintenance](https://img.shields.io/maintenance/yes/2025.svg)](https://github.com/byte-ranger-software/rdd-excel-addin 'GitHub Repository')

# RDD Excel Add‑in (Room Design Document + Puzzle Dependency Chart)

**Purpose**, an Excel add‑in that helps you author Room Design Documents and a Puzzle Dependency Chart in one place, then keep them consistent.  
Designed for Adventure Games made with AGS, with a 9‑verb UI and data‑oriented workflow.

---

## Features

- **Rooms**, create, remove, and rename Room sheets from a template.
- **Dropdown Lists**, two safe update modes,
  - **Update Dropdown Lists**, append missing items only, existing entries remain.
  - **Synchronize Dropdown Lists**, rebuild from Rooms, remove items that no longer exist.  
- **Puzzle List**, rebuild the internal puzzle index for validations and lookups.  
- **Dependency Data**, generate an edge list from all Rooms into the **Data** sheet.  
- **Puzzle Dependency Chart**, create or update the chart from the **Data** sheet.  
- **Validate**, check duplicate IDs, missing references, and hints for cycles.  
- **Export**, quick PDF and CSV exports for stakeholders.  

---

## Repository layout (this add‑in)

```txt
/                      # root
├─ src/
│  ├─ rdd-addin.xlsm   # editable source workbook (ribbon, modules)
│  ├─ modules/         # optional VBA exports (.bas/.cls)
│  ├─ ribbon/          # customUI.xml, callbacks map
│  └─ ribbon_icons/    # PNG icons (16, 24, 32, 48, 64, 128)
├─ bin/
│  └─ rdd.xlam         # add‑in build
├─ docs/
│  └─ MANUAL.md        # Description of the tool and the workflow with the tool
└─ README.md
```
