# 📁 ExportPartFileNamesToDataHub  
### Autodesk Inventor – iLogic Assembly Data Export Rule

## 📌 Overview

`ExportPartFileNamesToDataHub` is an iLogic rule designed to extract all leaf component file names from a top-level assembly and export them to a predefined worksheet in the Data Hub.

The rule prints the **Display Names** (file names without extension) of all leaf occurrences (Parts only) into the target worksheet previously configured using the `SetupDataHub` script.

This automation eliminates the need to manually extract part lists and ensures structured data consistency across projects.

---

## 🚦 Status

✅ Functional – Ready for Use  
✔ Tested within assembly environments  
✔ Requires prior Data Hub initialization  

---

## 🎯 Purpose

This script:

- Iterates through all leaf occurrences in the top-level assembly
- Identifies Part files (excluding subassemblies)
- Extracts each part’s Display Name (file name without extension)
- Writes the list into the target worksheet defined in the Data Hub

It is typically used as an early-stage data extraction step for:

- BOM preparation
- Cutlist processing
- Quantity aggregation
- Fabrication tracking

---

## ⚠️ Prerequisites

Before running this rule:

1. A Data Hub must already be configured using:
   - `SetupDataHub`
2. The rule must be executed from the **top-level assembly**
3. The target worksheet must exist in the Data Hub file

If the Data Hub has not been initialized, this rule will not function correctly.

---

## 📂 File Information

- **ExportPartFileNamesToDataHub.vb**

**Format:** iLogic Rule  
**Execution Context:** Top-level Assembly (.iam)

---

## 🛠️ How It Works

1. The user opens the top-level assembly.
2. The rule is executed from the iLogic browser.
3. The script:
   - Reads the Data Hub file path and worksheet name stored in document properties
   - Iterates through all leaf occurrences
   - Filters only Part documents
   - Extracts each part’s Display Name
4. The part names are written into the designated worksheet inside the Data Hub.

---

## 🔗 Dependencies

- Autodesk Inventor
- iLogic enabled
- Previously configured Data Hub (via `SetupDataHub`)
- Microsoft Excel (if using Excel as Data Hub)

---

## 📌 Requirements

- Must be executed from the top-level assembly
- Data Hub must already be linked to the document
- Worksheet must exist and be accessible
- Assembly structure must allow proper traversal of leaf occurrences

---

## 📈 Workflow Integration

Typical automation sequence:

1. Run `SetupDataHub`
2. Run `ExportPartFileNamesToDataHub`
3. Run downstream scripts (cutlist export, sheet dimensions, quantity aggregation)
4. Perform reporting or fabrication processing in Excel

This script provides the foundational dataset for subsequent automation routines.

---

## 💡 Why This Matters

Manually extracting part names from large assemblies can be:

- Time-consuming
- Error-prone
- Inconsistent across projects

This rule ensures:

- Standardized data extraction
- Reliable part tracking
- Seamless integration into Excel-driven workflows

---

## 🎥 Demo

_Add demo here (GIF showing the rule execution and worksheet output)._
