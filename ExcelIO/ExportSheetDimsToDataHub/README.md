# 📐 ExportSheetDimsToDataHub  
### Autodesk Inventor – iLogic Sheet Dimension Export Rule

## 📌 Overview

`ExportSheetDimsToDataHub` is an iLogic rule that extracts sheet-related dimensional data from sheet metal components within the active assembly and exports that information into the configured Data Hub worksheet.

This script is intended to automate the extraction of manufacturing-relevant sheet dimensions, reducing manual measurement and ensuring consistent fabrication data reporting.

It is typically used in sheet metal workflows where flat pattern dimensions are required for nesting, procurement, or production tracking.

---

## 🚦 Status

✅ Functional – Ready for Use  
✔ Designed for sheet metal assemblies  
✔ Requires Data Hub initialization  

---

## 🎯 Purpose

This script automates the extraction of sheet-related dimensional information such as:

- Flat pattern length
- Flat pattern width
- Bounding box dimensions
- Thickness
- Other predefined sheet parameters

The extracted data is written directly into the worksheet specified in the Data Hub configuration.

This enables:

- Nesting preparation
- Raw material planning
- Laser cutting optimization
- Production documentation

---

## ⚠️ Prerequisites

Before running this rule:

1. The document must be connected to a Data Hub via:
   - `SetupDataHub`
2. The rule must be executed from the top-level assembly
3. The target worksheet must exist
4. Sheet metal parts must contain valid flat patterns
5. Required dimensional parameters must be available

If flat patterns are not generated, the rule may create them automatically or skip those components depending on implementation.

---

## 📂 File Information

- **ExportSheetDimsToDataHub.vb**

**Format:** iLogic Rule  
**Execution Context:** Top-level Assembly (.iam)

---

## 🛠️ How It Works

1. The user opens the top-level assembly.
2. The rule is executed from the iLogic browser.
3. The script:
   - Reads the Data Hub file path and worksheet reference from document properties
   - Iterates through sheet metal components
   - Accesses flat pattern or bounding box data
   - Extracts predefined dimensional values
   - Writes structured rows into the designated worksheet

The resulting worksheet contains structured sheet dimension data for the entire assembly.

---

## 🔗 Dependencies

- Autodesk Inventor
- iLogic enabled
- Sheet Metal environment
- Previously configured Data Hub
- Microsoft Excel (if using Excel as Data Hub)

---

## 📌 Requirements

- Must run from the top-level assembly
- Data Hub must be initialized
- Worksheet must be accessible
- Sheet metal parts must follow standardized modeling practices
- Flat patterns must be valid

---

## 📈 Workflow Integration

Typical automation sequence:

1. `SetupDataHub`
2. `ExportPartFileNamesToDataHub`
3. `ExportCutlistDataToDataHub`
4. `ExportSheetDimsToDataHub`
5. `TotalPartQtyGenerator`

This script enriches the Data Hub with dimensional fabrication data required for downstream processes.

---

## 💡 Why This Matters

Manually measuring sheet metal components can be:

- Time-consuming
- Inconsistent
- Error-prone

This rule ensures:

- Standardized dimensional reporting
- Faster nesting preparation
- Improved material estimation
- Reduced engineering overhead

---

## 🎥 Demo

_Add demo here (GIF showing sheet dimensions being exported to the Data Hub worksheet)._
