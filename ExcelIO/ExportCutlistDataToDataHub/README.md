# 🪚 ExportCutlistDataToDataHub  
### Autodesk Inventor – iLogic Cutlist Data Export Rule

## 📌 Overview

`ExportCutlistDataToDataHub` is an iLogic rule that extracts cutlist-related data from all relevant components within the active assembly and exports that structured information into the configured Data Hub worksheet.

This script is intended to automate fabrication data preparation by centralizing cutlist information inside an Excel-based workflow.

It builds upon the initialization performed by `SetupDataHub` and typically follows the execution of `ExportPartFileNamesToDataHub`.

---

## 🚦 Status

✅ Functional – Ready for Use  
✔ Designed for structured fabrication workflows  
✔ Requires Data Hub initialization  

---

## 🎯 Purpose

This script automates the extraction of cutlist-related data such as:

- Component identifiers
- Dimensional parameters
- Material-related properties
- Other fabrication-relevant metadata

The extracted information is written directly into the worksheet specified in the Data Hub configuration.

This enables:

- Fabrication planning
- Procurement preparation
- Laser/CNC cut preparation
- Project-level reporting

---

## ⚠️ Prerequisites

Before running this rule:

1. The document must already be connected to a Data Hub via:
   - `SetupDataHub`
2. The rule must be executed from the top-level assembly
3. The target worksheet must exist
4. Parts must contain the required parameters or properties used for cutlist extraction

If required parameters are missing, the rule may skip entries or generate incomplete rows.

---

## 📂 File Information

- **ExportCutlistDataToDataHub.vb**

**Format:** iLogic Rule  
**Execution Context:** Top-level Assembly (.iam)

---

## 🛠️ How It Works

1. The user opens the top-level assembly.
2. The rule is executed from the iLogic browser.
3. The script:
   - Reads the Data Hub file path and worksheet reference from document properties
   - Iterates through relevant leaf components
   - Extracts predefined cutlist parameters and properties
   - Writes structured rows into the designated worksheet

The resulting worksheet becomes a centralized fabrication dataset for the entire assembly.

---

## 🔗 Dependencies

- Autodesk Inventor
- iLogic enabled
- Previously configured Data Hub
- Microsoft Excel (if using Excel as Data Hub)
- Required parameters defined consistently across parts

---

## 📌 Requirements

- Must run from the top-level assembly
- Data Hub must be initialized
- Worksheet must be accessible
- Parts must follow a standardized parameter naming convention

---

## 📈 Workflow Integration

Typical automation chain:

1. `SetupDataHub`
2. `Type the name of the target worksheet for the data`
3. `ExportCutlistDataToDataHub`

This script plays a central role in generating fabrication-ready datasets.

---

## 💡 Why This Matters

Manual cutlist extraction from large assemblies can be:

- Repetitive
- Inconsistent
- Prone to missing data
- Difficult to scale across projects

This rule ensures:

- Standardized cutlist data structure
- Faster fabrication package preparation
- Reduced human error
- Improved data traceability

---

## 🎥 Demo

_Add demo here (GIF showing cutlist data being exported to the Data Hub worksheet)._
