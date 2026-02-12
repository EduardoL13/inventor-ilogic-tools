# 🔄 ImportPropertiesAndMaterialDataFromExcel  
### Autodesk Inventor – iLogic Property & Material Assignment Rule

## 📌 Overview

`ImportPropertiesAndMaterialDataFromExcel` is an iLogic rule designed to import structured metadata from the configured Data Hub (Excel) and assign it to Part components within a top-level assembly.

This rule enables controlled property and material assignment directly from a centralized worksheet, ensuring consistency across all components in a product.

It completes the bidirectional data workflow by allowing Excel-edited metadata to be pushed back into the CAD environment.

---

## 🚦 Status

✅ Functional – Ready for Use  
✔ Designed for structured assembly environments  
✔ Requires prior Data Hub configuration  
✔ Requires prior part name export  

---

## 🎯 Purpose

This script automates the assignment of:

- Description
- Material
- Part Name
- Stock Number
- Other predefined properties (as defined in the rule)

It ensures that:

- Part metadata is standardized
- Manual property editing inside each part file is eliminated
- Excel becomes the single source of truth for part-level metadata

---

## ⚠️ Prerequisites

Before running this rule:

1. The Data Hub must be configured using:
   - `SetupDataHub`
2. Part names must already be exported using:
   - `ExportPartFileNamesToDataHub`
3. The user must manually populate the Excel worksheet with:
   - Desired Description values
   - Material values
   - Part Name values
   - Stock Number values
4. Column headers in Excel must match the property mapping defined inside the script.
5. The rule must be executed from the top-level assembly.

If column names do not match those expected in the script, property assignment may fail or assign incorrect values.

---

## 📂 File Information

- **ImportPropertiesAndMaterialDataFromExcel.vb**

**Format:** iLogic Rule  
**Execution Context:** Top-level Assembly (.iam)

---

## 🛠️ How It Works

1. The user opens the top-level assembly.
2. The rule is executed from the iLogic browser.
3. The script:
   - Reads the Data Hub file path and worksheet reference from document properties
   - Matches part names in the assembly with rows in the Excel worksheet
   - Retrieves property and material values from predefined columns
   - Opens each corresponding Part document
   - Assigns:
     - iProperties (Description, Stock Number, etc.)
     - Material property
4. The updated properties are saved to the respective part files.

---

## 🧠 Data Mapping Logic

The rule performs:

- Assembly traversal
- Part name matching
- Row-to-component alignment
- Column-to-property mapping
- Property assignment
- Material update

The script relies on consistent naming conventions between:

- Excel worksheet entries
- Inventor part display names

---

## 🔗 Dependencies

- Autodesk Inventor
- iLogic enabled
- Previously configured Data Hub
- Microsoft Excel
- Proper column-to-property alignment inside the script

---

## 📌 Requirements

- Must run from the top-level assembly
- Data Hub must be initialized
- Part names must already exist in Excel
- Column headers must match those defined in the script
- Parts must be writable (not read-only or Vault-locked)

---

## 📈 Workflow Integration

Typical automation sequence:

1. `SetupDataHub`
2. `ExportPartFileNamesToDataHub`
3. User edits Excel Data Hub (adds descriptions, materials, stock numbers)
4. `ImportPropertiesAndMaterialDataFromExcel`
5. Continue with export/aggregation scripts as needed

This establishes Excel as the centralized control interface for part metadata.

---

## 💡 Why This Matters

Manually assigning properties and materials to multiple parts can be:

- Repetitive
- Error-prone
- Inconsistent across revisions
- Difficult to audit

With this rule:

- Metadata becomes centrally managed
- Property updates are scalable
- CAD files remain synchronized with structured data
- Engineering workflow becomes significantly more efficient

---

## 🎥 Demo

_Add demo here (GIF showing Excel property editing followed by automatic property assignment in Inventor)._
