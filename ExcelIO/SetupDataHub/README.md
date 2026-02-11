# 📊 SetupDataHub  
### Autodesk Inventor – Data Communication Initialization Script

## 📌 Overview

`SetupDataHub` is a macro designed to initialize the connection between an Autodesk Inventor document (Part or Assembly) and an external data source referred to as the **Data Hub**.

Currently, the Data Hub is intended to be an Excel workbook, but the architecture is conceived so it can be extended to other file types or external data systems in the future.

The script establishes the required document properties that enable structured data exchange (I/O communication) between Inventor and a specified worksheet.

This macro serves as the foundational step for all subsequent automation routines that rely on centralized Excel-driven workflows.

---

## 🚦 Status

✅ Functional – Ready for Use  
✔ Designed for Autodesk Inventor  
✔ Executed from Part or Assembly files  

---

## 🎯 Purpose

The main objective of this script is to:

- Allow the user to select a Data Hub file (currently Excel-based)
- Prompt the user to define a worksheet name
- Create and configure document properties required for structured communication
- Standardize how Inventor documents reference external data

This ensures that downstream automation scripts can reliably read from and write to the designated worksheet.

---

## 🧠 Concept: What Is the Data Hub?

The **Data Hub** acts as a centralized external data container.

It enables:

- Bidirectional communication between Inventor and Excel
- Centralized project data tracking
- BOM-related data management
- Automation chaining across multiple macros

This script does not perform data export or import itself — it prepares the environment for that to happen.

---

## 📂 File Information

- **SetupDataHub.vb**

**Format:** VBA / VB Macro  
**Execution Context:** Run from a Part (.ipt) or Assembly (.iam)

---

## 🛠️ How It Works

1. The user runs the macro from an open Part or Assembly file.
2. A file browser dialog appears.
3. The user selects the desired Data Hub file (Excel workbook).
4. The user is prompted to enter the worksheet name.
5. The macro creates the necessary custom properties in the active document to:
   - Store the Data Hub file path
   - Store the worksheet reference
   - Enable future automation scripts to detect and use these properties

---

## 🔗 Dependencies

- Autodesk Inventor
- Microsoft Excel (if using Excel as Data Hub)
- Windows environment (for file dialog support)

---

## 📌 Requirements

- Must be executed from an open Part or Assembly document.
- The selected Data Hub file must exist.
- The specified worksheet must exist (or follow the intended naming convention).

---

## 📈 Workflow Integration

`SetupDataHub` is intended to be the **first step** in a larger automation ecosystem.

Typical workflow:

1. Run `SetupDataHub`
2. Run data export scripts (e.g., part names, cutlists, sheet dimensions)
3. Run aggregation or reporting scripts
4. Use Excel as centralized tracking system

---

## 💡 Why This Matters

Without this initialization step:

- Subsequent scripts would require manual file selection each time
- Risk of inconsistent data targeting increases
- Workflow automation becomes fragmented

With SetupDataHub:

- All documents become "Data Hub aware"
- Data consistency improves
- Automation becomes scalable

---

## 🔎 Notes

This script initializes the Data Hub connection but does not directly read or write data to Excel.  
It prepares the document environment for future automation routines.
