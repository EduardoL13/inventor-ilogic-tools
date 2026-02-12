# 🔢 TotalPartQtyGenerator  
### Autodesk Inventor – iLogic Quantity Aggregation Rule

## 📌 Overview

`TotalPartQtyGenerator` is an iLogic rule that calculates and aggregates the total quantity of each unique part within a top-level assembly and exports the results to the configured Data Hub worksheet.

Instead of manually reviewing assembly structures or relying on static BOM exports, this rule programmatically traverses the assembly tree and generates a structured quantity summary for downstream reporting and fabrication planning.

This script is typically executed after part names and dimensional data have already been exported to the Data Hub.

---

## 🚦 Status

✅ Functional – Ready for Use  
✔ Designed for structured assembly environments  
✔ Requires Data Hub initialization  

---

## 🎯 Purpose

This script automates the calculation of:

- Total quantity of each unique part
- Aggregated counts across nested subassemblies
- Structured quantity output into a Data Hub worksheet

It is particularly useful for:

- Procurement preparation
- Fabrication batching
- Material estimation
- Production planning
- BOM validation

---

## ⚠️ Prerequisites

Before running this rule:

1. The document must be connected to a Data Hub via:
   - `SetupDataHub`
2. The rule must be executed from the top-level assembly
3. The target worksheet must exist
4. Assembly structure must be valid and fully resolved

If the Data Hub has not been initialized, the rule will not function correctly.

---

## 📂 File Information

- **TotalPartQtyGenerator.vb**

**Format:** iLogic Rule  
**Execution Context:** Top-level Assembly (.iam)

---

## 🛠️ How It Works

1. The user opens the top-level assembly.
2. The rule is executed from the iLogic browser.
3. The script:
   - Reads the Data Hub file path and worksheet reference from document properties
   - Traverses the full assembly structure
   - Identifies all leaf components (Parts)
   - Groups identical parts
   - Calculates total occurrences across all levels
   - Writes aggregated quantity data into the designated worksheet

The resulting worksheet contains a clean, structured quantity summary per unique part.

---

## 🧠 Aggregation Logic

The rule performs:

- Assembly tree traversal
- Leaf-level filtering
- Unique part identification
- Quantity accumulation
- Structured output formatting

This ensures accurate total counts even in deeply nested assembly structures.

---

## 🔗 Dependencies

- Autodesk Inventor
- iLogic enabled
- Previously configured Data Hub
- Microsoft Excel (if using Excel as Data Hub)
- Fully resolved assembly structure

---

## 📌 Requirements

- Must run from the top-level assembly
- Data Hub must be initialized
- Worksheet must be accessible
- Suppressed or unresolved components may affect totals

---

## 📈 Workflow Integration

Typical automation sequence:

1. `SetupDataHub`
2. `ExportPartFileNamesToDataHub`
3. `ExportCutlistDataToDataHub`
4. `ExportSheetDimsToDataHub`
5. `TotalPartQtyGenerator`

This rule finalizes the dataset by adding structured quantity intelligence to the Data Hub.

---

## 💡 Why This Matters

Manually calculating part quantities in large assemblies can be:

- Time-consuming
- Error-prone
- Difficult to validate
- Inconsistent across revisions

This rule ensures:

- Accurate quantity aggregation
- Reliable procurement data
- Scalable assembly reporting
- Standardized production documentation

---

## 🎥 Demo

_Add demo here (GIF showing quantity aggregation and worksheet update)._
