# 🔢 TotalPartQtyGenerator  
### Autodesk Inventor – iLogic Production Quantity Multiplier Rule

## 📌 Overview

`TotalPartQtyGenerator` is an iLogic rule designed to calculate the total required quantity of parts for production based on a user-defined number of finished products.

When executed from the top-level assembly, the rule prompts the user to enter the desired number of units to fabricate. It then:

- Counts all part components within the assembly
- Multiplies their occurrence count by the user-defined production quantity
- Stores the resulting total quantity in a structured property/variable

This calculated quantity can then be accessed by downstream automation scripts such as:

- `ExportCutlistDataToDataHub`
- `ExportSheetDimsToDataHub`
- Other fabrication or reporting rules

This script acts as a **quantity multiplier and data propagation mechanism** within the automation workflow.

---

## 🚦 Status

✅ Functional – Ready for Use  
✔ Designed for production batch workflows  
✔ Must be executed from the top-level assembly  

---

## 🎯 Purpose

The primary objective of this rule is to:

- Allow the user to define how many finished products will be fabricated
- Automatically calculate total part requirements
- Store the computed quantities in a reusable format
- Enable other automation rules to retrieve consistent production-level quantities

This ensures fabrication data reflects real production volume rather than single-unit assembly counts.

---

## 🛠️ How It Works

1. The user opens the top-level assembly.
2. The rule is executed from the iLogic browser.
3. A text input dialog appears requesting:
   - Desired number of finished products to fabricate.
4. The script:
   - Traverses the assembly structure
   - Identifies all leaf-level Part components (excluding subassemblies)
   - Calculates their occurrence count
   - Multiplies each count by the user-defined production quantity
5. The resulting total quantities are stored in a variable or document property that can be accessed by other scripts.

---

## 🧠 Core Logic

The rule performs:

- User input handling
- Assembly tree traversal
- Leaf component filtering (Parts only)
- Quantity multiplication
- Centralized quantity storage for reuse

It does **not** currently calculate subassembly-level production quantities.

---

## 📂 File Information

- **TotalPartQtyGenerator.vb**

**Format:** iLogic Rule  
**Execution Context:** Top-level Assembly (.iam)

---

## 🔗 Dependencies

- Autodesk Inventor
- iLogic enabled
- Valid and fully resolved assembly structure
- Downstream scripts that reference the stored quantity variable/property

---

## 📌 Requirements

- Must run from the top-level assembly
- User must provide a valid numeric input
- Assembly structure must be fully loaded and resolved
- Leaf components must be properly defined as Parts

---

## 📈 Workflow Integration

Typical automation sequence:

1. `SetupDataHub`
2. `TotalPartQtyGenerator`
3. `ExportCutlistDataToDataHub`
4. `ExportSheetDimsToDataHub`

This ensures all exported fabrication data reflects the correct production batch size.

---

## 💡 Why This Matters

Without this rule:

- Exported cutlists would reflect single-unit assembly quantities
- Fabrication data would require manual scaling
- Risk of procurement errors increases
- Production documentation becomes inconsistent

With `TotalPartQtyGenerator`:

- Production scaling is automated
- Fabrication datasets are batch-aware
- Downstream scripts operate with consistent quantity logic
- Engineering time spent on manual recalculation is eliminated

---

## 🎥 Demo

_Add demo here (GIF showing quantity input dialog and resulting updated quantities being used by downstream scripts).
