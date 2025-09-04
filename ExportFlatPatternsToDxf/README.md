# Export Flat Patterns to DXF (Inventor iLogic Rule)

This iLogic rule automates the process of exporting **all sheet metal parts in the active assembly** as DXF files.  
Instead of opening each part manually and exporting its flat pattern, you can generate all DXFs in just one click.

---

## 🚦 Status
✅ Finished – ready to use.  
Tested in Autodesk Inventor 20XX.  

---

## 📂 Files
- `ExportFlatPatternsToDxf.vb` → iLogic rule.  
- `BrowseFileLocation.bas` → VBA macro required by the rule (provides the folder selection dialog).  

---

## 🛠️ Setup & Usage
1. **Add the macro to Inventor VBA:**
   - Open Inventor.  
   - Press `Alt + F11` to open the VBA editor.  
   - Insert a new module into the project.  
   - Copy-paste the contents of `BrowseFileLocation.bas` into that module.  
   - Save and close the VBA editor.  

2. **Add the iLogic rule:**
   - In Inventor, open the **iLogic browser**.  
   - Create a new rule and paste the contents of `ExportFlatPatternsToDxf.vb`.  
   - ⚠️ In lines **8 and 9** of the rule, update the values of the variables `projectVba` and `moduleVba` to match the name of the VBA project and module where you pasted the `BrowseFileLocation` macro.  

3. **Run the rule:**
   - Open an assembly that contains sheet metal parts.  
   - Run the iLogic rule.  
   - A folder dialog (powered by the VBA macro) will appear.  
   - Select the output folder.  
   - All sheet metal flat patterns will be exported as DXF files into that folder.  

---

## 🎥 Demo
![DXF Export Demo](../examples/ExportFlatPatternsToDxf.gif)  
*(Replace with the actual path to your GIF in the repo)*  

---

## ⚠️ Notes
- Non-sheet metal parts will be skipped automatically.  
- If a part doesn’t have a flat pattern yet, Inventor will prompt you to generate it.  
- The VBA macro **must** be installed before running the rule, and the variables `projectVba` and `moduleVba` must be set correctly.  

---

## 📬 Feedback
If you find this useful or have ideas for improvements, feel free to open an issue or reach out.  

---
