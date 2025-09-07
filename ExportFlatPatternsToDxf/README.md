# Export Flat Patterns to DXF (Inventor iLogic Rule)

This iLogic rule automates the process of exporting **all sheet metal parts in the active assembly** as DXF files.  
Instead of opening each part manually and exporting its flat pattern, you can generate all DXFs in just one click.

---

## 🚦 Status
✅ Finished – ready to use.  
Tested in Autodesk Inventor 2025.  

---

## 📂 Files
- `ExportFlatPatternsToDxf.vb` → iLogic rule + macro.
- `ExportFlatPatternsToDxfDirect.vb` → iLogic rule
- `BrowseFileLocation.bas` → VBA macro required by the rule (provides the folder selection dialog). This macro is available in CommonFolder with a demo that shows Setup & Usage step 1. described below

---

`ExportFlatPatternsToDxfDirect.vb`
## 🛠️ Setup & Usage

1. **Add the iLogic rule:**
   - Open an assembly that contains sheet metal parts. 
   - In Inventor, open the **iLogic browser**.  
   - Create a new rule and paste the contents of `ExportFlatPatternsToDxfDirect.vb`.   

2. **Run the rule:** 
   - Run the iLogic rule.  
   - A folder dialog will appear.  
   - Select the output folder.  
   - All sheet metal flat patterns will be exported as DXF files into that folder. 

---

`ExportFlatPatternsToDxf.vb`
## 🛠️ Setup & Usage
1. **Add the macro to Inventor VBA:**
   - Open Inventor.  
   - Press `Alt + F11` to open the VBA editor or go and click VBA Editor in the Tools tab above.  
   - Insert a new module into a project (ApplicationProject is a good option because it is available for every document) .  
   - Copy-paste the contents of `BrowseFileLocation.bas` into the inserted module.
   - Save and close the VBA editor.  

2. **Add the iLogic rule:**
   - Open an assembly that contains sheet metal parts. 
   - In Inventor, open the **iLogic browser**.  
   - Create a new rule and paste the contents of `ExportFlatPatternsToDxf.vb`.  
   - ⚠️ In lines **8 and 9** of the rule, update the values of the variables `projectVba` and `moduleVba` to match the name of the VBA project and module where you pasted the `BrowseFileLocation` macro.  

3. **Run the rule:** 
   - Run the iLogic rule.  
   - A folder dialog (powered by the VBA macro) will appear.  
   - Select the output folder.  
   - All sheet metal flat patterns will be exported as DXF files into that folder.  

---

## 🎥 Demo
![DXF Export Demo](examples/ExportFlatPatternsToDxf.gif)  
 

---

## ⚠️ Notes
- Non-sheet metal parts will be skipped automatically.  
- If a part doesn’t have a flat pattern yet, Inventor will create it.  
- The VBA macro **must** be installed before running the rule, and the variables `projectVba` and `moduleVba` must be set correctly.  

---

## 📬 Feedback
If you find this useful or have ideas for improvements, feel free to open an issue or reach out.  

---
