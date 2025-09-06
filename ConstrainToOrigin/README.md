# Constrain To Origin (Inventor iLogic Rule & Macro)

This tool automatically constrains the **origin planes of all component occurrences** in an assembly to the **origin planes of the assembly** itself.  

It is available in two versions:
- 📌 **iLogic rule** (`ConstrainToOrigin.vb`)  
- 📌 **VBA macro** (`ConstrainToOrigin.bas`)  

---

## 🚦 Status
✅ Finished – ready to use.  
Tested in Autodesk Inventor 2025.  

---

## 📂 Files
- `ConstrainToOrigin.vb` → iLogic rule.  
- `ConstrainToOrigin.bas` → VBA macro.  

---

## 🛠️ Setup & Usage

### Option 1 – iLogic Rule
1. Open Autodesk Inventor.  
2. In the **iLogic browser**, create a new rule.  
3. Copy-paste the contents of `ConstrainToOrigin.vb`.  
4. Save the rule.  
5. Open an assembly.  
6. Run the rule → all component origin planes will be constrained to the assembly’s origin planes.  

---

### Option 2 – VBA Macro
1. Open Autodesk Inventor.  
2. Press `Alt + F11` to open the VBA editor.  
3. Insert a new module into the project.  
4. Copy-paste the contents of `ConstrainToOrigin.bas`.  
5. Save and close the VBA editor.  
6. Open an assembly.  
7. Run the macro → all component origin planes will be constrained to the assembly’s origin planes.  

---

## 🎥 Demo
![Constrain To Origin Demo](ConstrainToOrigin/examples/OriginConstraint2.gif)  
  

---

## ⚠️ Notes
- Already constrained components will be skipped.  
- Patterned components are ignored to prevent redundant constraints.  
- Works with both assemblies and subassemblies.  

---

## 📬 Feedback
If you find this useful or have suggestions for improvements, feel free to open an issue or reach out.  

---
