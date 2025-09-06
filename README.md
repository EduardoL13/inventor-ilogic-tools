# Autodesk Inventor Automation – Macros & iLogic Rules

This repository contains a collection of Autodesk Inventor automation tools (macros and iLogic rules) that I’ve created to save time and simplify repetitive modeling tasks.  

The goal is to provide tools that are:  
- ✅ **Practical** → built for real modeling needs.  
- ✅ **Reusable** → adaptable to different projects.  
- ✅ **Open** → free to use and improve.  

---

## 📂 Repository Structure

- **`main` branch** → only finished, tested, and documented tools.  
- **`in-progress` branch** → includes experimental or unfinished tools.  

👉 Use `main` if you want **stable, reliable macros**.  
👉 Use `in-progress` if you’re curious about **what I’m currently developing**.  

---

## ✅ Finished Tools (in `main`)
| Tool | Type | Description |
|------|------|-------------|
| **CommonFolder** | Directory | Contains shared VBA macros that are required by some of the iLogic rules in this repository. They provide utility functions (e.g., browsing for a folder location). |
| **ConstrainToOrigin** | Directory | Contains macro and iLogic versions of a routine to automatically constrain the origin planes of all component occurrences in an assembly to the origin planes of the assembly itself. |
| **ExportFlatPatternsToDxf** | Directory | Contains an iLogic rule that automates the process of exporting all sheet metal parts in the active assembly as DXF files. Instead of opening each part manually and exporting its flat pattern, you can generate all DXFs in just one click. |

---

## 🚧 In-Progress Tools (in `in-progress`)
| Tool | Type | Description |
|------|------|-------------|
| **ExcelIO** | Directory | This directory will be updated in the coming weeks with macros and iLogic rules for importing/exporting data (properties, annotations, key parameters, etc.) into a specified worksheet within an Excel file. |

---

## 🛠️ How to Use

### For iLogic Rules
1. Download the `.vb` file you need.  
2. In Inventor, open the **iLogic browser** and create a new rule.  
3. Copy-paste the contents of the file into the rule editor.  
4. Save and run the rule.  

### For Macros
1. Download the `.bas` file you need.  
2. In Inventor, press `Alt + F11` to open the VBA editor.  
3. Insert a new module into the project.  
4. Copy-paste or import the `.bas` file into the module.  
5. Save and run the macro.  

---

## 💡 Contributing
- If you try one of the tools and improve it, feel free to fork this repo and submit a pull request.  
- If you have an idea for automation, open an issue so we can discuss it.  

---

## 📬 Contact
If you’d like to connect or discuss automation ideas, feel free to reach out on [LinkedIn](https://www.linkedin.com/in/eduardo-lopez-cobos) or leave a message here.  

---
