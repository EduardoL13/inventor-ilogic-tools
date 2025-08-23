# Autodesk Inventor Macros & iLogic Rules

This repository contains a collection of Autodesk Inventor automation tools (macros and iLogic rules) that I’ve created to save time and simplify repetitive modeling tasks.  

The goal is to provide tools that are:  
- ✅ **Practical** → built for real modeling needs.  
- ✅ **Reusable** → adaptable to different projects.  
- ✅ **Open** → free to use and improve.  

---

## 📂 Repository Structure

- **`main` branch** → only finished, tested, and documented tools.  
  - Folder: `finished/`  

- **`dev` branch** → includes experimental or in-progress tools.  
  - Folders: `finished/` + `in-progress/`  

This way, you can choose:
- If you want **stable, reliable macros** → stick to `main`.  
- If you’re curious about **what I’m currently working on** → check out `dev`.  

---

## ✅ Finished Tools (in `main`)
| Tool | Type | Description |
|------|------|-------------|
| `OriginConstraintRule.txt` | iLogic Rule | Automatically constrains the origin planes of all components in an assembly to the assembly’s origin planes. Useful for skeleton modeling workflows where grounding parts is not preferred. |
| `AssemblyOriginCons.txt` | Macro | Automatically constrains the origin planes of all components in an assembly to the assembly’s origin planes. Useful for skeleton modeling workflows where grounding parts is not preferred. |

---

## 🚧 In-Progress Tools (in `dev`)
| Tool | Type | Status |
|------|------|--------|
| `AutoDimensionMacro.bas` | Macro | 🚧 Working on automatic placement of dimensions for specific part templates. |
| `CustomExportRule.iLogic.vb` | iLogic Rule | 🚧 Early version of a rule to batch export components with custom naming. |

---

## 🛠️ How to Use
1. Download the file you need from the `finished/` folder in the `main` branch.  
2. For **iLogic rules** → copy the `.txt` code into an iLogic rule inside Inventor.  
3. For **Macros** → import the `.txt` file into the VBA editor in Inventor.  
4. Run the tool and save time 🚀.  

---

## 💡 Contributing
- If you try one of the tools and improve it, feel free to fork this repo and submit a pull request.  
- If you have an idea for automation, open an issue so we can discuss it.  

---

## 📬 Contact
If you’d like to connect or discuss automation ideas, feel free to reach out via [LinkedIn](www.linkedin.com/in/eduardo-lopez-cobos) or leave a message here.  

---


