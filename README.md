# 🔧 Inventor iLogic – Auto Constrain Planes to Assembly Origin

This iLogic rule automatically adds constraints between the **origin planes of every component** in an assembly and the **origin planes of the assembly itself**.  

It’s a small but powerful time-saver if you:  
- Work with **multibody skeleton modeling**.  
- Prefer **fully constrained parts** instead of grounding them by default.  
- Want to avoid the repetitive task of manually constraining each component’s planes.  

---

## 📌 Example Use Case
When you create an assembly from multibodies using **"Make Components"**, Inventor grounds each part. If you prefer a **fully constrained workflow**, this iLogic rule aligns all components’ origin planes with the assembly origin automatically.

Instead of adding constraints one by one, you can run this script and get everything constrained in **seconds**.

---

## ▶️ How to Use
For ilogic:
1. Open your Inventor assembly.  
2. Go to the **iLogic Rule Editor**.  
3. Copy-paste the code from [`AutoConstrainPlanes.vb`](AutoConstrainPlanes.vb).  
4. Run the rule → all components will be constrained to the assembly origin.  

To use it as a macro:
1. Access the vb editor in the manage tab
2. Create a module in the ApplicationDefault (if you want to be able to access the macro in every assembly you open)
3. Copy and paste de macro version code
4. run the macro for the ApplicationDefault project
---

## 💻 Code
Here’s the full code:  

```vbnet
' Auto Constrain Planes to Assembly Origin
' Author: Eduardo López


