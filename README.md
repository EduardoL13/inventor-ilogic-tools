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

' Sub Main ()

' Dim currentDoc As AssemblyDocument = ThisApplication.ActiveDocument
' AssemblyOriginCons(currentDoc)
	
' End Sub


' Sub AssemblyOriginCons(assemDef As AssemblyDocument)
	
' 'Sub that runs through the occurrences of the assembly to constraint their origin planes to those of the current aassembly 
' 'given that those occurrences have no prior existing constraints or that they are not part of pattern.

' 'Set planes assembly origin planes
' Dim PlanoE1 As WorkPlane = assemDef.ComponentDefinition.WorkPlanes.Item(1)
' Dim PlanoE2 As WorkPlane = assemDef.ComponentDefinition.WorkPlanes.Item(2)
' Dim PlanoE3 As WorkPlane = assemDef.ComponentDefinition.WorkPlanes.Item(3)
	
' For Each compOcc As ComponentOccurrence In assemDef.ComponentDefinition.Occurrences
' 	If compOcc.Constraints.Count = 0 And compOcc.PatternElement Is Nothing Then
	
' 	' Set occurrences ogigin planes
' 	    Dim Plano1 As WorkPlane = compOcc.Definition.Workplanes.Item(1) 
' 	    Dim Plano2 As WorkPlane = compOcc.Definition.Workplanes.Item(2) 
' 	    Dim Plano3 As WorkPlane = compOcc.Definition.Workplanes.Item(3) 
		
' 	' Set proxies for making the constraints
' 	    Dim APlano1 As WorkPlaneProxy
'        compOcc.CreateGeometryProxy(Plano1,APlano1)
' 	    Dim APlano2 As WorkPlaneProxy
'         compOcc.CreateGeometryProxy(Plano2,APlano2)
' 	    Dim APlano3 As WorkPlaneProxy
'        compOcc.CreateGeometryProxy(Plano3, APlano3)
		
' 	' Constraints application
' 	    assemDef.ComponentDefinition.Constraints.AddFlushConstraint(APlano1,PlanoE1,0)
' 	    assemDef.ComponentDefinition.Constraints.AddFlushConstraint(APlano2,PlanoE2,0)
' 	    assemDef.ComponentDefinition.Constraints.AddFlushConstraint(APlano3,PlanoE3,0)
	
'     End If
		
' Next

' End Sub	

