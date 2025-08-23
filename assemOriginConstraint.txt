Sub Main ()

Dim currentDoc As AssemblyDocument = ThisApplication.ActiveDocument
AssemblyOriginCons(currentDoc)
	
End Sub


Sub AssemblyOriginCons(assemDef As AssemblyDocument)
	
'Sub that runs through the occurrences of the assembly to constraint their origin planes to those of the current aassembly 
'given that those occurrences have no prior existing constraints or that they are not part of pattern.

'Set planes assembly origin planes
Dim planoAssem1 As WorkPlane = assemDef.ComponentDefinition.WorkPlanes.Item(1)
Dim planoAssem2 As WorkPlane = assemDef.ComponentDefinition.WorkPlanes.Item(2)
Dim planoAssem3 As WorkPlane = assemDef.ComponentDefinition.WorkPlanes.Item(3)
	
For Each compOcc As ComponentOccurrence In assemDef.ComponentDefinition.Occurrences
	If compOcc.Constraints.Count = 0 And compOcc.PatternElement Is Nothing Then
	
	' Set occurrences ogigin planes
	    Dim plano1 As WorkPlane = compOcc.Definition.Workplanes.Item(1) 
	    Dim plano2 As WorkPlane = compOcc.Definition.Workplanes.Item(2) 
	    Dim plano3 As WorkPlane = compOcc.Definition.Workplanes.Item(3) 
		
	' Set proxies for making the constraints
	    Dim planoA1 As WorkPlaneProxy
        compOcc.CreateGeometryProxy(plano1,planoA1)
	    Dim planoA2 As WorkPlaneProxy
        compOcc.CreateGeometryProxy(plano2,planoA2)
	    Dim planoA3 As WorkPlaneProxy
        compOcc.CreateGeometryProxy(plano3, planoA3)
		
	' Constraints application
	    assemDef.ComponentDefinition.Constraints.AddFlushConstraint(planoA1,planoAssem1,0)
	    assemDef.ComponentDefinition.Constraints.AddFlushConstraint(planoA2,planoAssem2,0)
	    assemDef.ComponentDefinition.Constraints.AddFlushConstraint(planoA3,planoAssem3,0)
	
    End If
		
Next

End Sub	
