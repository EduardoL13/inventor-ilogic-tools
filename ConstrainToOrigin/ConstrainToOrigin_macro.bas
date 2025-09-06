Sub AssemblyOriginCons()
    
'Sub that runs through the occurrences of the assembly to constraint their origin planes to those of the current aassembly
'given that those occurrences have no prior existing constraints or that they are not part of pattern.

    Dim currentDoc As AssemblyDocument
    Set currentDoc = ThisApplication.ActiveDocument

'Set planes assembly origin planes
    Dim planoAssem1 As WorkPlane
    Set planoAssem1 = currentDoc.ComponentDefinition.WorkPlanes.Item(1)
    
    Dim planoAssem2 As WorkPlane
    Set planoAssem2 = currentDoc.ComponentDefinition.WorkPlanes.Item(2)
    
    Dim planoAssem3 As WorkPlane
    Set planoAssem3 = currentDoc.ComponentDefinition.WorkPlanes.Item(3)
    
    Dim compOcc As ComponentOccurrence
    
    For Each compOcc In currentDoc.ComponentDefinition.Occurrences
        If compOcc.Constraints.Count = 0 And compOcc.PatternElement Is Nothing Then
    
    ' Set occurrences ogigin planes
        Dim plano1 As WorkPlane
        Set plano1 = compOcc.Definition.WorkPlanes.Item(1)
        
        Dim plano2 As WorkPlane
        Set plano2 = compOcc.Definition.WorkPlanes.Item(2)
        
        Dim plano3 As WorkPlane
        Set plano3 = compOcc.Definition.WorkPlanes.Item(3)
        
    ' Set proxies for making the constraints
        Dim planoA1 As WorkPlaneProxy
        compOcc.CreateGeometryProxy plano1, planoA1
        
        Dim planoA2 As WorkPlaneProxy
        compOcc.CreateGeometryProxy plano2, planoA2
        
        Dim planoA3 As WorkPlaneProxy
        compOcc.CreateGeometryProxy plano3, planoA3
        
    ' Constraints application
        currentDoc.ComponentDefinition.Constraints.AddFlushConstraint planoA1, planoAssem1, 0

        currentDoc.ComponentDefinition.Constraints.AddFlushConstraint planoA2, planoAssem2, 0
        
        currentDoc.ComponentDefinition.Constraints.AddFlushConstraint planoA3, planoAssem3, 0
    
        End If
        
    Next

End Sub
