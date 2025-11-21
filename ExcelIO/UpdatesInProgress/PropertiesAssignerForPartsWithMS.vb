Sub main

' HACER PRUEBA PARA ASIGNAR PROPIEDADES DESDE DENTRO DE UN ENSAMBLE
' INCORPORAR EL ASIGNADOR DENTRO DEL ACTUAL Y SEPARAR CASO SYM DE CASO MS
	
Dim esteDoc As PartDocument = ThisDoc.Document

Dim oCurrentScope As MemberEditScopeEnum
Dim oFactoryDoc As PartDocument

oFactoryDoc = esteDoc.ComponentDefinition.FactoryDocument
oCurrentScope = oFactoryDoc.ComponentDefinition.ModelStates.MemberEditScope


For Each ms As ModelState In oFactoryDoc.ComponentDefinition.ModelStates
	
	ms.Activate
	If ms.Name = "panel_h2" Then
        ms.FactoryDocument.PropertySets.Item("Design Tracking Properties").Item("Description").Value = "Prueba individual"
    End If
	
Next





'For Each state As Object In oFactoryDoc.ComponentDefinition.ModelStates.ModelStateTable.TableRows
'    state    
'	'MsgBox(state.MemberName)
	
'Next


'MsgBox(oFactoryDoc.ComponentDefinition.ModelStates.ModelStateTable.TableRows.Item(1).MemberName)



End sub


' ASSEMBLY PROP ASSIGNER IN PROGRESS: La idea es que los LO se listen en en el excel y que exista una columna adicional donde se agreguen los MS de la forma " (modelStateName)"

Sub Main ()

Dim currentDoc As AssemblyDocument = ThisDoc.Document
	

setPropertiesAndMat(currentDoc) 
currentDoc.Update2
MsgBox("Done")
	
End Sub	


Sub setPropertiesAndMat(assembComp As AssemblyDocument) 
	
	Dim listkeyStrings As New List(Of String)

    For Each compOcc As ComponentOccurrence In assembComp.ComponentDefinition.Occurrences.AllLeafOccurrences
		
		
		
	    If compOcc.Suppressed Or TypeOf compOcc.Definition.Document IsNot PartDocument Or compOcc.BOMStructure.Equals(kNormalBOMStructure) <> True Then
	        
	    Else	
			
		    Dim occDoc As PartDocument = compOcc.Definition.Document ' Acceso al part document de la instancia

	        nameToCompare = occDoc.DisplayName.Substring(0, occDoc.DisplayName.LastIndexOf("."))
'			MsgBox("Full Name: " & occDoc.DisplayName)
'			MsgBox("Test name Trimmed: " & nameToCompare)			
'			MsgBox(occDoc.DisplayName.LastIndexOf("."))
'		    testName = occDoc.DisplayName.Substring(occDoc.DisplayName.LastIndexOf("."), occDoc.DisplayName.Length - occDoc.DisplayName.LastIndexOf(".")) ' occDoc.DisplayName.LastIndexOf(".")
			
'			MsgBox(testName)
			
		    If listkeyStrings.Contains(nameToCompare) Then
				'MsgBox("Repetido: " & occDoc.DisplayName)
				If occDoc.ComponentDefinition.IsModelStateMember Then
				    MsgBox("Repetido con MS: " & occDoc.DisplayName)
					oFactoryDoc = occDoc.ComponentDefinition.FactoryDocument
					oFactoryDoc.PropertySets.Item("Design Tracking Properties").Item("Description").Value = "Test"

				End If
		    Else
		        'MsgBox(nameToCheck)
			     listkeyStrings.Add(nameToCompare)
				 MsgBox("Entry: " & occDoc.DisplayName)
				
'					 Dim oCurrentScope As MemberEditScopeEnum
'					 Dim oFactoryDoc As PartDocument
				
				
'				     If occDoc.ComponentDefinition.IsModelStateMember Then
					
'					     oFactoryDoc = occDoc.ComponentDefinition.FactoryDocument
'					     oCurrentScope = oFactoryDoc.ComponentDefinition.ModelStates.MemberEditScope
		
'				         If oCurrentScope = MemberEditScopeEnum.kEditActiveMember Then
'					         oFactoryDoc.ComponentDefinition.ModelStates.MemberEditScope = MemberEditScopeEnum.kEditAllMembers
'				    	 End If
'				    	 occDoc = oFactoryDoc
'					 End If				
				
					'nameToCompare = occDoc.DisplayName.Substring(0, occDoc.DisplayName.LastIndexOf("."))
				
				
'                 	 If nameToCompare = nameOccDS Then 
'				         MsgBox(nameToCompare)
'                    	 propsAssigner(nameToCompare & ".ipt", file, tab, rowCounter, colPartNo, colStockNumber, colDescription, colMaterial)    'compOcc.Name
					
'	                	 If oFactoryDoc IsNot Nothing Then
'						     oFactoryDoc.ComponentDefinition.ModelStates.MemberEditScope = oCurrentScope
'						 End If
'			         End If
		     End If
	    End If
    Next
    			
End Sub







