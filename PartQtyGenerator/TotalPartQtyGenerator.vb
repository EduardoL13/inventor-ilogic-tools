Sub Main ()

Dim currentDoc As AssemblyDocument = ThisDoc.Document


prodQty = InputBox("Enter the desired total qty of products", "TOTAL QTY") ' Multiplicador según el número de productos requerido

If prodQty = ""
	Exit Sub
End If

prodQty = CDbl(prodQty)
'MsgBox(prodQty)

setQty(currentDoc,prodQty)

MsgBox("Done")
	
End Sub	

Sub setQty(assembComp As AssemblyDocument,pQty As Double) ', file As String, tab As String)
	
	Dim listkeyStrings As New List(Of String) ' Inicialización de lista de keystrings individuales

    For Each compOcc As ComponentOccurrence In assembComp.ComponentDefinition.Occurrences.AllLeafOccurrences
		
		If compOcc.Suppressed
			
		Else
		
	    	If TypeOf compOcc.Definition.Document IsNot PartDocument Or compOcc.BOMStructure.Equals(kNormalBOMStructure) <> True Then
	        	'MsgBox(compOcc.Name)
	    	Else	
		    	Dim occDoc As PartDocument = compOcc.Definition.Document ' Acceso al part document de la instancia
	        	nameToCompare = occDoc.DisplayName.Substring(0, occDoc.DisplayName.LastIndexOf("."))
		
			
				Dim oCurrentScope As MemberEditScopeEnum
				Dim oFactoryDoc As PartDocument	
		
		    	If listkeyStrings.Contains(nameToCompare) Then
			
					If occDoc.ComponentDefinition.IsModelStateMember Then

						oFactoryDoc = occDoc.ComponentDefinition.FactoryDocument
						oCurrentScope = oFactoryDoc.ComponentDefinition.ModelStates.MemberEditScope
							
				    	If oCurrentScope = MemberEditScopeEnum.kEditActiveMember Then
					    	oFactoryDoc.ComponentDefinition.ModelStates.MemberEditScope = MemberEditScopeEnum.kEditAllMembers
				    	End If
				    	occDoc = oFactoryDoc
			    	End If					
				
			    	Dim occDefaultProps As PropertySet = occDoc.PropertySets.Item("Design Tracking Properties")
			    	' suma un contador a la propiedad de qty

					valueToUpdate = CDbl(occDefaultProps.Item("Cost Center").Value + 1*pQty)
					occDoc.PropertySets.Item("Design Tracking Properties").Item("Cost Center").Value = valueToUpdate.ToString
					'occProps.Item("Cost Center").Value = valueToUpdate.ToString 
	                 
					If oFactoryDoc IsNot Nothing Then
				    	oFactoryDoc.ComponentDefinition.ModelStates.MemberEditScope = oCurrentScope
					End If				
				
				
				
		    	Else
		        	'MsgBox(nameToCompare)
			     	listkeyStrings.Add(nameToCompare)
							
				
				 	If occDoc.ComponentDefinition.IsModelStateMember Then
					
			        	oFactoryDoc = occDoc.ComponentDefinition.FactoryDocument
				    	oCurrentScope = oFactoryDoc.ComponentDefinition.ModelStates.MemberEditScope
		
				    	If oCurrentScope = MemberEditScopeEnum.kEditActiveMember Then
					    	oFactoryDoc.ComponentDefinition.ModelStates.MemberEditScope = MemberEditScopeEnum.kEditAllMembers
				    	End If
				    	occDoc = oFactoryDoc
				 	End If		

				 	Dim occDefaultProps As PropertySet = occDoc.PropertySets.Item("Design Tracking Properties")
				 	startValue = 1*pQty
                 	occDefaultProps.Item("Cost Center").Value = startValue.ToString
				 	'MsgBox(occDoc.PropertySets.Item("Design Tracking Properties").Item("Cost Center").Value)	
				 
	             	If oFactoryDoc IsNot Nothing Then
				    	oFactoryDoc.ComponentDefinition.ModelStates.MemberEditScope = oCurrentScope
				 	End If
		     End If
	    End If
	End If
    Next
    			
End Sub
