Sub Main ()

Dim currentDoc As AssemblyDocument = ThisDoc.Document

Dim sFile As String
Dim tabDS As String

'sFile = currentDoc.PropertySets.Item("Spreadsheet Document").Item("File Name").Value
'tabDS = currentDoc.PropertySets.Item("Worksheet Data").Item("Worksheet Name").Value	
	

'setQty(currentDoc, sFile, tabDS)
setQty(currentDoc)

MsgBox("Done")
	
End Sub	

Sub setQty(assembComp As AssemblyDocument) ', file As String, tab As String)
	
	Dim listkeyStrings As New List(Of String) ' Inicialización de lista de keystrings individuales

    For Each compOcc As ComponentOccurrence In assembComp.ComponentDefinition.Occurrences.AllLeafOccurrences
		
	    If compOcc.Suppressed Or TypeOf compOcc.Definition.Document IsNot PartDocument Or compOcc.BOMStructure.Equals(kNormalBOMStructure) <> True Then
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
				
			    'Dim occProps As PropertySet = occDoc.PropertySets.Item("Design Tracking Properties")
			    ' suma un contador a la propiedad de qty

				valueToUpdate = CDbl(occDoc.PropertySets.Item("Design Tracking Properties").Item("Cost Center").Value + 1)
				occDoc.PropertySets.Item("Design Tracking Properties").Item("Cost Center").Value = valueToUpdate.ToString
				'occProps.Item("Cost Center").Value = valueToUpdate.ToString 
	                 
				If oFactoryDoc IsNot Nothing Then
				    oFactoryDoc.ComponentDefinition.ModelStates.MemberEditScope = oCurrentScope
				End If				
				
				
				
		    Else
		        'MsgBox(nameToCheck)
			     listkeyStrings.Add(nameToCompare)
							
				
				 If occDoc.ComponentDefinition.IsModelStateMember Then
					
			         oFactoryDoc = occDoc.ComponentDefinition.FactoryDocument
				     oCurrentScope = oFactoryDoc.ComponentDefinition.ModelStates.MemberEditScope
		
				     If oCurrentScope = MemberEditScopeEnum.kEditActiveMember Then
					     oFactoryDoc.ComponentDefinition.ModelStates.MemberEditScope = MemberEditScopeEnum.kEditAllMembers
				     End If
				     occDoc = oFactoryDoc
				 End If		

				 'Dim occProps As PropertySet = occDoc.front
				 startValue = 1
                 occDoc.PropertySets.Item("Design Tracking Properties").Item("Cost Center").Value = startValue.ToString
				 'MsgBox(occDoc.PropertySets.Item("Design Tracking Properties").Item("Cost Center").Value)	
				 
	             If oFactoryDoc IsNot Nothing Then
				     oFactoryDoc.ComponentDefinition.ModelStates.MemberEditScope = oCurrentScope
				 End If
						 
						 

		     End If
	    End If
    Next
    			
End Sub

