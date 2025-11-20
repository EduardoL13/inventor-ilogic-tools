Sub main

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
