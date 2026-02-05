Sub Main ()
'NOTA: Las dimensiones de los valores está dados en cm por defecto
Dim esteDoc As AssemblyDocument = ThisDoc.Document
Dim count As Integer = 0
Dim leafOccurrences As ComponentOccurrencesEnumerator = esteDoc.ComponentDefinition.Occurrences.AllLeafOccurrences


Dim file As String = esteDoc.PropertySets.Item("Spreadsheet Document").Item("File Name").Value
Dim tab As String = InputBox("Enter the excel worksheet name", "Worksheet to Export")

If tab = ""
	Exit Sub
End If
	
Dim ConvFactor As Double = 1/2.54 'Factor de conversión de cm a in
startRow = 8
rowCounter = startRow

partIDColumn = "F" 'Excel Column with PartID
partLengthColumn = "B" 'Dimension or property 
partHeightColumn = "C" 'Atrribute Value

partQtyColumn = "D" 'Excel Column with PartID for thk and materials
partMaterialColumn = "H" 'Material for each sheet/plate 
partThkColumn = "G"' thk for each sheet/plate

Dim listkeyStrings As New List(Of String)

For Each compOccurrence As ComponentOccurrence In leafOccurrences
	Dim occDoc As PartDocument = compOccurrence.Definition.Document
	
	If compOccurrence.Suppressed Then
		
	Else	
	
		If compOccurrence.BOMStructure.Equals(kNormalBOMStructure) <> True Or TypeOf compOccurrence.Definition.Document IsNot PartDocument Or compOccurrence.Suppressed Then
	
		Else
		
			nameToCompare = occDoc.DisplayName.Substring(0, occDoc.DisplayName.LastIndexOf("."))
			'MsgBox(nameToCompare)
    		nameToPrint = occDoc.PropertySets.Item("Design Tracking Properties").Item("Stock Number").Value ' que muestre el part ID
		
	    	If listkeyStrings.Contains(nameToCompare) Then
				
			Else
		    	'MsgBox(nameToCompare)
				listkeyStrings.Add(nameToCompare)


	    		If occDoc.ComponentDefinition.Type.ToString = "kSheetMetalComponentDefinitionObject" Then
				
				
				
		    		If occDoc.ComponentDefinition.HasFlatPattern Then
					
						Dim partModDims As ModelDimensions = occDoc.ComponentDefinition.FlatPattern.ModelAnnotations.ModelDimensions

					 	propsPrinter(occDoc, file, tab, partIDColumn, partQtyColumn, partDescriptionColumn, partLengthColumn, partHeightColumn, partThkColumn, partMaterialColumn, startRow, rowCounter)
					
          	        	rowCounter = rowCounter + 1						
					End If
				End If
			End If
		End If
	End If			
Next
GoExcel.Save
MsgBox("Export Done")
End Sub

Sub propsPrinter(currentPart As PartDocument, file As String, tab As String, partIDColumn As String, partQtyColumn As String, partDescriptionColumn As String, partLengthColumn As String, partHeightColumn As String, partThkColumn As String, partMaterialColumn As String, startRow As Integer, currentRow As Integer)
    Dim ConvFactor As Double = 1 / 2.54 'Factor de conversión de cm a in
'    If currentRow = startRow Then

		Try
			
			
	        GoExcel.CellValue(file, tab, partIDColumn & currentRow) = currentPart.PropertySets.Item("Design Tracking Properties").Item("Stock Number").Value  'confirmar esto
			GoExcel.CellValue(partQtyColumn & currentRow) = Single.Parse(currentPart.PropertySets.Item("Design Tracking Properties").Item("Cost Center").Value) 'Total qty
			GoExcel.CellValue(partLengthColumn & currentRow) = currentPart.ComponentDefinition.FlatPattern.ModelAnnotations.ModelDimensions.Item("length").ModelValue * ConvFactor
			GoExcel.CellValue(partHeightColumn & currentRow) = currentPart.ComponentDefinition.FlatPattern.ModelAnnotations.ModelDimensions.Item("height").ModelValue * ConvFactor			
			GoExcel.CellValue(partThkColumn & currentRow) = currentPart.ComponentDefinition.FlatPattern.ModelAnnotations.ModelDimensions.Item("thk").ModelValue * ConvFactor	
			
			GoExcel.CellValue(partMaterialColumn & currentRow) = currentPart.PropertySets.Item("Design Tracking Properties").Item("Material").Value				

		Catch :

			MsgBox(currentPart.DisplayName & " has missing properties or properties that are not in the valid format")
		    'GoTo ExitHere 
		
		End Try

End Sub
