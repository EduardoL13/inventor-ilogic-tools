Sub Main ()
'NOTA: Las dimensiones de los valores está dados en cm por defecto
Dim esteDoc As AssemblyDocument = ThisDoc.Document
Dim count As Integer = 0
Dim leafOccurrences As ComponentOccurrencesEnumerator = esteDoc.ComponentDefinition.Occurrences.AllLeafOccurrences


Dim file As String = esteDoc.PropertySets.Item("Spreadsheet Document").Item("File Name").Value
Dim tab As String = InputBox("Enter the excel worksheet name", "Worksheet to Export")

Dim ConvFactor As Double = 1/2.54 'Factor de conversión de cm a in
Dim rowCounter = 6

partIDColumn = "D" 'Excel Column with PartID
partLengthColumn = "B" 'Dimension or property 

partQtyColumn = "C" 'Excel Column with PartID for thk and materials
partDescriptionColumn = "E"
'partMaterialColumn = "H" 'Material for each sheet/plate 

Dim listkeyStrings As New List(Of String)

For Each compOccurrence As ComponentOccurrence In leafOccurrences
	Dim occDoc As PartDocument = compOccurrence.Definition.Document
	
	
	

	
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
				
			Else
					
			    Dim partModDims As ModelDimensions = occDoc.ComponentDefinition.ModelAnnotations.ModelDimensions
				If rowCounter = 8 Then
	        	    GoExcel.CellValue(file, tab, partIDColumn & rowCounter) = nameToPrint 'confirmar esto
					GoExcel.CellValue(partQtyColumn & rowCounter) = Single.Parse(occDoc.PropertySets.Item("Design Tracking Properties").Item("Cost Center").Value) 'Total qty
		
'					GoExcel.CellValue(partMaterialColumn & rowCounter) = occDoc.PropertySets.Item("Design Tracking Properties").Item("Material").Value
					GoExcel.CellValue(partDescriptionColumn & rowCounter) = occDoc.PropertySets.Item("Design Tracking Properties").Item("Description").Value
						
					GoExcel.CellValue(partLengthColumn & rowCounter) = partModDims.Item("length").ModelValue * ConvFactor
						
						
				Else
	        		GoExcel.CellValue(file, tab, partIDColumn & rowCounter) = nameToPrint 'confirmar esto
					GoExcel.CellValue(partQtyColumn & rowCounter) = Single.Parse(occDoc.PropertySets.Item("Design Tracking Properties").Item("Cost Center").Value) 'Total qty
		
'					GoExcel.CellValue(partMaterialColumn & rowCounter) = occDoc.PropertySets.Item("Design Tracking Properties").Item("Material").Value
					GoExcel.CellValue(partDescriptionColumn & rowCounter) = occDoc.PropertySets.Item("Design Tracking Properties").Item("Description").Value
						
					GoExcel.CellValue(partLengthColumn & rowCounter) = partModDims.Item("length").ModelValue * ConvFactor

						
			    End If
          	    rowCounter = rowCounter + 1						

			End If
		End If
	End If
Next
GoExcel.Save
MsgBox("Export Done")
End Sub
