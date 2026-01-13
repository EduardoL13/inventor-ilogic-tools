Sub Main ()
'NOTA: Las dimensiones de los valores está dados en cm por defecto
Dim esteDoc As AssemblyDocument = ThisDoc.Document
Dim count As Integer = 0
Dim leafOccurrences As ComponentOccurrencesEnumerator = esteDoc.ComponentDefinition.Occurrences.AllLeafOccurrences


Dim file As String = esteDoc.PropertySets.Item("Spreadsheet Document").Item("File Name").Value
Dim tab As String = InputBox("Enter the excel worksheet name", "Worksheet to Export")

Dim ConvFactor As Double = 1/2.54 'Factor de conversión de cm a in
Dim rowCounter = 8

partIDColumn = "F" 'Excel Column with PartID
partLengthColumn = "B" 'Dimension or property 
partHeightColumn = "C" 'Atrribute Value

partQtyColumn = "D" 'Excel Column with PartID for thk and materials
partMaterialColumn = "H" 'Material for each sheet/plate 
partThkColumn = "G"' thk for each sheet/plate

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
				
				
				
		    	If occDoc.ComponentDefinition.HasFlatPattern Then
					
					Dim partModDims As ModelDimensions = occDoc.ComponentDefinition.FlatPattern.ModelAnnotations.ModelDimensions
				    If rowCounter = 8 Then
	        			GoExcel.CellValue(file, tab, partIDColumn & rowCounter) = nameToPrint 'confirmar esto
						GoExcel.CellValue(partQtyColumn & rowCounter) = Single.Parse(occDoc.PropertySets.Item("Design Tracking Properties").Item("Cost Center").Value) 'Total qty
		
						GoExcel.CellValue(partMaterialColumn & rowCounter) = occDoc.PropertySets.Item("Design Tracking Properties").Item("Material").Value
'						GoExcel.CellValue(partThkColumn & rowCounter) = occDoc.ComponentDefinition.Parameters.ModelParameters("Thickness").Value*ConvFactor
						GoExcel.CellValue(partThkColumn & rowCounter) = partModDims.Item("thk").ModelValue * ConvFactor
						
						GoExcel.CellValue(partLengthColumn & rowCounter) = partModDims.Item("length").ModelValue * ConvFactor
						GoExcel.CellValue(partHeightColumn & rowCounter) = partModDims.Item("Height").ModelValue * ConvFactor
						
						
					Else
	        			GoExcel.CellValue(file, tab, partIDColumn & rowCounter) = nameToPrint 'confirmar esto
						GoExcel.CellValue(partQtyColumn & rowCounter) = Single.Parse(occDoc.PropertySets.Item("Design Tracking Properties").Item("Cost Center").Value) 'Total qty
		
						GoExcel.CellValue(partMaterialColumn & rowCounter) = occDoc.PropertySets.Item("Design Tracking Properties").Item("Material").Value
'						GoExcel.CellValue(partThkColumn & rowCounter) = occDoc.ComponentDefinition.Parameters.ModelParameters("Thickness").Value*ConvFactor
						GoExcel.CellValue(partThkColumn & rowCounter) = partModDims.Item("thk").ModelValue * ConvFactor
						
						GoExcel.CellValue(partLengthColumn & rowCounter) = partModDims.Item("length").ModelValue * ConvFactor
						GoExcel.CellValue(partHeightColumn & rowCounter) = partModDims.Item("height").ModelValue * ConvFactor 
						
			        End If
          	        rowCounter = rowCounter + 1						
				End If
			End If
		End If
	End If
Next
GoExcel.Save
MsgBox("Export Done")
End Sub
