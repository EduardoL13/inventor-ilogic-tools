ub Main ()
'NOTA: Las dimensiones de los valores está dados en cm por defecto
Dim esteDoc As AssemblyDocument = ThisDoc.Document
Dim count As Integer = 0
Dim leafOccurrences As ComponentOccurrencesEnumerator = esteDoc.ComponentDefinition.Occurrences.AllLeafOccurrences


Dim file As String = esteDoc.PropertySets.Item("Spreadsheet Document").Item("File Name").Value
'Dim tab As String = "CAD_output_shpl"
Dim tab As String = InputBox("Enter the excel worksheet name", "Worksheet to Export")

Dim ConvFactor As Double = 1/2.54 'Factor de conversión de cm a in
Dim rowCounter = 2
For Each compOccurrence As ComponentOccurrence In leafOccurrences
	Dim occDoc As PartDocument = compOccurrence.Definition.Document
	'nameToPrint = occDoc.DisplayName.Substring(0, occDoc.DisplayName.LastIndexOf("."))
    nameToPrint = occDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value
	
	If compOccurrence.Visible = False Or TypeOf compOccurrence.Definition.Document IsNot PartDocument Or compOccurrence.Suppressed Then
	
	Else
	If occDoc.ComponentDefinition.Type.ToString = "kSheetMetalComponentDefinitionObject" Then
		If occDoc.ComponentDefinition.HasFlatPattern Then
	        For Each partModDim As ModelDimension In occDoc.ComponentDefinition.FlatPattern.ModelAnnotations.ModelDimensions
	            If rowCounter = 2 Then
	                GoExcel.CellValue(file, tab, "A" & rowCounter) = nameToPrint 'confirmar esto
					GoExcel.CellValue("B" & rowCounter) = partModDim.Name
					GoExcel.CellValue("C" & rowCounter) = partModDim.ModelValue * ConvFactor
				Else
	   		    	GoExcel.CellValue("A" & rowCounter) = nameToPrint 'confirmar esto
					GoExcel.CellValue("B" & rowCounter) = partModDim.Name
					GoExcel.CellValue("C" & rowCounter) = partModDim.ModelValue * ConvFactor
				End If
	    	rowCounter = rowCounter + 1 
			Next
	        For Each partModAnno As Object In occDoc.ComponentDefinition.FlatPattern.ModelAnnotations.ModelLeaderNotes	
	            If rowCounter = 2 Then
	                GoExcel.CellValue(file, tab, "A" & rowCounter) = nameToPrint 'confirmar esto
					GoExcel.CellValue("B" & rowCounter) = partModAnno.Name
					GoExcel.CellValue("C" & rowCounter) = partModAnno.InternalName
				Else
	   		    	GoExcel.CellValue("A" & rowCounter) = nameToPrint 'confirmar esto
					GoExcel.CellValue("B" & rowCounter) = partModAnno.Name
					GoExcel.CellValue("C" & rowCounter) = partModAnno.Definition.Text.Text
				End If
	    	rowCounter = rowCounter + 1 
			Next			    
			
			
 	    Else
	       For Each partModDim As ModelDimension In occDoc.ComponentDefinition.ModelAnnotations.ModelDimensions
	            If rowCounter = 2 Then
	                GoExcel.CellValue(file, tab, "A" & rowCounter) = nameToPrint 'confirmar esto
					GoExcel.CellValue("B" & rowCounter) = partModDim.Name
					GoExcel.CellValue("C" & rowCounter) = partModDim.ModelValue * ConvFactor
				Else
	   		    	GoExcel.CellValue("A" & rowCounter) = nameToPrint 'confirmar esto
					GoExcel.CellValue("B" & rowCounter) = partModDim.Name
					GoExcel.CellValue("C" & rowCounter) = partModDim.ModelValue * ConvFactor
				End If
	    	rowCounter = rowCounter + 1 
			Next
		    For Each partModAnno As Object In occDoc.ComponentDefinition.ModelAnnotations.ModelLeaderNotes	
	            If rowCounter = 2 Then
	                GoExcel.CellValue(file, tab, "A" & rowCounter) = nameToPrint 'confirmar esto
					GoExcel.CellValue("B" & rowCounter) = partModAnno.Name
					GoExcel.CellValue("C" & rowCounter) = partModAnno.InternalName
				Else
	   		    	GoExcel.CellValue("A" & rowCounter) = nameToPrint 'confirmar esto
					GoExcel.CellValue("B" & rowCounter) = partModAnno.Name
					GoExcel.CellValue("C" & rowCounter) = partModAnno.Definition.Text.Text
				End If
	    	rowCounter = rowCounter + 1 
			Next		
		End If
	Else
	    For Each partModDim As ModelDimension In occDoc.ComponentDefinition.ModelAnnotations.ModelDimensions
	        If rowCounter = 2 Then
	            GoExcel.CellValue(file, tab, "A" & rowCounter) = nameToPrint 'confirmar esto
				GoExcel.CellValue("B" & rowCounter) = partModDim.Name
				GoExcel.CellValue("C" & rowCounter) = partModDim.ModelValue * ConvFactor
			Else
	   		    GoExcel.CellValue("A" & rowCounter) = nameToPrint 'confirmar esto
				GoExcel.CellValue("B" & rowCounter) = partModDim.Name
				GoExcel.CellValue("C" & rowCounter) = partModDim.ModelValue * ConvFactor
			End If
	    rowCounter = rowCounter + 1 
		Next
	    For Each partModAnno As Object In occDoc.ComponentDefinition.ModelAnnotations.ModelLeaderNotes	
	        If rowCounter = 2 Then
	            GoExcel.CellValue(file, tab, "A" & rowCounter) = nameToPrint 'confirmar esto
				GoExcel.CellValue("B" & rowCounter) = partModAnno.Name
				GoExcel.CellValue("C" & rowCounter) = partModAnno.InternalName
			Else
	   		    GoExcel.CellValue("A" & rowCounter) = nameToPrint 'confirmar esto
				GoExcel.CellValue("B" & rowCounter) = partModAnno.Name
				GoExcel.CellValue("C" & rowCounter) = partModAnno.Definition.Text.Text
			End If
	    	rowCounter = rowCounter + 1 
			Next		
	End If
End If
Next
GoExcel.Save
MsgBox("Export Done")
End Sub
