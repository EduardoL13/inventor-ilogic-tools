Sub Main ()
'NOTA: Las dimensiones de los valores está dados en cm por defecto
Dim esteDoc As AssemblyDocument = ThisDoc.Document
Dim count As Integer = 0
Dim leafOccurrences As ComponentOccurrencesEnumerator = esteDoc.ComponentDefinition.Occurrences.AllLeafOccurrences


Dim file As String = esteDoc.PropertySets.Item("Spreadsheet Document").Item("File Name").Value
'Dim tab As String = "CAD_output_prof"
Dim tab As String = InputBox("Enter the excel worksheet name", "Worksheet to Export")


Dim ConvFactor As Double = 1/2.54 'Factor de conversión de cm a in
Dim rowCounter = 2
For Each compOccurrence As ComponentOccurrence In leafOccurrences
	
	Dim occDoc As PartDocument = compOccurrence.Definition.Document
	'nameToPrint = occDoc.DisplayName.Substring(0, occDoc.DisplayName.LastIndexOf("."))
	nameToPrint = occDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value


    If compOccurrence.Visible = False Or TypeOf compOccurrence.Definition.Document IsNot PartDocument Or compOccurrence.Suppressed Then
	
	Else
	    For Each partModNote As ModelGeneralNote In occDoc.ComponentDefinition.ModelAnnotations.ModelGeneralNotes
	        If rowCounter = 2 Then
	            GoExcel.CellValue(file, tab, "E" & rowCounter) = nameToPrint 'confirmar esto
				GoExcel.CellValue("F" & rowCounter) = partModNote.Name
				GoExcel.CellValue("G" & rowCounter) = partModNote.Definition.Text.Text
			Else
	   		    GoExcel.CellValue("E" & rowCounter) = nameToPrint 'confirmar esto
				GoExcel.CellValue("F" & rowCounter) = partModNote.Name
				GoExcel.CellValue("G" & rowCounter) = partModNote.Definition.Text.Text
			End If
	    rowCounter = rowCounter + 1 
		Next
	End If
Next
GoExcel.Save
MsgBox("Export Done")
End Sub
