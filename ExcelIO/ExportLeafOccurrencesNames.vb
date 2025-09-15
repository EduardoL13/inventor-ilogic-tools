Sub Main ()

Dim currentDoc As AssemblyDocument = ThisDoc.Document
Dim count As Integer = 0
Dim leafOccurrences As ComponentOccurrencesEnumerator = currentDoc.ComponentDefinition.Occurrences.AllLeafOccurrences

Dim listkeyStrings As New List(Of String)

Dim sFile As String = currentDoc.PropertySets.Item("Spreadsheet Document").Item("File Name").Value
Dim tab As String = currentDoc.PropertySets.Item("Worksheet Data").Item("Worksheet Name").Value	

Dim rowCounter = 2

For Each compOccurrence As ComponentOccurrence In leafOccurrences
	
    If compOccurrence.BOMStructure.ToString = "kNormalBOMStructure" And compOccurrence.Suppressed = False Then	
		Dim occDoc As PartDocument = compOccurrence.Definition.Document
		nameToCheck = compOccurrence.Name.Substring(0, compOccurrence.Name.LastIndexOf(":"))
	
		If listkeyStrings.Contains(nameToCheck) Then
		
		Else
			'listOccsNames.Add(occDoc.DisplayName.Substring(0, occDoc.DisplayName.LastIndexOf(".")))
			listkeyStrings.Add(nameToCheck)
        	If rowCounter = 2 Then
                GoExcel.CellValue(sFile, tab, "A" & rowCounter) = occDoc.DisplayName.Substring(0, occDoc.DisplayName.LastIndexOf("."))
        	Else
            	GoExcel.CellValue("A" & rowCounter) = occDoc.DisplayName.Substring(0, occDoc.DisplayName.LastIndexOf("."))
       		End If
	        rowCounter = rowCounter + 1  		
	
		End If
	End If
Next

GoExcel.Save

End Sub



