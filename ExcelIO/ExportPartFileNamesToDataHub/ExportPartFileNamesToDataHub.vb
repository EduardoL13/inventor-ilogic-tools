Sub Main ()

Dim currentDoc As AssemblyDocument = ThisDoc.Document
Dim count As Integer = 0
Dim leafOccurrences As ComponentOccurrencesEnumerator = currentDoc.ComponentDefinition.Occurrences.AllLeafOccurrences

Dim numOccs As Integer = leafOccurrences.Count ' 
'Dim listOccsNames(numOccs) As String  ' Definición de la variable que se le va a asignar al multiparameter list

Dim listOccsNames As New List(Of String)
Dim listkeyStrings As New List(Of String)


For Each compOccurrence As ComponentOccurrence In leafOccurrences
    If compOccurrence.BOMStructure.ToString = "kNormalBOMStructure" And compOccurrence.Suppressed = False Then	
		Dim occDoc As PartDocument = compOccurrence.Definition.Document
		nameToCheck = compOccurrence.Name.Substring(0, compOccurrence.Name.LastIndexOf(":"))
		If listkeyStrings.Contains(nameToCheck) Then
		
		Else
            'listOccsNames.Add(compOccurrence.Name)
			listOccsNames.Add(occDoc.DisplayName.Substring(0, occDoc.DisplayName.LastIndexOf(".")))
			listkeyStrings.Add(nameToCheck)
		End If
	End If
Next

Dim sFile As String
Dim tabDS As String

sFile = currentDoc.PropertySets.Item("Spreadsheet Document").Item("File Name").Value
tabDS = currentDoc.PropertySets.Item("Worksheet Data").Item("Worksheet Name").Value	

Dim file As String = sFile
Dim tab As String = tabDS
Dim rowCounter = 2

For Each name As String In listOccsNames
        If rowCounter = 2 Then
            GoExcel.CellValue(file, tab, "A" & rowCounter) = name
        Else
            GoExcel.CellValue("A" & rowCounter) = name
        End If
	    rowCounter = rowCounter + 1   
Next

GoExcel.Save
MsgBox("Done")

End Sub
