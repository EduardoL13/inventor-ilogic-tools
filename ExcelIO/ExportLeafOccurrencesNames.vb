Sub Main ()

	Dim currentDoc As AssemblyDocument = ThisDoc.Document
	Dim count As Integer = 0
	Dim leafOccurrences As ComponentOccurrencesEnumerator = currentDoc.ComponentDefinition.Occurrences.AllLeafOccurrences

	Dim listkeyStrings As New List(Of String)

	Dim sFile As String = currentDoc.PropertySets.Item("Spreadsheet Document").Item("File Name").Value
	Dim tab As String = currentDoc.PropertySets.Item("Worksheet Data").Item("Worksheet Name").Value	

	Dim rowCounter = 2
	
	Dim excelNames As New List(Of String)

	If GoExcel.CellValue(sFile, tab, "A2") <> "" Then

        Dim lastRow As Integer = findLastDataRow(sFile, tab)

    	For i As Integer = 2 To lastRow - 1
        	excelNames.Add(GoExcel.CellValue(sFile, tab, "A" & i).ToString())
    	Next

    	rowCounter = lastRow

	Else

    	rowCounter = 2

	End If
	
	For Each compOccurrence As ComponentOccurrence In leafOccurrences
	
    	If compOccurrence.BOMStructure.ToString = "kNormalBOMStructure" And compOccurrence.Suppressed = False Then	
			Dim occDoc As PartDocument = compOccurrence.Definition.Document
			nameToCheck = compOccurrence.Name.Substring(0, compOccurrence.Name.LastIndexOf(":"))
	
		If listkeyStrings.Contains(nameToCheck) Then

		Else

    	listkeyStrings.Add(nameToCheck)

    	Dim partName As String = occDoc.DisplayName.Substring(0, occDoc.DisplayName.LastIndexOf("."))

    	'Solo escribir si no existe ya en Excel
    	If Not excelNames.Contains(partName) Then

        	If rowCounter = 2 Then
            	GoExcel.CellValue(sFile, tab, "A" & rowCounter) = partName
        	Else
            	GoExcel.CellValue("A" & rowCounter) = partName
        	End If

        	excelNames.Add(partName)
        	rowCounter = rowCounter + 1

    	End If

End If
	End If
Next

GoExcel.Save

End Sub

Function findLastDataRow(file As String, tab As String)

    Dim cellVal As Object
    Dim range As Integer = 1000
    Dim lastDataRow As Integer = 0

    For rowNum As Integer = 1 To range

        cellVal = GoExcel.CellValue(file, tab, "A" & rowNum)

        If cellVal Is Nothing Then
            If rowNum > range Then Exit For
        Else
            lastDataRow = rowNum + 1
        End If

    Next

    Return lastDataRow

End Function


