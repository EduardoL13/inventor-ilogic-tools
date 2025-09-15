Sub Main ()

Dim currentDoc As AssemblyDocument = ThisDoc.Document

Dim sFile As String
Dim tabDS As String

' Properties assigned by the macro "SetupSpreadsheetData"
sFile = currentDoc.PropertySets.Item("Spreadsheet Document").Item("File Name").Value
tabDS = currentDoc.PropertySets.Item("Worksheet Data").Item("Worksheet Name").Value	
	
Dim lastDataValue As Integer = findLastDataRow(sFile, tabDS) '4 'Definición de Range 

setPropertiesAndMat(currentDoc, sFile, tabDS, lastDataValue) 
	
End Sub	


Function findLastDataRow(file As String, tab As String)
' Returns the value of the last populated cell in the Column A.
	Dim cellVal As Object 
	Dim range As Integer = 1000
	Dim lastDataRow As Integer = 0
	
    For rowNum As Integer = 1 To range
		
		cellVal = GoExcel.CellValue(file, tab, "A" & rowNum)
	    If cellVal Is Nothing  Then '(cellVal Is Nothing OrElse String.IsNullOrEmpty(cellVal.ToString()))
			If (rowNum > range) Then Exit For
	    Else
			lastDataRow = rowNum + 1
		End If
		
	Next
	Return lastDataRow
	
End Function


Sub setPropertiesAndMat(assembComp As AssemblyDocument, file As String, tab As String, lastValue As Integer) 

For Each compOcc As ComponentOccurrence In assembComp.ComponentDefinition.Occurrences.AllLeafOccurrences
    If compOcc.Suppressed Then
	
	Else
	    For rowCounter=2 To lastValue
		Dim nameOccDS As String = GoExcel.CellValue(file, tab, "A" & rowCounter) 
		'MsgBox(compOcc.Name)
		
		
		'Make sure to put here below the corresponding column letter to the property you want to assign
		colPartNo = "M" 'Part Number
		colStockNumber = "N" 'Stock Number
		colDescription = "D" 'Description
		colMaterial = "E" 'Material 
		
		If compOcc.Name = nameOccDS Then
			
		    ' Assign Properties (make sure that assigned column matches the desired property)
        	iProperties.Expression(compOcc.Name, "Project", "Part Number") = GoExcel.CellValue(file, tab, colPartNo & rowCounter)
			iProperties.Expression(compOcc.Name, "Project", "Stock Number") = GoExcel.CellValue(file, tab, colStockNumber & rowCounter)
	    	iProperties.Expression(compOcc.Name, "Project", "Description") = GoExcel.CellValue(file, tab, colDescription & rowCounter)
	    	iProperties.MaterialOfComponent(compOcc.Name) = GoExcel.CellValue(file, tab, colMaterial & rowCounter) ' 
			
		End If
	Next
	End If
	Next

End Sub
