Sub Main()

    Dim currentDoc As AssemblyDocument = ThisDoc.Document

    Dim sFile As String
    Dim tabDS As String

    Try

        sFile = currentDoc.PropertySets.Item("Spreadsheet Document").Item("File Name").Value
        tabDS = currentDoc.PropertySets.Item("Worksheet Data").Item("Worksheet Name").Value

    Catch

        MsgBox("Please set up the file and worksheet to proceed")
        Exit Sub

    End Try

    Dim colName As String = "A"
    Dim rowStart As Integer = 2

    Dim lastDataRow As Integer = FindLastDataRow(sFile, tabDS)

    'Guardar todos los nombres existentes en Excel
    Dim existingNames As New HashSet(Of String)

    For rowNum As Integer = rowStart To lastDataRow

        Dim cellVal As Object = GoExcel.CellValue(sFile, tabDS, colName & rowNum)

        If Not cellVal Is Nothing Then
            existingNames.Add(cellVal.ToString.Trim)
			'MsgBox(cellVal.ToString.Trim)
        End If

    Next

    Dim addedCount As Integer = 0
    Dim addedNames As String = ""
	
	Dim listOccsNames As New List(Of String)
	Dim listkeyStrings As New List(Of String)

    Dim oAsmDef As AssemblyComponentDefinition = currentDoc.ComponentDefinition

    For Each compOccurrence As ComponentOccurrence In oAsmDef.Occurrences.AllLeafOccurrences
        If compOccurrence.BOMStructure.ToString = "kNormalBOMStructure" And compOccurrence.Suppressed = False Then	
		    Dim occDoc As PartDocument = compOccurrence.Definition.Document
     		nameToCheck = compOccurrence.Name.Substring(0, compOccurrence.Name.LastIndexOf(":"))
      		If listkeyStrings.Contains(nameToCheck) Then
		
     		Else
            	'listOccsNames.Add(compOccurrence.Name)
				occName = occDoc.DisplayName.Substring(0, occDoc.DisplayName.LastIndexOf("."))
			    listOccsNames.Add(occName) 'Add(occDoc.DisplayName.Substring(0, occDoc.DisplayName.LastIndexOf(".")))
				listkeyStrings.Add(nameToCheck)
		
            	If Not existingNames.Contains(occName) Then

                	lastDataRow += 1
                    'Msgbox(lastDataRow)
	                GoExcel.CellValue(sFile, tabDS, colName & lastDataRow) = occName
	                
	            	existingNames.Add(occName)
	
	            	addedCount += 1
	            	addedNames &= occName & vbCrLf

            	End If
			End If
        End If
    Next
    GoExcel.Save
    If addedCount > 0 Then

        MsgBox(addedCount & " names added:" & vbCrLf & vbCrLf & addedNames)

    Else

        MsgBox("No new names were found.")

    End If

End Sub


Function FindLastDataRow(file As String, tab As String) As Integer

    Dim range As Integer = 1000
    Dim lastDataRow As Integer = 1

    For rowNum As Integer = 1 To range

        Dim cellVal As Object = GoExcel.CellValue(file, tab, "A" & rowNum)

        If Not cellVal Is Nothing Then
            lastDataRow = rowNum
        End If

    Next

    Return lastDataRow
	
End Function
