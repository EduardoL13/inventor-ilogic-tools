Sub Main ()

Dim currentDoc As AssemblyDocument = ThisDoc.Document

Dim sFile As String
Dim tabDS As String

sFile = currentDoc.PropertySets.Item("Spreadsheet Document").Item("File Name").Value
tabDS = currentDoc.PropertySets.Item("Worksheet Data").Item("Worksheet Name").Value	
	
Dim lastDataValue As Integer = findLastDataRow(sFile, tabDS) '4 'Definición de Range 

'-----
setPropertiesAndMat(currentDoc, sFile, tabDS, lastDataValue) 
currentDoc.Update2
Msgbox("Properties updated successfully")
	
End Sub	


Function findLastDataRow(file As String, tab As String)
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

	colPartNo = "M" 'Part Number
    colStockNumber = "N" 'Stock Number
    colDescription = "D" 'Description
	colMaterial = "E" 'Material 

    For Each compOcc As ComponentOccurrence In assembComp.ComponentDefinition.Occurrences.AllLeafOccurrences
        If compOcc.Suppressed Then
	
	    Else
			
			
	        For rowCounter = 2 To lastValue
				
		        Dim nameOccDS As String = GoExcel.CellValue(file, tab, "A" & rowCounter) 
				Dim occDoc As PartDocument = compOcc.Definition.Document ' Acceso al part document de la instancia
				Dim occProps As PropertySet = occDoc.PropertySets.Item("Design Tracking Properties")
				
		        If compOcc.Name = nameOccDS Then

                    propsAssigner(compOcc.Name, file, tab, rowCounter, colPartNo, colStockNumber, colDescription, colMaterial)    
					'Dim occProps As PropertySet = occDoc.PropertySets.Item("Design Tracking Properties")
					'Msgbox(occProps.Item("Categories").Value)
					'occDoc.ComponentDefinition.Material = ThisApplication.ActiveDocument.Materials.Item(occProps.Item("Categories").Value)
					occDoc.Update2
					
                    Dim compModelStates As ModelStates = occDoc.ComponentDefinition.ModelStates
					
					
					stockNoProp = occProps.Item("Stock Number").Value
					descriptionProp = occProps.Item("Description").Value
					partNoProp = occProps.Item("Part Number").Value
					materialProp= occProps.Item("Material").Value
					MsgBox(materialProp)
					
                    If compModelStates.Count <> 1 Then
						
						'MsgBox(compModelStates.Count)
 
                        Dim openDoc As PartDocument = ThisApplication.Documents.Open(occDoc.FullDocumentName,False)

                        For Each modState As ModelState In compModelStates '- 1
							
                            If modState.Name.Contains("sym") Then ' Verifica si es una parte simétrica (igual pero con otro posicionamiento)
								openDoc.ComponentDefinition.ModelStates(modState.Name).Activate
								'MsgBox(modState.Name)
								openDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value = partNoProp
								openDoc.PropertySets.Item("Design Tracking Properties").Item("Stock Number").Value = stockNoProp
								openDoc.PropertySets.Item("Design Tracking Properties").Item("Description").Value = descriptionProp
								openDoc.ComponentDefinition.Material = ThisApplication.ActiveDocument.Materials.Item(materialProp)
								
							End If
							
						occDoc.Update2
						occDoc.Save2
						
						Next
						
						openDoc.Update2
						openDoc.Save2
					    openDoc.Close
						
					Else
						
						propsAssigner(compOcc.Name, file, tab, rowCounter, colPartNo, colStockNumber, colDescription, colMaterial)
                        'occDoc.ComponentDefinition.Material = ThisApplication.ActiveDocument.Materials.Item(occProps.Item("Categories").Value)
						'occDoc.ComponentDefinition.Material = ThisApplication.ActiveDocument.Materials.Item(occProps.Item(materialProp).Value)
						
                    End If
	 					
			    End If
			Next 
		 End If
    Next
    			
End Sub

Sub propsAssigner(compName As String, doc As String, tab As String, row As Integer, partNo As String, stockNumber As String, description As String, material As String)
	
    iProperties.Expression(compName, "Project", "Part Number") = GoExcel.CellValue(doc, tab, partNo & row)
	iProperties.Expression(compName, "Project", "Stock Number") = GoExcel.CellValue(doc, tab, stockNumber & row)
	iProperties.Expression(compName, "Project", "Description") = GoExcel.CellValue(doc, tab, description & row)
	'iProperties.Expression(compName, "Project", "Categories") = GoExcel.CellValue(doc, tab, material & row)
	iProperties.MaterialOfComponent(compName) = GoExcel.CellValue(doc, tab, material & row)
	
End Sub
