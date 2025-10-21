Sub Main ()

Dim currentDoc As AssemblyDocument = ThisDoc.Document

Dim sFile As String
Dim tabDS As String

sFile = currentDoc.PropertySets.Item("Spreadsheet Document").Item("File Name").Value
tabDS = currentDoc.PropertySets.Item("Worksheet Data").Item("Worksheet Name").Value	
	
Dim lastDataValue As Integer = findLastDataRow(sFile, tabDS) '4 'Definición de Range 

setPropertiesAndMat(currentDoc, sFile, tabDS, lastDataValue) 
currentDoc.Update2
MsgBox("Properties updated successfully")
	
End Sub	


Function findLastDataRow(file As String, tab As String)
	Dim cellVal As Object 
	Dim range As Integer = 1000
	Dim lastDataRow As Integer = 0
    For rowNum As Integer = 1 To range
		cellVal = GoExcel.CellValue(file, tab, "A" & rowNum)
	    If cellVal Is Nothing  Then 
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
	
	Dim listkeyStrings As New List(Of String)

    For Each compOcc As ComponentOccurrence In assembComp.ComponentDefinition.Occurrences.AllLeafOccurrences
		
	    If compOcc.Suppressed Or TypeOf compOcc.Definition.Document IsNot PartDocument Then 
	        'MsgBox(compOcc.Name)
	    Else	
		    Dim occDoc As PartDocument = compOcc.Definition.Document ' Acceso al part document de la instancia
	        nameToCompare = occDoc.DisplayName.Substring(0, occDoc.DisplayName.LastIndexOf("."))
		   
		    If listkeyStrings.Contains(nameToCompare) Then
				
		    Else
		        'MsgBox(nameToCheck)
			     listkeyStrings.Add(nameToCompare)
		
	             For rowCounter = 2 To lastValue
				
		             Dim nameOccDS As String = GoExcel.CellValue(file, tab, "A" & rowCounter) 
				     'Dim occDoc As PartDocument = compOcc.Definition.Document ' Acceso al part document de la instancia
					 Dim occProps As PropertySet = occDoc.PropertySets.Item("Design Tracking Properties")
				
					 Dim oCurrentScope As MemberEditScopeEnum
					 Dim oFactoryDoc As PartDocument
				
				
				     If occDoc.ComponentDefinition.IsModelStateMember Then
					
					     oFactoryDoc = occDoc.ComponentDefinition.FactoryDocument
					     oCurrentScope = oFactoryDoc.ComponentDefinition.ModelStates.MemberEditScope
		
				         If oCurrentScope = MemberEditScopeEnum.kEditActiveMember Then
					         oFactoryDoc.ComponentDefinition.ModelStates.MemberEditScope = MemberEditScopeEnum.kEditAllMembers
				    	 End If
				    	 occDoc = oFactoryDoc
					 End If				
				
					'nameToCompare = occDoc.DisplayName.Substring(0, occDoc.DisplayName.LastIndexOf("."))
				
				
			        'If compOcc.Name = nameOccDS Then
                 	 If nameToCompare = nameOccDS Then 
				         MsgBox(nameToCompare)
                    	 propsAssigner(nameToCompare & ".ipt", file, tab, rowCounter, colPartNo, colStockNumber, colDescription, colMaterial)    'compOcc.Name
					
	                	 If oFactoryDoc IsNot Nothing Then
						     oFactoryDoc.ComponentDefinition.ModelStates.MemberEditScope = oCurrentScope
						 End If
	              
			         End If
			     Next 
		     End If
	    End If
    Next
    			
End Sub

Sub propsAssigner(compName As String, doc As String, tab As String, row As Integer, partNo As String, stockNumber As String, description As String, material As String)
	
    iProperties.Expression(compName, "Project", "Part Number") = GoExcel.CellValue(doc, tab, partNo & row)
	iProperties.Expression(compName, "Project", "Stock Number") = GoExcel.CellValue(doc, tab, stockNumber & row)
	iProperties.Expression(compName, "Project", "Description") = GoExcel.CellValue(doc, tab, description & row)
	iProperties.Material(compName) = GoExcel.CellValue(doc, tab, material & row)
	
End Sub
