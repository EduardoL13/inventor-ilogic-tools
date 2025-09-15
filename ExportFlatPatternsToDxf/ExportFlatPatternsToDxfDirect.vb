Sub Main ()

    Dim currentDoc As AssemblyDocument = ThisDoc.Document
    Dim leafOccurrences As ComponentOccurrencesEnumerator = currentDoc.ComponentDefinition.Occurrences.AllLeafOccurrences
	
	Dim oDlg As New System.Windows.Forms.FolderBrowserDialog
	Dim dlgResult As New System.Windows.Forms.DialogResult
	
    With oDlg
	
    .ShowNewFolderButton = True
    .InitialDirectory = "C:\Temp"

    If .ShowDialog = dlgResult.Cancel
	    Exit Sub
	End If
	
    End With
	
	targetLocation = oDlg.SelectedPath
	
	Dim listkeyStrings As New List(Of String)
	
    For Each compOccurrence As ComponentOccurrence In leafOccurrences
		
		If compOccurrence.Suppressed Or TypeOf compOccurrence.Definition.Document IsNot PartDocument Then
        
		Else

            Dim occDoc As PartDocument = compOccurrence.Definition.Document
            
            ' Check for sheet metal
            If occDoc.ComponentDefinition.Type = kSheetMetalComponentDefinitionObject Then
		    
			' Check if part is a replica of an existing one    
		        nameToCheck = compOccurrence.Name.Substring(0, compOccurrence.Name.LastIndexOf(":"))  'trimName(compOccurrence.Name, 2) ' Removes the 2 last chars of the occ name, for example: ":1"
				
		        If listkeyStrings.Contains(nameToCheck) Then
		
		        Else
				   'MsgBox(nameToCheck)
			       listkeyStrings.Add(nameToCheck)
		    
                    Dim smCompDef As SheetMetalComponentDefinition = occDoc.ComponentDefinition
                
                ' verifies if there is flat pattern and creates one if there is not
                    If Not smCompDef.HasFlatPattern Then
                        smCompDef.Unfold()
                    End If
                
                ' Gets flat pattern
                    Dim flatPattern As FlatPattern = smCompDef.FlatPattern
               
	            ' Gets name for dxf file
			        Dim desiredDisplayName As String = occDoc.DisplayName.Substring(0, occDoc.DisplayName.LastIndexOf(".")) 'removes ".ipt"
			    'MsgBox(desiredDisplayName)
				
                ' Creates file name
                    Dim fileName As String = System.IO.Path.Combine(targetLocation, desiredDisplayName & ".dxf") 
                'MsgBox(fileName & " has been added")
				
                ' Export to DXF
                flatPattern.DataIO.WriteDataToFile("FLAT PATTERN DXF?AcadVersion=2018", fileName)
                
                End If
            End If
        End If
        Next
    MsgBox("All dxf files have been added to the chosen location")
End Sub
