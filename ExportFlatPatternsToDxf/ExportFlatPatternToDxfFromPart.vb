Sub Main ()

    Dim currentDoc As PartDocument = ThisDoc.Document
	
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
	
            
            ' Check for sheet metal
            If currentDoc.ComponentDefinition.Type = kSheetMetalComponentDefinitionObject Then

                Dim smCompDef As SheetMetalComponentDefinition = currentDoc.ComponentDefinition
                
                ' verifies if there is flat pattern and creates one if there is not
                    If Not smCompDef.HasFlatPattern Then
                        smCompDef.Unfold()
                    End If
                
                ' Gets flat pattern
                    Dim flatPattern As FlatPattern = smCompDef.FlatPattern
               
	            ' Gets name for dxf file

					Try
			            desiredDisplayName = currentDoc.PropertySets.Item("Design Tracking Properties").Item("Stock Number").Value
			           	'MsgBox(desiredDisplayName)
					   
					Catch
					    MsgBox("Current part " & nameToCheck & "has no Stock number/Part ID assigned")
				        Exit Sub
						
					End Try

			    'MsgBox(desiredDisplayName)
                ' Creates file name
                    Dim fileName As String = System.IO.Path.Combine(targetLocation, desiredDisplayName & ".dxf") 
                'MsgBox(fileName & " has been added")
				
                ' Export to DXF
                flatPattern.DataIO.WriteDataToFile("FLAT PATTERN DXF?AcadVersion=2018", fileName)
            Else
			    MsgBox("No Existing Flat Pattern within this part")
				Exit Sub
				
            End If
			
    MsgBox("Done")
End Sub
