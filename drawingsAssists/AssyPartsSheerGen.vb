Sub Main()

    Dim dwgDoc As DrawingDocument = ThisDoc.Document

    '------------------------------------------------------
    ' Let the user pick a drawing view
    '------------------------------------------------------
    Dim targetView As DrawingView

    targetView = ThisApplication.CommandManager.Pick( _
        SelectionFilterEnum.kDrawingViewFilter, _
        "Select an assembly view")

    If targetView Is Nothing Then Exit Sub

    '------------------------------------------------------
    ' Get referenced document
    '------------------------------------------------------
    Dim targetDoc As Document
    targetDoc = targetView.ReferencedDocumentDescriptor.ReferencedDocument

    '------------------------------------------------------
    ' Verify it is an assembly
    '------------------------------------------------------
    If targetDoc.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then

        MessageBox.Show( _
            "The selected view does not reference an assembly.", _
            "Invalid Selection")

        Exit Sub

    End If

    Dim asmDoc As AssemblyDocument = CType(targetDoc, AssemblyDocument)
    Dim asmDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition

    '------------------------------------------------------
    ' Used to avoid duplicated sheets
    '------------------------------------------------------
    Dim processedDocs As New System.Collections.Generic.HashSet(Of String)

    '------------------------------------------------------
    ' Loop through all leaf occurrences
    '------------------------------------------------------
    For Each occ As ComponentOccurrence In asmDef.Occurrences.AllLeafOccurrences

        'Skip suppressed occurrences
        If occ.Suppressed Then Continue For

        'Skip non-normal BOM structures
        If occ.Definition.BOMStructure <> BOMStructureEnum.kNormalBOMStructure Then
            Continue For
        End If

        Dim partDoc As Document = occ.Definition.Document

        'Avoid duplicated parts
        If processedDocs.Contains(partDoc.FullFileName) Then
            Continue For
        End If

        processedDocs.Add(partDoc.FullFileName)

        '----------------------------------------------
        'Find Format
        '----------------------------------------------
		
		
		Dim sf As SheetFormat = Nothing

		For Each fmt As SheetFormat In dwgDoc.SheetFormats
		
		    If fmt.Name = "PartsFab" Then
		        sf = fmt
		        Exit For
		    End If
		
		Next
		
		If sf Is Nothing Then
		    MessageBox.Show("Sheet format not found.")
		    Exit Sub
		End If
		

        '----------------------------------------------
        'Create sheet
        '----------------------------------------------
		

        Dim sheetName As String

'        sheetName = System.IO.Path.GetFileNameWithoutExtension( _
'                        partDoc.DisplayName) _
'                        & " Fabrication dwg"
        sheetName = partDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value _
                        & " Fabrication dwg"

        Try

            dwgDoc.Sheets.AddUsingSheetFormat( _
                sf, _
                partDoc, _
                sheetName)

        Catch ex As Exception

            MessageBox.Show( _
                "Could not create sheet for:" & vbCrLf & _
                partDoc.DisplayName & vbCrLf & vbCrLf & _
                ex.Message)

        End Try

    Next

    MessageBox.Show("Finished.")

End Sub
