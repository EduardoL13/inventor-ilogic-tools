Sub Main()

    Dim currentDoc As AssemblyDocument = ThisDoc.Document
    Dim oAsmDef As AssemblyComponentDefinition = currentDoc.ComponentDefinition

    Dim missingAnnotations As New List(Of String)

    'Controla qué documentos ya fueron revisados
    Dim processedParts As New HashSet(Of String)

    For Each leafOcc As ComponentOccurrence In oAsmDef.Occurrences.AllLeafOccurrences

        'Ignorar ocurrencias suprimidas
        If leafOcc.Suppressed Then Continue For

        'Ignorar componentes cuya BOM Structure no sea Normal
        If leafOcc.BOMStructure <> BOMStructureEnum.kNormalBOMStructure Then Continue For

        Dim partDoc As PartDocument = leafOcc.Definition.Document

        'Ignorar piezas que ya fueron revisadas
        If Not processedParts.Add(partDoc.FullFileName) Then Continue For

        If Not HasModelDimensions(partDoc) Then
            missingAnnotations.Add(leafOcc.Name)
        End If

    Next

    If missingAnnotations.Count = 0 Then

        MsgBox("All valid parts contain model annotations.")

    Else

        Dim report As String = ""

        For Each partName As String In missingAnnotations
            report &= partName & vbCrLf
        Next

        MsgBox("The following parts do not contain model annotations:" _
            & vbCrLf & vbCrLf & report)

    End If

End Sub


Function HasModelDimensions(partDoc As PartDocument) As Boolean

    Dim compDef As PartComponentDefinition = partDoc.ComponentDefinition

    '-----------------------------
    ' Sheet Metal Part
    '-----------------------------
    If TypeOf compDef Is SheetMetalComponentDefinition Then

        Dim smDef As SheetMetalComponentDefinition = compDef

        If Not smDef.HasFlatPattern Then
            Return False
        End If

        If smDef.FlatPattern.ModelAnnotations.ModelDimensions.Count > 0 Then
            Return True
        Else
            Return False
        End If

    End If

    '-----------------------------
    ' Standard Part
    '-----------------------------
    If compDef.ModelAnnotations.ModelDimensions.Count > 0 Then
        Return True
    Else
        Return False
    End If

End Function
