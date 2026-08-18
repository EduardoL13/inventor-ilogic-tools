Sub Main()

    'NOTA: Las dimensiones de los valores están dados en cm por defecto

    Dim esteDoc As AssemblyDocument = ThisDoc.Document

    Dim leafOccurrences As ComponentOccurrencesEnumerator = _
        esteDoc.ComponentDefinition.Occurrences.AllLeafOccurrences

    Dim file As String = _
        esteDoc.PropertySets.Item("Spreadsheet Document").Item("File Name").Value

    Dim tab As String = _
        InputBox("Enter the excel worksheet name", "Worksheet to Export")

    If tab = "" Then Exit Sub

    '-----------------------------
    'Excel Columns
    '-----------------------------

    Dim partTypeColumn As String = "A"

    Dim partLengthColumn As String = "B"

    Dim partQtyColumn As String = "C"

    Dim partIDColumn As String = "D"

    Dim partDescriptionColumn As String = "F"

    '-----------------------------
    'Locate Part Section
    '-----------------------------

	Dim startRow As Integer
	Dim nextAvailableRow As Integer

	Dim rowDictionary As Dictionary(Of String, Integer)

	rowDictionary = BuildExcelIndex( _
                    file, _
                    tab, _
                    partTypeColumn, _
                    partIDColumn, _
                    startRow, _
                    nextAvailableRow)

    '-----------------------------
    'Avoid processing duplicate parts
    '-----------------------------

    Dim processedParts As New HashSet(Of String)

    '-----------------------------
    'Process Assembly
    '-----------------------------

    For Each compOccurrence As ComponentOccurrence In leafOccurrences

        If compOccurrence.Suppressed Then Continue For

        If compOccurrence.BOMStructure <> _
            BOMStructureEnum.kNormalBOMStructure Then Continue For

        If TypeOf compOccurrence.Definition.Document IsNot PartDocument Then Continue For

        Dim occDoc As PartDocument = compOccurrence.Definition.Document

        If Not processedParts.Add(occDoc.FullFileName) Then Continue For

        If occDoc.ComponentDefinition.Type.ToString = _
            "kSheetMetalComponentDefinitionObject" Then Continue For

        Dim partID As String

        partID = occDoc.PropertySets.Item( _
            "Design Tracking Properties").Item( _
            "Stock Number").Value

        Dim currentRow As Integer

        If rowDictionary.ContainsKey(partID) Then

            currentRow = rowDictionary(partID)

        Else

            currentRow = nextAvailableRow

            rowDictionary.Add(partID, currentRow)

            nextAvailableRow += 1

        End If

        propsPrinter( _
            occDoc, _
            file, _
            tab, _
            partTypeColumn, _
            partIDColumn, _
            partQtyColumn, _
            partDescriptionColumn, _
            partLengthColumn, _
            currentRow)

    Next

    GoExcel.Save()

    MsgBox("Export Done")

End Sub



Sub propsPrinter( _
    currentPart As PartDocument, _
    file As String, _
    tab As String, _
    partTypeColumn As String, _
    partIDColumn As String, _
    partQtyColumn As String, _
    partDescriptionColumn As String, _
    partLengthColumn As String, _
    currentRow As Integer)

    Dim ConvFactor As Double = 1 / 2.54

    Try

        'Type
        GoExcel.CellValue(file, tab, partTypeColumn & currentRow) = "P"

        'Part ID
        GoExcel.CellValue(file, tab, partIDColumn & currentRow) = _
            currentPart.PropertySets.Item( _
            "Design Tracking Properties").Item( _
            "Stock Number").Value

        'Quantity
        GoExcel.CellValue(file, tab, partQtyColumn & currentRow) = _
            Single.Parse(currentPart.PropertySets.Item( _
            "Design Tracking Properties").Item( _
            "Cost Center").Value)

        'Description
        GoExcel.CellValue(file, tab, partDescriptionColumn & currentRow) = _
            currentPart.PropertySets.Item( _
            "Design Tracking Properties").Item( _
            "Description").Value

        'Length
        GoExcel.CellValue(file, tab, partLengthColumn & currentRow) = _
            currentPart.ComponentDefinition.ModelAnnotations. _
            ModelDimensions.Item("length").ModelValue * ConvFactor

    Catch

        MsgBox(currentPart.DisplayName & _
               " has missing properties or properties that are not in the valid format")

    End Try

End Sub

Function BuildExcelIndex( _
    file As String, _
    tab As String, _
    partTypeColumn As String, _
    partIDColumn As String, _
    ByRef startRow As Integer, _
    ByRef nextAvailableRow As Integer) _
    As Dictionary(Of String, Integer)

    Dim dict As New Dictionary(Of String, Integer)

    Dim searchRange As Integer = 1000

    Dim firstEmptyRow As Integer = 0
    Dim firstPartRow As Integer = 0
    Dim lastPartRow As Integer = 0

    For rowNum As Integer = 2 To searchRange

        Dim typeValue As Object

        typeValue = GoExcel.CellValue(file, tab, partTypeColumn & rowNum)

        'Guardar la primera fila vacía
        If typeValue Is Nothing Then

            If firstEmptyRow = 0 Then
                firstEmptyRow = rowNum
            End If

            Continue For

        End If

        If typeValue.ToString.Trim.ToUpper <> "P" Then Continue For

        'Primera fila del bloque
        If firstPartRow = 0 Then
            firstPartRow = rowNum
        End If

        lastPartRow = rowNum

        Dim idValue As Object

        idValue = GoExcel.CellValue(file, tab, partIDColumn & rowNum)

        If idValue Is Nothing Then Continue For

        Dim partID As String

        partID = idValue.ToString.Trim

        If Not dict.ContainsKey(partID) Then

            dict.Add(partID, rowNum)

        End If

    Next

    If firstPartRow = 0 Then

        'No existe todavía bloque de Parts
        startRow = firstEmptyRow
        nextAvailableRow = firstEmptyRow

    Else

        'Ya existe un bloque de Parts
        startRow = firstPartRow
        nextAvailableRow = lastPartRow + 1

    End If

    Return dict

End Function
