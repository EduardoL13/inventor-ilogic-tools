Sub Main()

    Dim currentDoc As AssemblyDocument = ThisDoc.Document
    Dim sFile As String = currentDoc.PropertySets.Item("Spreadsheet Document").Item("File Name").Value

    ' Column where the immediate parent is written (4 = D)
    Const PARENT_COL As Integer = 4

    ' Ask for the worksheet name
    Dim tab As String = InputBox("Enter the worksheet name:", "Sub-assemblies export", "")
    If tab.Trim() = "" Then Exit Sub
    tab = tab.Trim()

    ' Open the workbook on the requested worksheet (fails if it does not exist)
    Try
        GoExcel.Open(sFile, tab)
        Dim testValue As Object = GoExcel.CellValue("A1")
    Catch ex As Exception
        MessageBox.Show("Worksheet '" & tab & "' could not be opened in:" & vbCrLf & sFile & vbCrLf & vbCrLf & ex.Message, "Error")
        Exit Sub
    End Try

    ' Read existing rows (column A and parent column) to avoid duplicates
    Dim seen As New HashSet(Of String)
    Dim rowCounter As Integer = 2

    While CellText("A" & rowCounter) <> "" And rowCounter < 5000
        Dim key As String = CellText("A" & rowCounter) & "|" & CellText(ColumnLetter(PARENT_COL) & rowCounter)
        seen.Add(key)
        rowCounter = rowCounter + 1
    End While

    ' Collect sub-assemblies recursively
    Dim results As New List(Of String())
    CollectSubAssemblies(currentDoc.ComponentDefinition.Occurrences, "", results, seen)

    ' Write results: name in column A, columns B and C left empty,
    ' immediate parent (if any) in column D
    For Each row As String() In results
        GoExcel.CellValue("A" & rowCounter) = row(0)
        If row(1) <> "" Then
            GoExcel.CellValue(ColumnLetter(PARENT_COL) & rowCounter) = row(1)
        End If
        rowCounter = rowCounter + 1
    Next

    GoExcel.Save
    GoExcel.Close

End Sub

Sub CollectSubAssemblies(occs As System.Collections.IEnumerable, parentName As String, results As List(Of String()), seen As HashSet(Of String))

    For Each occ As ComponentOccurrence In occs

        If occ.Suppressed Then Continue For
        If occ.DefinitionDocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then Continue For

        Dim nextParent As String = parentName

        If occ.BOMStructure.ToString = "kNormalBOMStructure" Then

            Dim asmName As String = CleanName(occ.Definition.Document.DisplayName)

            ' Row = own name + immediate parent (empty if it sits directly under the top level)
            Dim key As String = asmName & "|" & parentName
            If Not seen.Contains(key) Then
                seen.Add(key)
                results.Add(New String() {asmName, parentName})
            End If

            ' This sub-assembly becomes the immediate parent for its children
            nextParent = asmName

        End If

        CollectSubAssemblies(occ.SubOccurrences, nextParent, results, seen)

    Next

End Sub

Function CleanName(displayName As String) As String
    Dim idx As Integer = displayName.LastIndexOf(".")
    If idx > 0 Then Return displayName.Substring(0, idx)
    Return displayName
End Function

Function CellText(address As String) As String
    Dim v As Object = GoExcel.CellValue(address)
    If v Is Nothing Then Return ""
    Return v.ToString().Trim()
End Function

Function ColumnLetter(index As Integer) As String
    Dim result As String = ""
    While index > 0
        Dim remainder As Integer = (index - 1) Mod 26
        result = Chr(65 + remainder) & result
        index = (index - 1) \ 26
    End While
    Return result
End Function
