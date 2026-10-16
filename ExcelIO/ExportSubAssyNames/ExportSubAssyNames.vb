Sub Main()

    Dim currentDoc As AssemblyDocument = ThisDoc.Document
    Dim sFile As String = currentDoc.PropertySets.Item("Spreadsheet Document").Item("File Name").Value

    ' Column where the containing sub-assemblies start (4 = D)
    Const PARENT_START_COL As Integer = 4

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

    ' Read existing rows (column A and columns from D onward) to avoid duplicates
    Dim seen As New HashSet(Of String)
    Dim rowCounter As Integer = 2

    While CellText("A" & rowCounter) <> "" And rowCounter < 5000
        Dim parts As New List(Of String)
        parts.Add(CellText("A" & rowCounter))
        For c As Integer = PARENT_START_COL To PARENT_START_COL + 14
            Dim txt As String = CellText(ColumnLetter(c) & rowCounter)
            If txt = "" Then Exit For
            parts.Add(txt)
        Next
        seen.Add(String.Join("|", parts.ToArray()))
        rowCounter = rowCounter + 1
    End While

    ' Collect sub-assemblies recursively
    Dim results As New List(Of List(Of String))
    Dim emptyChain As New List(Of String)
    CollectSubAssemblies(currentDoc.ComponentDefinition.Occurrences, emptyChain, results, seen)

    ' Write results: name in column A, columns B and C left empty,
    ' containing sub-assemblies (nearest parent first) starting at column D
    For Each chain As List(Of String) In results
        GoExcel.CellValue("A" & rowCounter) = chain(0)
        For c As Integer = 1 To chain.Count - 1
            GoExcel.CellValue(ColumnLetter(PARENT_START_COL + c - 1) & rowCounter) = chain(c)
        Next
        rowCounter = rowCounter + 1
    Next

    GoExcel.Save
    GoExcel.Close

End Sub

Sub CollectSubAssemblies(occs As System.Collections.IEnumerable, parentChain As List(Of String), results As List(Of List(Of String)), seen As HashSet(Of String))

    For Each occ As ComponentOccurrence In occs

        If occ.Suppressed Then Continue For
        If occ.DefinitionDocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then Continue For

        Dim nextChain As New List(Of String)(parentChain)

        If occ.BOMStructure.ToString = "kNormalBOMStructure" Then

            Dim asmName As String = CleanName(occ.Definition.Document.DisplayName)

            ' Row = own name followed by containing sub-assemblies (nearest parent first)
            Dim row As New List(Of String)
            row.Add(asmName)
            row.AddRange(parentChain)

            Dim key As String = String.Join("|", row.ToArray())
            If Not seen.Contains(key) Then
                seen.Add(key)
                results.Add(row)
            End If

            ' This sub-assembly becomes the nearest parent for its children
            nextChain.Insert(0, asmName)

        End If

        CollectSubAssemblies(occ.SubOccurrences, nextChain, results, seen)

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
