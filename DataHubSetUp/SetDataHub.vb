Sub SetupSpreadsheet()

    Dim currentDoc As Document = ThisApplication.ActiveDocument

    Dim ssData As PropertySet

    '--------------------------------------------------------
    ' Set Spreadsheet Document
    '--------------------------------------------------------

    ' Browse Spreadsheet File
    Dim docNameString As String = BrowseFile()

    ' User cancelled
    If String.IsNullOrWhiteSpace(docNameString) Then
        Exit Sub
    End If

    ' Create PropertySet if it doesn't exist
    If Not currentDoc.PropertySets.PropertySetExists("Spreadsheet Document") Then

        ssData = currentDoc.PropertySets.Add("Spreadsheet Document")

        Dim propDocName As Inventor.Property =
            ssData.Add(docNameString, "File Name")

        MsgBox("File Name: " & propDocName.Value & " has been added")

    Else

        ssData = currentDoc.PropertySets.Item("Spreadsheet Document")

        ssData.Item("File Name").Value = docNameString

        MsgBox("New File Name: " & ssData.Item("File Name").Value)

    End If


    '--------------------------------------------------------
    ' Set Worksheet
    '--------------------------------------------------------

    Dim worksheetData As PropertySet

    Dim tabNameString As String =
        InputBox("Enter worksheet tab name")

    ' User cancelled
    If String.IsNullOrWhiteSpace(tabNameString) Then
        Exit Sub
    End If

    If Not currentDoc.PropertySets.PropertySetExists("Worksheet Data") Then

        worksheetData = currentDoc.PropertySets.Add("Worksheet Data")

        Dim propTabName As Inventor.Property =
            worksheetData.Add(tabNameString, "Worksheet Name")

        MsgBox("Worksheet Name: " & propTabName.Value & " has been added")

    Else

        worksheetData = currentDoc.PropertySets.Item("Worksheet Data")

        worksheetData.Item("Worksheet Name").Value = tabNameString

        MsgBox("New Worksheet Name: " &
               worksheetData.Item("Worksheet Name").Value)

    End If

End Sub


Function BrowseFile() As String

    Dim dlg As FileDialog = Nothing

    ThisApplication.CreateFileDialog(dlg)

    With dlg
        .DialogTitle = "Select a file"
        '.Filter = "Inventor Files (*.ipt;*.iam;*.idw;*.ipn;*.dwg)|*.ipt;*.iam;*.idw;*.ipn;*.dwg|All Files (*.*)|*.*"
        .FilterIndex = 1
        .InitialDirectory = ThisApplication.DesignProjectManager.ActiveDesignProject.WorkspacePath
        .MultiSelectEnabled = False
        .CancelError = True
    End With

    Try

        dlg.ShowOpen()
        Return dlg.FileName

    Catch

        Return String.Empty

    End Try

End Function
