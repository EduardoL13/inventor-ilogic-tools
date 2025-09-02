Sub SetupSpreadsheet()

Dim currentDoc As Object
Set currentDoc = ThisApplication.ActiveDocument

Dim DSData As PropertySet

' Set Spreadsheet Document
If currentDoc.PropertySets.PropertySetExists("Spreadsheet Document") = False Then

    Set SSData = currentDoc.PropertySets.Add("Spreadsheet Document")
    docNameString = BrowseFile()
    
    Dim propDocName As Property
    Set propDocName = SSData.Add(docNameString, "File Name")
    
    MsgBox ("File Name:" & propDocName.Value & " has been added")
    
Else
    Set SSData = currentDoc.PropertySets.Item("Spreadsheet Document")
    docNameString = BrowseFile()
    SSData.Item("File Name").Value = docNameString
    MsgBox ("New File Name:" & SSData.Item("File Name").Value)
    
End If

' Set worksheet

If currentDoc.PropertySets.PropertySetExists("Worksheet Data") = False Then

    Set worksheetData = currentDoc.PropertySets.Add("Worksheet Data")
    tabNameString = InputBox("Enter worksheet tab name")
    
    Dim propTabName As Property
    Set propTabName = worksheetData.Add(tabNameString, "Worksheet Name")
    
    MsgBox ("Worksheet Name:" & propTabName.Value & " has been added")
    
Else
    Set worksheetData = currentDoc.PropertySets.Item("Worksheet Data")
    tabNameString = InputBox("Enter worksheet tab name")
    worksheetData.Item("Worksheet Name").Value = tabNameString
    MsgBox ("New Worksheet Name:" & worksheetData.Item("Worksheet Name").Value)
    
End If

End Sub

Function BrowseFile() As String
    Dim oDlg As FileDialog
    Call ThisApplication.CreateFileDialog(oDlg)

    With oDlg
        .DialogTitle = "Select a file"
        '.Filter = "Inventor Files (.ipt;.iam;.idw;.ipn;.dwg)|.ipt;.iam;.idw;.ipn;.dwg|All Files (.)|."
        .FilterIndex = 1
        .InitialDirectory = ThisApplication.DesignProjectManager.ActiveDesignProject.WorkspacePath
        .MultiSelectEnabled = False
        .CancelError = True
    End With

    On Error Resume Next
    oDlg.ShowOpen              ' shows the Open dialog
    If Err.Number <> 0 Then    ' user canceled
        BrowseFile = ""
        Err.Clear
        Exit Function
    End If

    BrowseFile = oDlg.FileName ' full path of the chosen file
End Function

