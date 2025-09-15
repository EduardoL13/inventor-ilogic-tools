Sub SetupSpreadsheet()

    Dim currentDoc As Object
    Set currentDoc = ThisApplication.ActiveDocument

    Dim SSData As PropertySet

    ' Set Spreadsheet Document -----------------------------------------------------

    ' Browse Spreadsheet File
    docNameString = BrowseFile()
    
    ' Cancels if user cancel in the dialogbox
    If docNameString = "" Then
        Exit Sub
    End If

' Conditional to see if property already exists
    If currentDoc.PropertySets.PropertySetExists("Spreadsheet Document") = False Then
    
        ' Adds Spreadsheet File Data
        Set SSData = currentDoc.PropertySets.Add("Spreadsheet Document")

    
        ' Creates and adds Property
        Dim propDocName As Property
        Set propDocName = SSData.Add(docNameString, "File Name")
     

        MsgBox ("File Name:" & propDocName.Value & " has been added")
  
    ' Set the property given that it already exists
    Else
    
        'Declares and updates Spreadsheet Worksheet Data
        Set SSData = currentDoc.PropertySets.Item("Spreadsheet Document")

    
        ' Updates Property
        SSData.Item("File Name").Value = docNameString
        MsgBox ("New File Name:" & SSData.Item("File Name").Value)
    
    End If

' Set worksheet --------------------------------------------------------------------

    Dim worksheetData As PropertySet

    ' Assign name of worksheet to access within the spreadsheet
    tabNameString = InputBox("Enter worksheet tab name")

    ' Cancels if user cancel in the dialogbox
    If tabNameString = "" Then
        Exit Sub
    End If


    If currentDoc.PropertySets.PropertySetExists("Worksheet Data") = False Then

        ' Adds Spreadsheet Worksheet Data
        Set worksheetData = currentDoc.PropertySets.Add("Worksheet Data")

    
        ' Creates Property
        Dim propTabName As Property
        Set propTabName = worksheetData.Add(tabNameString, "Worksheet Name")
        
    
        MsgBox ("Worksheet Name:" & propTabName.Value & " has been added")
    
    Else
         
        'Declares worksheet Data Property set
        Set worksheetData = currentDoc.PropertySets.Item("Worksheet Data")

        ' Updates Property
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

