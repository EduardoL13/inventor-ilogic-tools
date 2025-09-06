Sub BrowseLocation()


    Dim objDialog As Object
    Dim selectedPath As String
    
    Dim currentDoc As Object
    Set currentDoc = ThisApplication.ActiveDocument
    
    Dim dxfData As PropertySet
    Dim dxfLocation As Property
    
    If currentDoc.PropertySets.PropertySetExists("dxf Export Data") = False Then

        Set dxfData = currentDoc.PropertySets.Add("dxf Export Data")
        Set dxfLocation = dxfData.Add("", "Location")
        'MsgBox ("File Name:" & dxfLocation.Value & " has been added")
    
    Else
    
        Set dxfData = currentDoc.PropertySets.Item("dxf Export Data")
        Set dxfLocation = dxfData.Item("Location")
        
    End If
    
    ' Create a FolderBrowserDialog
    Set objDialog = CreateObject("Shell.Application").BrowseForFolder(0, "Select folder to save occurrences", 0, 0)
    
    If Not objDialog Is Nothing Then
        selectedPath = objDialog.Items().Item().Path
        MsgBox "You selected: " & selectedPath
        ' You can now use selectedPath to save your files
        dxfLocation.Value = selectedPath
        
    Else
        MsgBox "No folder selected."
        dxfLocation.Value = ""
        
    End If
    
    
End Sub
