Sub Main
'    ' Verificar que el documento activo sea un plano (DrawingDocument)
'    If ThisApplication.ActiveDocument.DocumentType <> kDrawingDocumentObject Then
'        MessageBox.Show("Esta regla debe ejecutarse desde un archivo de dibujo (.idw/.dwg).", "iLogic")
'        Return
'    End If

    Dim oDoc As DrawingDocument = ThisDoc.Document
    Dim oSheet As Sheet = oDoc.ActiveSheet

    Dim qtyInput As String
    qtyInput = InputBox("Ingrese la cantidad (QTY):", "Cantidad de Ensamblajes", "1")
    
    If qtyInput = "" Then Return
    	
    ' Definir el texto de la nota con el QTY ingresado
    Dim noteText As String = "FABRICATE " & qtyInput & " ASSEMBLIES PER PRODUCT"
	Dim sizeString As String = "0.32"
	
    ' Definir el punto de inserción de la nota en la hoja (en centímetros, centro de la hoja)
    Dim oTG As TransientGeometry = ThisApplication.TransientGeometry
    Dim oPoint As Point2d = oTG.CreatePoint2d(oSheet.Width / 2, oSheet.Height / 2)
    
    ' Añadir la nota general a la hoja
	Dim oNote As GeneralNote
    oNote = oSheet.DrawingNotes.GeneralNotes.AddFitted(oPoint, baseText)
	
    oNote.FormattedText = "<StyleOverride FontSize='" & sizeString & "'>" & noteText & "</StyleOverride>"
    
    MessageBox.Show("Nota agregada correctamente.", "iLogic")
End Sub
