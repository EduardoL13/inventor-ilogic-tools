Sub Main () ' v1
	'Objetivo: escribir en un DS uno o más parámetros que se requieran de Inventor
	'Declarations
	Dim invDoc As PartDocument = ThisDoc.Document ' Documento activo

	Dim inventorParamList As UserParameters = invDoc.ComponentDefinition.Parameters.UserParameters
	Dim noPDS As Integer = inventorParamList.Count ' Insertar número de parámetros que se desean
    Dim noListPDS As Integer = noPDS - 1 'Número de parámetros para poner en los arrays

'-----------------INPUT----------------------

   'Dim nomParametersForDS(noListPDS) As String ' Listado de nombres de parámetros que se quieren escribir en el DS
    Dim writeParamsList(noListPDS) As Object
   'MsgBox(writeParamsList.Length)

    '---------------------------------------------

   'Dim writeParamsList As ReferenceParameter = inventorParamList.Item(nomParametersForDS) 'Genera listado de parámetros vacío

   Dim i As Integer

   For i=0 To writeParamsList.Length-1
      writeParamsList(i) = inventorParamList.Item(i+1) 'llena el listado de los parámetros que se quieren 
   Next


' Comms with data hub excel

	Dim file As String
    Dim tab As String = InputBox("Enter the excel worksheet name", "Worksheet to Export")

    If tab = ""
		Exit Sub
	End If

    Try
		
        file = invDoc.PropertySets.Item("Spreadsheet Document").Item("File Name").Value
	
    Catch:

	    MsgBox("Communications with Data Hub are not set. Please try again after setting them")
	Exit Sub
    End Try


	'Desired row and column input in given worksheet
	colParameterName = "A" ' Column in which parameter names are listed
	colParameterValue = "B" 'Column in which parameter values are listed
	colParameterUnits = "C" 'Column in which parameter units are listed
	

	'Row counter definition
	Dim RowCounter As Integer = 2 'Row de inicio para escribir datos


	For Each param As UserParameter In writeParamsList
	    'If param.IsKey = True Then
		If param.Comment.Contains("Profile Parameter")
            If RowCounter = 2 Then
	            GoExcel.CellValue(file, tab, colParameterName & RowCounter) = param.Name
		    	cf = unitsEval(param.Units)
	        	GoExcel.CellValue(colParameterValue & RowCounter) = param.Value * cf
	        	GoExcel.CellValue(colParameterUnits & RowCounter) = param.Units
    		Else
	        	GoExcel.CellValue(colParameterName & RowCounter) = param.Name
		    	cf = unitsEval(param.Units)
	        	GoExcel.CellValue(colParameterValue & RowCounter) = param.Value * cf
	        	GoExcel.CellValue(colParameterUnits & RowCounter) = param.Units
			End If
			RowCounter = RowCounter + 1 
    	End If

	Next

	GoExcel.Save

End Sub


Function unitsEval(units As String)
    If units = "in"
	    cf = 1 / 2.54
	Else If units = "ul"
		cf = 1
	Else
		cf = 10
	End If
	Return cf
End Function

