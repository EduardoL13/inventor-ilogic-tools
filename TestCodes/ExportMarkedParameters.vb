Sub Main () ' v1
'Objetivo: escribir en un DS uno o más parámetros que se requieran de Inventor
' Declarations
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

' Target file to write the parameter


Dim file As String
Dim tab As String

file = invDoc.PropertySets.Item("Spreadsheet Document").Item("File Name").Value
tab = invDoc.PropertySets.Item("Worksheet Data").Item("Worksheet Name").Value

'Row counter definition
Dim RowCounter As Integer = 2 'Row de inicio para escribir datos


For Each param As UserParameter In writeParamsList
	If param.IsKey = True Then
        If RowCounter = 2 Then
	        GoExcel.CellValue(file, tab, "A" & RowCounter) = param.Name
		    cf = unitsEval(param.Units)
	        GoExcel.CellValue("B" & RowCounter) = param.Value * cf
	        GoExcel.CellValue("C" & RowCounter) = param.Units
    	Else
	        GoExcel.CellValue("A" & RowCounter) = param.Name
		    cf = unitsEval(param.Units)
	        GoExcel.CellValue("B" & RowCounter) = param.Value * cf
	        GoExcel.CellValue("C" & RowCounter) = param.Units
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
