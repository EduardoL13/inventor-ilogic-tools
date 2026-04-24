Sub Main()
	
	Dim esteDoc As PartDocument = ThisDoc.Document
	Dim partDef As PartComponentDefinition = esteDoc.ComponentDefinition


	sufix = InputBox("Enter Body Profile Name", "Body Name", "ej:CoreTube")

	Dim pLength1 As String = "pLength1"
	paramName = pLength1 & sufix 

	partDef.Parameters.UserParameters.AddByValue(pLength1 & sufix, 1, "in")
	partDef.Parameters.UserParameters(paramName).Comment = "Profile Parameter"

	Dim pLength2 As String = "pLength2"
	paramName = pLength2 & sufix 

	partDef.Parameters.UserParameters.AddByValue(pLength2 & sufix, 1, "in")
	partDef.Parameters.UserParameters(paramName).Comment = "Profile Parameter"

	Dim pThk As String = "pThk"
	paramName = pThk & sufix 

	partDef.Parameters.UserParameters.AddByValue(pThk & sufix, 1, "in")
	partDef.Parameters.UserParameters(paramName).Comment = "Profile Parameter"

End Sub
