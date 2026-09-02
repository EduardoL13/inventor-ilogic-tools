Sub Main ()
    
	'Set Parameters and features access
	'MsgBox("Conf triggered")
	Dim currentDoc As PartDocument = ThisDoc.Document
    Dim oCompDef As PartComponentDefinition = currentDoc.ComponentDefinition
	Dim mParams As ModelParameters = oCompDef.Parameters.ModelParameters
    Dim uParams As UserParameters = oCompDef.Parameters.UserParameters
	convFactor = 2.54
 
	'Configuration list
	configuration1 = "MMC"
	configuration2 = "LMC"
	configuration3 = "NOMINAL"
	configuration4 = "CRITICAL"
	
	' + = ModelValueTypeEnum.kUpperModelValue
	' - = ModelValueTypeEnum.kLowerModelValue
	' Nom = ModelValueTypeEnum.kNominalModelValue
	
	
	'Changes configuration
	

    If parameter("matCondition") = configuration1 Then 'MMC
	' Listar Supresiones y cambios en kep parameters
        mParams("mLengthTotal").ModelValueType = ModelValueTypeEnum.kUpperValue
		mParams("mGapCladdingFrame").ModelValueType = ModelValueTypeEnum.kLowerValue
		
		uParams("tolCompenser").Value = 1/32*convFactor ' Compensa el 1/32 que se aumenta la lengthTotal
		
		MsgBox("Active Condition: " & parameter("matCondition"))
		
    Else If parameter("matCondition") = configuration2 Then 'LMC
        mParams("mLengthTotal").ModelValueType = ModelValueTypeEnum.kLowerValue
		mParams("mGapCladdingFrame").ModelValueType = ModelValueTypeEnum.kUpperValue
		
		uParams("tolCompenser").Value = 0
		
		MsgBox("Active Condition: " & parameter("matCondition"))
		
	Else If parameter("matCondition") = configuration3 Then 'NOMINAL
        mParams("mLengthTotal").ModelValueType = ModelValueTypeEnum.kNominalValue
		mParams("mGapCladdingFrame").ModelValueType = ModelValueTypeEnum.kNominalValue
		
		uParams("tolCompenser").Value = 0
		
		MsgBox("Active Condition: " & parameter("matCondition"))



	Else If parameter("matCondition") = configuration4 Then 'CRITICAL
        mParams("mLengthTotal").ModelValueType = ModelValueTypeEnum.kLowerValue
		mParams("mGapCladdingFrame").ModelValueType = ModelValueTypeEnum.kLowerValue
		
		uParams("tolCompenser").Value = 0
		
		MsgBox("Active Condition: " & parameter("matCondition"))

	End If


End Sub
