Sub main 

    Dim esteDoc As PartDocument = ThisDoc.Document
	Dim partDef As PartComponentDefinition = esteDoc.ComponentDefinition
	Dim listParams As UserParameters = partDef.Parameters.UserParameters
	convFactor = 2.54
	
			
	If listParams.Item("noFastener").Value = "#4" Then
		listParams.Item("dia").Value = 0.112 * convFactor
		MsgBox("dia set to " & listParams.Item("noFastener").Value & " fastener")	
				
	Else If listParams.Item("noFastener").Value = "#6"
		listParams.Item("dia").Value = 0.136 * convFactor
		MsgBox("dia set to " & listParams.Item("noFastener").Value & " fastener")	
				
	Else If listParams.Item("noFastener").Value = "#8"
		listParams.Item("dia").Value = 0.164 * convFactor
		MsgBox("dia set to " & listParams.Item("noFastener").Value & " fastener")	
				
	Else If listParams.Item("noFastener").Value = "#10"
		listParams.Item("dia").Value = 0.19 * convFactor
		MsgBox("dia set to " & listParams.Item("noFastener").Value & " fastener")	
		
	Else If listParams.Item("noFastener").Value = "#12"
		listParams.Item("dia").Value = 0.216 * convFactor
		MsgBox("dia set to " & listParams.Item("noFastener").Value & " fastener")			

	Else If listParams.Item("noFastener").Value = "1/4"
		listParams.Item("dia").Value = 0.25 * convFactor
		MsgBox("dia set to " & listParams.Item("noFastener").Value & " fastener")
		
	Else If listParams.Item("noFastener").Value = "5/16"
		listParams.Item("dia").Value = 0.313 * convFactor
		MsgBox("dia set to " & listParams.Item("noFastener").Value & " fastener")

	Else If listParams.Item("noFastener").Value = "3/8"
		listParams.Item("dia").Value = 0.375 * convFactor
		MsgBox("dia set to " & listParams.Item("noFastener").Value & " fastener")

	Else If listParams.Item("noFastener").Value = "7/16"
		listParams.Item("dia").Value = 0.438 * convFactor
		MsgBox("dia set to " & listParams.Item("noFastener").Value & " fastener")

	Else If listParams.Item("noFastener").Value = "1/2"
		listParams.Item("dia").Value = 0.5 * convFactor
		MsgBox("dia set to " & listParams.Item("noFastener").Value & " fastener")
		
	Else If listParams.Item("noFastener").Value = "9/16"
		listParams.Item("dia").Value = 0.5625 * convFactor
		MsgBox("dia set to " & listParams.Item("noFastener").Value & " fastener")
		
	Else If listParams.Item("noFastener").Value = "5/8"
		listParams.Item("dia").Value = 0.625 * convFactor
		MsgBox("dia set to " & listParams.Item("noFastener").Value & " fastener")	
		
	Else If listParams.Item("noFastener").Value = "3/4"
		listParams.Item("dia").Value = 0.75 * convFactor
		MsgBox("dia set to " & listParams.Item("noFastener").Value & " fastener")
		
	End If
				
	
    
End Sub
