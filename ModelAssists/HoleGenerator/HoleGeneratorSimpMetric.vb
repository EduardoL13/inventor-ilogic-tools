Sub main 
    ' Simplified: Asume 50% hard materials y standard fit
	
    Dim esteDoc As PartDocument = ThisDoc.Document
	Dim partDef As PartComponentDefinition = esteDoc.ComponentDefinition
	Dim listParams As UserParameters = partDef.Parameters.UserParameters
    convFactor = 1/10 'Conversion factor mm -> cm
	
	sufix = InputBox("Hole Generator","Feature Name","ex:AtCover")
		
	If listParams.Item("HoleDataType").Value = "Simplified" Then
		
	    If listParams.Item("HoleClearanceType").Value = "Clear" Then
			
	    	If listParams.Item("HoleTable").Value = "M1.5" Then
				
		    	listParams.AddByValue("diaHole" & sufix, 1.65 * convFactor, "mm")
				listParams.Item("diaHole" & sufix).Comment = "Free fit for M1.5 fastener"
				MsgBox("Parameters for holes have been generated")	
				
			Else If listParams.Item("HoleTable").Value = "M1.6"
				
			    listParams.AddByValue("diaHole" & sufix, 1.75 * convFactor, "mm")	
				listParams.Item("diaHole" & sufix).Comment = "Free fit for M1.6 fastener"
				MsgBox("Parameters for holes have been generated")	
				
			Else If listParams.Item("HoleTable").Value = "M1.8"
				
			    listParams.AddByValue("diaHole" & sufix, 2 * convFactor, "mm")	
				listParams.Item("diaHole" & sufix).Comment = "Free fit for M1.8 fastener"
				MsgBox("Parameters for holes have been generated")	
				
			Else If listParams.Item("HoleTable").Value = "M2"

			    listParams.AddByValue("diaHole" & sufix, 2.2 * convFactor, "in")
				listParams.Item("diaHole" & sufix).Comment = "Free fit for M2 fastener"
				MsgBox("Parameters for holes have been generated")	
				
			Else If listParams.Item("HoleTable").Value = "M2.2"

			    listParams.AddByValue("diaHole" & sufix, 2.4 * convFactor, "mm")
				listParams.Item("diaHole" & sufix).Comment = "Free fit for M2.2 fastener"
				MsgBox("Parameters for holes have been generated")	
				
			Else If listParams.Item("HoleTable").Value = "M3"

			    listParams.AddByValue("diaHole" & sufix, 3.3 * convFactor, "mm")	
				listParams.Item("diaHole" & sufix).Comment = "Free fit for M3 fastener"
				MsgBox("Parameters for holes have been generated")	
				
				
			Else If listParams.Item("HoleTable").Value = "M3.5"

			    listParams.AddByValue("diaHole" & sufix, 3.85 * convFactor, "mm")
				listParams.Item("diaHole" & sufix).Comment = "Free fit for M3.5 fastener"
				MsgBox("Parameters for holes have been generated")		
				
			Else If listParams.Item("HoleTable").Value = "M4"

			    listParams.AddByValue("diaHole" & sufix, 4.4 * convFactor, "mm")
				listParams.Item("diaHole" & sufix).Comment = "Free fit for M4 fastener"
				MsgBox("Parameters for holes have been generated")		
				
			Else If listParams.Item("HoleTable").Value = "M4.5"

			    listParams.AddByValue("diaHole" & sufix, 5 * convFactor, "mm")
				listParams.Item("diaHole" & sufix).Comment = "Free fit for M4.5 fastener"
				MsgBox("Parameters for holes have been generated")	
				
			Else If listParams.Item("HoleTable").Value = "M5"

			    listParams.AddByValue("diaHole" & sufix, 5.5 * convFactor, "mm")
				listParams.Item("diaHole" & sufix).Comment = "Free fit for M5 fastener"
				MsgBox("Parameters for holes have been generated")		
				
			Else If listParams.Item("HoleTable").Value = "M5.5"

			    listParams.AddByValue("diaHole" & sufix,  6.1*convFactor, "mm")
				listParams.Item("diaHole" & sufix).Comment = "Free fit for M5.5 fastener"
				MsgBox("Parameters for holes have been generated")		
				
			Else If listParams.Item("HoleTable").Value = "M6"

			    listParams.AddByValue("diaHole" & sufix,  6.6*convFactor, "mm")	
				listParams.Item("diaHole" & sufix).Comment = "Free fit for M6 fastener"
				MsgBox("Parameters for holes have been generated")
				
			Else If listParams.Item("HoleTable").Value = "M7"

			    listParams.AddByValue("diaHole" & sufix,  7.7*convFactor, "mm") '0.397
				listParams.Item("diaHole" & sufix).Comment = "Free fit for M7 fastener"
				
				MsgBox("Parameters for holes have been generated")				
			Else If listParams.Item("HoleTable").Value = "M8"

			    listParams.AddByValue("diaHole" & sufix, 8.8 * convFactor, "mm") '0.5938
				listParams.Item("diaHole" & sufix).Comment = "Free fit for M8 fastener"
				MsgBox("Parameters for holes have been generated")
				
			Else If listParams.Item("HoleTable").Value = "M9"

			    listParams.AddByValue("diaHole" & sufix,  9.9*convFactor, "mm") '0.5938
				listParams.Item("diaHole" & sufix).Comment = "Free fit for M9 fastener"
				MsgBox("Parameters for holes have been generated")	
				
			Else If listParams.Item("HoleTable").Value = "M10"

			    listParams.AddByValue("diaHole" & sufix, 11 * convFactor, "mm") '0.5938
				listParams.Item("diaHole" & sufix).Comment = "Free fit for M10 fastener"
				MsgBox("Parameters for holes have been generated")
				
			Else If listParams.Item("HoleTable").Value = "M11"

			    listParams.AddByValue("diaHole" & sufix, 12.1 * convFactor, "mm") '0.5938
				listParams.Item("diaHole" & sufix).Comment = "Free fit for M11 fastener"
				MsgBox("Parameters for holes have been generated")		
				
			Else If listParams.Item("HoleTable").Value = "M12"

			    listParams.AddByValue("diaHole" & sufix, 13.2 * convFactor, "mm") '0.5938	
				listParams.Item("diaHole" & sufix).Comment = "Free fit for M12 fastener"
				MsgBox("Parameters for holes have been generated")			
				
			Else If listParams.Item("HoleTable").Value = "M14"

			    listParams.AddByValue("diaHole" & sufix, 15.5 * convFactor, "mm") '0.5938	
				listParams.Item("diaHole" & sufix).Comment = "Free fit for M14 fastener"
				MsgBox("Parameters for holes have been generated")	
				
			Else If listParams.Item("HoleTable").Value = "M15"

			    listParams.AddByValue("diaHole" & sufix, 16.5 * convFactor, "mm") '0.5938	
				listParams.Item("diaHole" & sufix).Comment = "Free fit for M15 fastener"
				MsgBox("Parameters for holes have been generated")		
				
			Else If listParams.Item("HoleTable").Value = "M16"

			    listParams.AddByValue("diaHole" & sufix, 17.5 * convFactor, "mm") '0.5938	
				listParams.Item("diaHole" & sufix).Comment = "Free fit for M16 fastener"
				MsgBox("Parameters for holes have been generated")			
				
			Else If listParams.Item("HoleTable").Value = "M17"

			    listParams.AddByValue("diaHole" & sufix, 18.5 * convFactor, "mm") '0.5938	
				listParams.Item("diaHole" & sufix).Comment = "Free fit for M17 fastener"
				MsgBox("Parameters for holes have been generated")	
				
			Else If listParams.Item("HoleTable").Value = "M18"

			    listParams.AddByValue("diaHole" & sufix, 20 * convFactor, "mm") '0.5938
				listParams.Item("diaHole" & sufix).Comment = "Free fit for M18 fastener"
				MsgBox("Parameters for holes have been generated")
			
			
			Else If listParams.Item("HoleTable").Value = "M19"

			    listParams.AddByValue("diaHole" & sufix, 21 * convFactor, "mm") '0.5938
				listParams.Item("diaHole" & sufix).Comment = "Free fit for M19 fastener"
				MsgBox("Parameters for holes have been generated")
			
			Else If listParams.Item("HoleTable").Value = "M20"

			    listParams.AddByValue("diaHole" & sufix, 22 * convFactor, "mm") '0.5938
				listParams.Item("diaHole" & sufix).Comment = "Free fit for M20 fastener"
				MsgBox("Parameters for holes have been generated")
						
			End If				
			
		Else If listParams.Item("HoleClearanceType").Value = "Tapped" Then
			' Assumptions:- hard material - threads per unit: coarse
			
	    	If listParams.Item("HoleTable").Value = "M1.5" Then
				
		    	listParams.AddByValue("diaHole" & sufix, 1.25 * convFactor, "mm")
				listParams.Item("diaHole" & sufix).Comment = "Tapped fit for M1.5 fastener"
				MsgBox("Parameters for holes have been generated")	
				
			Else If listParams.Item("HoleTable").Value = "M1.6"
				
			    listParams.AddByValue("diaHole" & sufix, 1.35 * convFactor, "mm")	
				listParams.Item("diaHole" & sufix).Comment = "Tapped fit for M1.6 fastener"
				MsgBox("Parameters for holes have been generated")	
				
			Else If listParams.Item("HoleTable").Value = "M1.8"
				
			    listParams.AddByValue("diaHole" & sufix, 1.55 * convFactor, "mm")	
				listParams.Item("diaHole" & sufix).Comment = "Tapped fit for M1.8 fastener"
				MsgBox("Parameters for holes have been generated")	
				
			Else If listParams.Item("HoleTable").Value = "M2"

			    listParams.AddByValue("diaHole" & sufix, 1.7 * convFactor, "in")
				listParams.Item("diaHole" & sufix).Comment = "Tapped fit for M2 fastener"
				MsgBox("Parameters for holes have been generated")	
				
			Else If listParams.Item("HoleTable").Value = "M2.2"

			    listParams.AddByValue("diaHole" & sufix, 1.9 * convFactor, "mm")
				listParams.Item("diaHole" & sufix).Comment = "Tapped fit for M2.2 fastener"
				MsgBox("Parameters for holes have been generated")	
				
			Else If listParams.Item("HoleTable").Value = "M3"

			    listParams.AddByValue("diaHole" & sufix, 2.6 * convFactor, "mm")	
				listParams.Item("diaHole" & sufix).Comment = "Tapped fit for M3 fastener"
				MsgBox("Parameters for holes have been generated")	
				
				
			Else If listParams.Item("HoleTable").Value = "M3.5"

			    listParams.AddByValue("diaHole" & sufix, 3.1 * convFactor, "mm")
				listParams.Item("diaHole" & sufix).Comment = "Tapped fit for M3.5 fastener"
				MsgBox("Parameters for holes have been generated")		
				
			Else If listParams.Item("HoleTable").Value = "M4"

			    listParams.AddByValue("diaHole" & sufix, 3.5 * convFactor, "mm")
				listParams.Item("diaHole" & sufix).Comment = "Tapped fit for M4 fastener"
				MsgBox("Parameters for holes have been generated")		
				
			Else If listParams.Item("HoleTable").Value = "M4.5"

			    listParams.AddByValue("diaHole" & sufix, 4 * convFactor, "mm")
				listParams.Item("diaHole" & sufix).Comment = "Tapped fit for M4.5 fastener"
				MsgBox("Parameters for holes have been generated")	
				
			Else If listParams.Item("HoleTable").Value = "M5"

			    listParams.AddByValue("diaHole" & sufix, 4.4 * convFactor, "mm")
				listParams.Item("diaHole" & sufix).Comment = "Tapped fit for M5 fastener"
				MsgBox("Parameters for holes have been generated")		
				
			Else If listParams.Item("HoleTable").Value = "M5.5"

			    listParams.AddByValue("diaHole" & sufix,  4.9*convFactor, "mm")
				listParams.Item("diaHole" & sufix).Comment = "Tapped fit for M5.5 fastener"
				MsgBox("Parameters for holes have been generated")		
				
			Else If listParams.Item("HoleTable").Value = "M6"

			    listParams.AddByValue("diaHole" & sufix,  5.4*convFactor, "mm")	
				listParams.Item("diaHole" & sufix).Comment = "Tapped fit for M6 fastener"
				MsgBox("Parameters for holes have been generated")
				
			Else If listParams.Item("HoleTable").Value = "M7"

			    listParams.AddByValue("diaHole" & sufix,  6.4*convFactor, "mm") '0.397
				listParams.Item("diaHole" & sufix).Comment = "Tapped fit for M7 fastener"
				
				MsgBox("Parameters for holes have been generated")				
			Else If listParams.Item("HoleTable").Value = "M8"

			    listParams.AddByValue("diaHole" & sufix, 7.2 * convFactor, "mm") '0.5938
				listParams.Item("diaHole" & sufix).Comment = "Tapped fit for M8 fastener"
				MsgBox("Parameters for holes have been generated")
				
			Else If listParams.Item("HoleTable").Value = "M9"

			    listParams.AddByValue("diaHole" & sufix,  8.2*convFactor, "mm") '0.5938
				listParams.Item("diaHole" & sufix).Comment = "Tapped fit for M9 fastener"
				MsgBox("Parameters for holes have been generated")	
				
			Else If listParams.Item("HoleTable").Value = "M10"

			    listParams.AddByValue("diaHole" & sufix, 9 * convFactor, "mm") '0.5938
				listParams.Item("diaHole" & sufix).Comment = "Tapped fit for M10 fastener"
				MsgBox("Parameters for holes have been generated")
				
			Else If listParams.Item("HoleTable").Value = "M11"

			    listParams.AddByValue("diaHole" & sufix, 10 * convFactor, "mm") '0.5938
				listParams.Item("diaHole" & sufix).Comment = "Tapped fit for M11 fastener"
				MsgBox("Parameters for holes have been generated")		
				
			Else If listParams.Item("HoleTable").Value = "M12"

			    listParams.AddByValue("diaHole" & sufix, 10.9 * convFactor, "mm") '0.5938	
				listParams.Item("diaHole" & sufix).Comment = "Tapped fit for M12 fastener"
				MsgBox("Parameters for holes have been generated")			
				
			Else If listParams.Item("HoleTable").Value = "M14"

			    listParams.AddByValue("diaHole" & sufix, 12.7 * convFactor, "mm") '0.5938	
				listParams.Item("diaHole" & sufix).Comment = "Tapped fit for M14 fastener"
				MsgBox("Parameters for holes have been generated")	
				
			Else If listParams.Item("HoleTable").Value = "M15"

			    listParams.AddByValue("diaHole" & sufix, 14 * convFactor, "mm") '0.5938	
				listParams.Item("diaHole" & sufix).Comment = "Tapped fit for M15 fastener"
				MsgBox("Parameters for holes have been generated")		
				
			Else If listParams.Item("HoleTable").Value = "M16"

			    listParams.AddByValue("diaHole" & sufix, 14.75 * convFactor, "mm") '0.5938	
				listParams.Item("diaHole" & sufix).Comment = "Tapped fit for M16 fastener"
				MsgBox("Parameters for holes have been generated")			
				
			Else If listParams.Item("HoleTable").Value = "M17"

			    listParams.AddByValue("diaHole" & sufix, 16 * convFactor, "mm") '0.5938	
				listParams.Item("diaHole" & sufix).Comment = "Tapped fit for M17 fastener"
				MsgBox("Parameters for holes have been generated")	
				
			Else If listParams.Item("HoleTable").Value = "M18"

			    listParams.AddByValue("diaHole" & sufix, 16.5 * convFactor, "mm") '0.5938
				listParams.Item("diaHole" & sufix).Comment = "Tapped fit for M18 fastener"
				MsgBox("Parameters for holes have been generated")
						
			
			Else If listParams.Item("HoleTable").Value = "M19"

			    listParams.AddByValue("diaHole" & sufix, 17.5 * convFactor, "mm") '0.5938
				listParams.Item("diaHole" & sufix).Comment = "Tapped fit for M19 fastener"
				MsgBox("Parameters for holes have been generated")
			
			Else If listParams.Item("HoleTable").Value = "M20"

			    listParams.AddByValue("diaHole" & sufix, 18.5 * convFactor, "mm") '0.5938
				listParams.Item("diaHole" & sufix).Comment = "Tapped fit for M20 fastener"
				MsgBox("Parameters for holes have been generated")
			 
			End If
		End If	
    End If
    
End Sub
