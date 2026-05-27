Sub main 

    Dim esteDoc As PartDocument = ThisDoc.Document
	  Dim partDef As PartComponentDefinition = esteDoc.ComponentDefinition
	  Dim listParams As UserParameters = partDef.Parameters.UserParameters
    convFactor = 2.54 'Conversion factor in -> cm
	
	sufix = InputBox("Hole Generator","Feature Name","ex:AtCover")
	
	If listParams.Item("DrilledHoleType").Value = "SD" Then
		
	    If listParams.Item("HoleClearanceType").Value = "Clear" Then
			
	    	If listParams.Item("HoleTableImp").Value = "#0" Then
		    	'listParams.AddByValue("diaFastener" & sufix, 0.06 * convFactor, "in")
		    	listParams.AddByValue("diaHole" & sufix, 0.07 * convFactor, "in")
				listParams.Item("diaHole" & sufix).Comment = "Free fit for .06in dia fastener (#0)"
				MsgBox("Parameters for holes have been generated")	
				
			Else If listParams.Item("HoleTableImp").Value = "#1"
			    'listParams.AddByValue("diaFastener" & sufix, 0.073 * convFactor, "in")
			    listParams.AddByValue("diaHole" & sufix, 0.081 * convFactor, "in")	
				listParams.Item("diaHole" & sufix).Comment = "Free fit for .073in dia fastener (#1)"
				MsgBox("Parameters for holes have been generated")	
				
			Else If listParams.Item("HoleTableImp").Value = "#2"
			    'listParams.AddByValue("diaFastener" & sufix, 0.086 * convFactor, "in")
			    listParams.AddByValue("diaHole" & sufix, 0.096 * convFactor, "in")	
				listParams.Item("diaHole" & sufix).Comment = "Free fit for .086in dia fastener (#2)"
				MsgBox("Parameters for holes have been generated")	
				
			Else If listParams.Item("HoleTableImp").Value = "#3"
			    'listParams.AddByValue(diaFastener & sufix, 0.099 * convFactor, "in")
			    listParams.AddByValue("diaHole" & sufix, 0.11 * convFactor, "in")
				listParams.Item("diaHole" & sufix).Comment = "Free fit for .099in dia fastener (#3)"
				MsgBox("Parameters for holes have been generated")	
				
			Else If listParams.Item("HoleTableImp").Value = "#4"
			    'listParams.AddByValue(diaFastener & sufix, 0.112 * convFactor, "in")
			    listParams.AddByValue("diaHole" & sufix, 0.1285 * convFactor, "in")
				listParams.Item("diaHole" & sufix).Comment = "Free fit for .112in dia fastener (#4)"
				MsgBox("Parameters for holes have been generated")	
				
			Else If listParams.Item("HoleTableImp").Value = "#5"
			    'listParams.AddByValue(diaFastener & sufix, 0.125 * convFactor, "in")
			    listParams.AddByValue("diaHole" & sufix, 0.1285 * convFactor, "in")	
				listParams.Item("diaHole" & sufix).Comment = "Free fit for .125in dia fastener (#5)"
				MsgBox("Parameters for holes have been generated")	
				
				
			Else If listParams.Item("HoleTableImp").Value = "#6"
			    'listParams.AddByValue(diaFastener & sufix, 0.138 * convFactor, "in")
			    listParams.AddByValue("diaHole" & sufix, 0.1495 * convFactor, "in")
				listParams.Item("diaHole" & sufix).Comment = "Free fit for .138in dia fastener (#6)"
				MsgBox("Parameters for holes have been generated")		
				
			Else If listParams.Item("HoleTableImp").Value = "#8"
			    'listParams.AddByValue(diaFastener & sufix, 0.164 * convFactor, "in")
			    listParams.AddByValue("diaHole" & sufix, 0.177 * convFactor, "in")
				listParams.Item("diaHole" & sufix).Comment = "Free fit for .164in dia fastener (#8)"
				MsgBox("Parameters for holes have been generated")		
				
			Else If listParams.Item("HoleTableImp").Value = "#10"
			    'listParams.AddByValue(diaFastener & sufix, 0.19 * convFactor, "in")
			    listParams.AddByValue("diaHole" & sufix, 0.201 * convFactor, "in")
				listParams.Item("diaHole" & sufix).Comment = "Free fit for .19in dia fastener (#10)"
				MsgBox("Parameters for holes have been generated")	
				
			Else If listParams.Item("HoleTableImp").Value = "#12"
			    'listParams.AddByValue(diaFastener & sufix, 0.216 * convFactor, "in")
			    listParams.AddByValue("diaHole" & sufix, 0.228 * convFactor, "in")
				listParams.Item("diaHole" & sufix).Comment = "Free fit for .216in dia fastener (#12)"
				MsgBox("Parameters for holes have been generated")		
				
			Else If listParams.Item("HoleTableImp").Value = "1/4"
			    'listParams.AddByValue(diaFastener & sufix, 0.25 * convFactor, "in")
			    listParams.AddByValue("diaHole" & sufix, (0.25+.125) * convFactor, "in")
				listParams.Item("diaHole" & sufix).Comment = "Free fit for .25in dia fastener (1/4)"
				MsgBox("Parameters for holes have been generated")		
				
			Else If listParams.Item("HoleTableImp").Value = "5/16"
			    'listParams.AddByValue(diaFastener & sufix, 0.3125 * convFactor, "in")
			    listParams.AddByValue("diaHole" & sufix, (0.332+0.125) * convFactor, "in")	
				listParams.Item("diaHole" & sufix).Comment = "Free fit for .3125in dia fastener (5/16)"
				MsgBox("Parameters for holes have been generated")
				
			Else If listParams.Item("HoleTableImp").Value = "3/8"
			    'listParams.AddByValue(diaFastener & sufix, 0.375 * convFactor, "in")
			    listParams.AddByValue("diaHole" & sufix, (0.375+0.125) * convFactor, "in") '0.397
				listParams.Item("diaHole" & sufix).Comment = "Free fit for .375in dia fastener (3/8)"
				
				MsgBox("Parameters for holes have been generated")				
			Else If listParams.Item("HoleTableImp").Value = "7/16"
			    'listParams.AddByValue(diaFastener & sufix, 0.4375 * convFactor, "in")
			    listParams.AddByValue("diaHole" & sufix, (0.4375 + 0.125) * convFactor, "in") '0.5938
				listParams.Item("diaHole" & sufix).Comment = "Free fit for .4375in dia fastener (7/16)"
				MsgBox("Parameters for holes have been generated")
				
			Else If listParams.Item("HoleTableImp").Value = "1/2"
			    'listParams.AddByValue(diaFastener & sufix, 0.5 * convFactor, "in")
			    listParams.AddByValue("diaHole" & sufix, (0.5 + 0.125) * convFactor, "in") '0.5938
				listParams.Item("diaHole" & sufix).Comment = "Free fit for .5in dia fastener (1/2)"
				MsgBox("Parameters for holes have been generated")	
				
			Else If listParams.Item("HoleTableImp").Value = "9/16"
			    'listParams.AddByValue(diaFastener & sufix, 0.5625 * convFactor, "in")
			    listParams.AddByValue("diaHole" & sufix, (0.5625 + 0.125) * convFactor, "in") '0.5938
				listParams.Item("diaHole" & sufix).Comment = "Free fit for .5625in dia fastener (9/16)"
				MsgBox("Parameters for holes have been generated")
				
			Else If listParams.Item("HoleTableImp").Value = "5/8"
			    'listParams.AddByValue(diaFastener & sufix, 0.625 * convFactor, "in")
			    listParams.AddByValue("diaHole" & sufix, (0.625 + 0.125) * convFactor, "In") '0.5938
				listParams.Item("diaHole" & sufix).Comment = "Free fit for .625in dia fastener (5/8)"
				MsgBox("Parameters for holes have been generated")		
				
			Else If listParams.Item("HoleTableImp").Value = "11/16"
			    'listParams.AddByValue(diaFastener & sufix, 0.6875 * convFactor, "in")
			    listParams.AddByValue("diaHole" & sufix, (0.6875 + 0.125) * convFactor, "in") '0.5938	
				listParams.Item("diaHole" & sufix).Comment = "Free fit for .6875in dia fastener (11/16)"
				MsgBox("Parameters for holes have been generated")			
				
			Else If listParams.Item("HoleTableImp").Value = "3/4"
			    'listParams.AddByValue(diaFastener & sufix, 0.75 * convFactor, "in")
			    listParams.AddByValue("diaHole" & sufix, (0.75 + 0.125) * convFactor, "in") '0.5938	
				listParams.Item("diaHole" & sufix).Comment = "Free fit for .75in dia fastener (3/4)"
				MsgBox("Parameters for holes have been generated")	
				
			Else If listParams.Item("HoleTableImp").Value = "13/16"
			    'listParams.AddByValue(diaFastener & sufix, 0.8125 * convFactor, "in")
			    listParams.AddByValue("diaHole" & sufix, (0.8125 + 0.125) * convFactor, "in") '0.5938	
				listParams.Item("diaHole" & sufix).Comment = "Free fit for .8125in dia fastener (13/16)"
				MsgBox("Parameters for holes have been generated")		
				
			Else If listParams.Item("HoleTableImp").Value = "7/8"
			    'listParams.AddByValue(diaFastener & sufix, 0.875 * convFactor, "in")
			    listParams.AddByValue("diaHole" & sufix, (0.875 + 0.125) * convFactor, "in") '0.5938	
				listParams.Item("diaHole" & sufix).Comment = "Free fit for .875in dia fastener (7/8)"
				MsgBox("Parameters for holes have been generated")			
				
			Else If listParams.Item("HoleTableImp").Value = "15/16"
			    'listParams.AddByValue(diaFastener & sufix, 0.9375 * convFactor, "in")
			    listParams.AddByValue("diaHole" & sufix, (0.9375 + 0.125) * convFactor, "in") '0.5938	
				listParams.Item("diaHole" & sufix).Comment = "Free fit for .9375 dia fastener (15/16)"
				MsgBox("Parameters for holes have been generated")	
				
			Else If listParams.Item("HoleTableImp").Value = "1"
			    'listParams.AddByValue(diaFastener & sufix, 1 * convFactor, "in")
			    listParams.AddByValue("diaHole" & sufix, (1 + 0.125) * convFactor, "in") '0.5938
				listParams.Item("diaHole" & sufix).Comment = "Free fit for 1in dia fastener"
				MsgBox("Parameters for holes have been generated")
			End If
			
		Else If listParams.Item("HoleClearanceType").Value = "Tapped" Then
			' Assumptions:- hard material - threads per unit: coarse
			
	    	If listParams.Item("HoleTableImp").Value = "#0" Then
		    	'listParams.AddByValue("diaFastener" & sufix, 0.06 * convFactor, "in")
		    	listParams.AddByValue("diaHole" & sufix, 0.052 * convFactor, "in")
				listParams.Item("diaHole" & sufix).Comment = "Tapped for .06in dia fastener (#0)"
				MsgBox("Parameters for holes have been generated")	
				
			Else If listParams.Item("HoleTableImp").Value = "#1"
			    'listParams.AddByValue("diaFastener" & sufix, 0.073 * convFactor, "in")
			    listParams.AddByValue("diaHole" & sufix, 0.0625 * convFactor, "in")	
				listParams.Item("diaHole" & sufix).Comment = "Tapped for .073in dia fastener (#1)"
				MsgBox("Parameters for holes have been generated")	
				
			Else If listParams.Item("HoleTableImp").Value = "#2"
			    'listParams.AddByValue("diaFastener" & sufix, 0.086 * convFactor, "in")
			    listParams.AddByValue("diaHole" & sufix, 0.073 * convFactor, "in")	
				listParams.Item("diaHole" & sufix).Comment = "Tapped for .086in dia fastener (#2)"
				MsgBox("Parameters for holes have been generated")	
				
			Else If listParams.Item("HoleTableImp").Value = "#3"
			    'listParams.AddByValue(diaFastener & sufix, 0.099 * convFactor, "in")
			    listParams.AddByValue("diaHole" & sufix, 0.086 * convFactor, "in")
				listParams.Item("diaHole" & sufix).Comment = "Tapped for .099in dia fastener (#3)"
				MsgBox("Parameters for holes have been generated")	
				
			Else If listParams.Item("HoleTableImp").Value = "#4"
			    'listParams.AddByValue(diaFastener & sufix, 0.112 * convFactor, "in")
			    listParams.AddByValue("diaHole" & sufix, 0.096 * convFactor, "in")
				listParams.Item("diaHole" & sufix).Comment = "Tapped for .112in dia fastener (#4)"
				MsgBox("Parameters for holes have been generated")	
				
			Else If listParams.Item("HoleTableImp").Value = "#5"
			    'listParams.AddByValue(diaFastener & sufix, 0.125 * convFactor, "in")
			    listParams.AddByValue("diaHole" & sufix, 0.1094 * convFactor, "in")	
				listParams.Item("diaHole" & sufix).Comment = "Tapped for .125in dia fastener (#5)"
				MsgBox("Parameters for holes have been generated")	
				
				
			Else If listParams.Item("HoleTableImp").Value = "#6"
			    'listParams.AddByValue(diaFastener & sufix, 0.138 * convFactor, "in")
			    listParams.AddByValue("diaHole" & sufix, 0.116 * convFactor, "in")
				listParams.Item("diaHole" & sufix).Comment = "Tapped for .138in dia fastener (#6)"
				MsgBox("Parameters for holes have been generated")		
				
			Else If listParams.Item("HoleTableImp").Value = "#8"
			    'listParams.AddByValue(diaFastener & sufix, 0.164 * convFactor, "in")
			    listParams.AddByValue("diaHole" & sufix, 0.144 * convFactor, "in")
				listParams.Item("diaHole" & sufix).Comment = "Tapped for .164in dia fastener (#8)"
				MsgBox("Parameters for holes have been generated")		
				
			Else If listParams.Item("HoleTableImp").Value = "#10"
			    'listParams.AddByValue(diaFastener & sufix, 0.19 * convFactor, "in")
			    listParams.AddByValue("diaHole" & sufix, 0.161 * convFactor, "in")
				listParams.Item("diaHole" & sufix).Comment = "Tapped for .19in dia fastener (#10)"
				MsgBox("Parameters for holes have been generated")	
				
			Else If listParams.Item("HoleTableImp").Value = "#12"
			    'listParams.AddByValue(diaFastener & sufix, 0.216 * convFactor, "in")
			    listParams.AddByValue("diaHole" & sufix, 0.189 * convFactor, "in")
				listParams.Item("diaHole" & sufix).Comment = "Tapped for .216in dia fastener (#12)"
				MsgBox("Parameters for holes have been generated")		
				
			Else If listParams.Item("HoleTableImp").Value = "1/4"
			    'listParams.AddByValue(diaFastener & sufix, 0.25 * convFactor, "in")
			    listParams.AddByValue("diaHole" & sufix, 0.2188 * convFactor, "in")
				listParams.Item("diaHole" & sufix).Comment = "Tapped for .25in dia fastener (1/4)"
				MsgBox("Parameters for holes have been generated")		
				
			Else If listParams.Item("HoleTableImp").Value = "5/16"
			    'listParams.AddByValue(diaFastener & sufix, 0.3125 * convFactor, "in")
			    listParams.AddByValue("diaHole" & sufix, 0.277 * convFactor, "in")	
				listParams.Item("diaHole" & sufix).Comment = "Tapped for .3125in dia fastener (5/16)"
				MsgBox("Parameters for holes have been generated")
				
			Else If listParams.Item("HoleTableImp").Value = "3/8"
			    'listParams.AddByValue(diaFastener & sufix, 0.375 * convFactor, "in")
			    listParams.AddByValue("diaHole" & sufix, 0.332 * convFactor, "in") '0.397
				listParams.Item("diaHole" & sufix).Comment = "Tapped for .375in dia fastener (3/8)"
				
				MsgBox("Parameters for holes have been generated")				
			Else If listParams.Item("HoleTableImp").Value = "7/16"
			    'listParams.AddByValue(diaFastener & sufix, 0.4375 * convFactor, "in")
			    listParams.AddByValue("diaHole" & sufix, 0.3096 * convFactor, "in") '0.5938
				listParams.Item("diaHole" & sufix).Comment = "Tapped for .4375in dia fastener (7/16)"
				MsgBox("Parameters for holes have been generated")
				
			Else If listParams.Item("HoleTableImp").Value = "1/2"
			    'listParams.AddByValue(diaFastener & sufix, 0.5 * convFactor, "in")
			    listParams.AddByValue("diaHole" & sufix, 0.4531 * convFactor, "in") '0.5938
				listParams.Item("diaHole" & sufix).Comment = "Tapped for .5in dia fastener (1/2)"
				MsgBox("Parameters for holes have been generated")	
				
			Else If listParams.Item("HoleTableImp").Value = "9/16"
			    'listParams.AddByValue(diaFastener & sufix, 0.5625 * convFactor, "in")
			    listParams.AddByValue("diaHole" & sufix, 0.5156 * convFactor, "in") '0.5938
				listParams.Item("diaHole" & sufix).Comment = "Tapped for .5625in dia fastener (9/16)"
				MsgBox("Parameters for holes have been generated")
				
			Else If listParams.Item("HoleTableImp").Value = "5/8"
			    'listParams.AddByValue(diaFastener & sufix, 0.625 * convFactor, "in")
			    listParams.AddByValue("diaHole" & sufix, 0.5625 * convFactor, "In") '0.5938
				listParams.Item("diaHole" & sufix).Comment = "Tapped for .625in dia fastener (5/8)"
				MsgBox("Parameters for holes have been generated")		
				
			Else If listParams.Item("HoleTableImp").Value = "11/16"
			    'listParams.AddByValue(diaFastener & sufix, 0.6875 * convFactor, "in")
			    listParams.AddByValue("diaHole" & sufix, 0.6562 * convFactor, "in") '0.5938	
				listParams.Item("diaHole" & sufix).Comment = "Tapped for .6875in dia fastener (11/16)"
				MsgBox("Parameters for holes have been generated")			
				
			Else If listParams.Item("HoleTableImp").Value = "3/4"
			    'listParams.AddByValue(diaFastener & sufix, 0.75 * convFactor, "in")
			    listParams.AddByValue("diaHole" & sufix, 0.6875 * convFactor, "in") '0.5938	
				listParams.Item("diaHole" & sufix).Comment = "Tapped for .75in dia fastener (3/4)"
				MsgBox("Parameters for holes have been generated")	
				
			Else If listParams.Item("HoleTableImp").Value = "13/16"
			    'listParams.AddByValue(diaFastener & sufix, 0.8125 * convFactor, "in")
			    listParams.AddByValue("diaHole" & sufix, 0.7812 * convFactor, "in") '0.5938	
				listParams.Item("diaHole" & sufix).Comment = "Tapped for .8125in dia fastener (13/16)"
				MsgBox("Parameters for holes have been generated")		
				
			Else If listParams.Item("HoleTableImp").Value = "7/8"
			    'listParams.AddByValue(diaFastener & sufix, 0.875 * convFactor, "in")
			    listParams.AddByValue("diaHole" & sufix, 0.7969 * convFactor, "in") '0.5938	
				listParams.Item("diaHole" & sufix).Comment = "Tapped for .875in dia fastener (7/8)"
				MsgBox("Parameters for holes have been generated")			
				
			Else If listParams.Item("HoleTableImp").Value = "15/16"
			    'listParams.AddByValue(diaFastener & sufix, 0.9375 * convFactor, "in")
			    listParams.AddByValue("diaHole" & sufix, 0.9062 * convFactor, "in") '0.5938	
				listParams.Item("diaHole" & sufix).Comment = "Tapped for .9375 dia fastener (15/16)"
				MsgBox("Parameters for holes have been generated")	
				
			Else If listParams.Item("HoleTableImp").Value = "1"
			    'listParams.AddByValue(diaFastener & sufix, 1 * convFactor, "in")
			    listParams.AddByValue("diaHole" & sufix, 0.9219 * convFactor, "in") '0.5938
				listParams.Item("diaHole" & sufix).Comment = "Tapped for 1in dia fastener"
				MsgBox("Parameters for holes have been generated")

			 
			End If
		End If	
    End If
    
End Sub
