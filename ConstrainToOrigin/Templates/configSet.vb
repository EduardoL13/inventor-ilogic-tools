Sub Main ()
    
	Dim esteDoc As PartDocument = ThisDoc.Document
    Dim featureList As PartFeatures = esteDoc.ComponentDefinition.Features 
    

	'Changes configuration

    If currentModel = "G3" Then
	' Listar Supresiones y cambios en kep parameters
        Parameter("kepWidth") = Parameter("inputThkG3")
	    Parameter("input2LengthOffsetSides") = Parameter("inputLengthOffsetSidesG3")
		Parameter("input2NoScreens") = Parameter("inputNoScreensPerFaceG3")
		Parameter("kepOffsetFromCenterVT") = Parameter("inputOffsetFromCenterVT")
		'Parameter("lengthGapBetScreens") = 0
		'featureList("ThkExtSrfcSupSheetB").Suppressed = False
		'featureList("HoleForWoBBack").Suppressed = False
		
		
    Else If currentModel = "D3_2S" Then
		
	    Parameter("kepWidth") = Parameter("inputThkD3")
		Parameter("input2LengthOffsetSides") = Parameter("inputLengthOffsetSides2S")
		Parameter("input2NoScreens") = Parameter("inputNoScreensPerFaceD32S")
		Parameter("kepOffsetFromCenterVT") = 0
        'Parameter("lengthGapBetScreens") = 0
		'featureList("ThkExtSrfcSupSheetB").Suppressed = True
		'featureList("HoleForWoBBack").Suppressed = True
		
		
	Else If currentModel = "D3_4S" Then
		
		Parameter("kepWidth") = Parameter("inputThkD3")
		Parameter("input2LengthOffsetSides") = Parameter("inputLengthOffsetSides2S")
		Parameter("input2NoScreens") = Parameter("inputNoScreensPerFaceD34S")
		Parameter("kepOffsetFromCenterVT") = 0
		
		'featureList("ThkExtSrfcSupSheetB").Suppressed = True
		'featureList("HoleForWoBBack").Suppressed = True
		
	End If


End Sub
