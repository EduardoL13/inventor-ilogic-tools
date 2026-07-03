Sub Main ()
    
	'Set Parameters and features access
	Dim currentDoc As PartDocument = ThisDoc.Document
    Dim featureList As PartFeatures = currentDoc.ComponentDefinition.Features 
    Dim mParamsList As ModelParameters = currentDoc.ComponentDefinition.Parameters.ModelParameters
	
	'Expressions to add if required
    Expression = "(kepWidth + noSigns * kepExcessTop + ( noSigns - 1 ul ) * kepGapBetSigns)/2"
	
	'Configuration list
	configuration1 = "22B"
	configuration2 = "23B"
	configuration3 = "24C"
	
	
	'Changes configuration
	

    If currentModel = configuration1 Then
	' Listar Supresiones y cambios en kep parameters
'        Parameter("kepWidth") = 84
'		Parameter("kepHeight") = 24.38
'		Parameter("noSigns") = 2
'		mParamsList.Item("mOffsetStartPlaneProfile").Expression = ExpForOffset
'		featureList("MirrLidAtTop").Suppressed = True		
		'Parameter("lengthGapBetScreens") = 0

		'featureList("HoleForWoBBack").Suppressed = False
		
    Else If currentModel = configuration2 Then
		
'	    Parameter("kepWidth") = 84
'		Parameter("kepHeight") = 27.88
'		Parameter("noSigns") = 2        
'		mParamsList.Item("mOffsetStartPlaneProfile").Expression = ExpForOffset
'		featureList("MirrLidAtTop").Suppressed = True
		
    Else If currentModel = configuration3 Then
		
'	    Parameter("kepWidth") = 84
'		Parameter("kepHeight") = 22
'		Parameter("noSigns") = 1		
'		mParamsList.Item("mOffsetStartPlaneProfile").Expression = "0"
'		featureList("MirrLidAtTop").Suppressed = False
		
	End If


End Sub
