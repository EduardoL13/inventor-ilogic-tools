Sub Main ()
    
	'Set Parameters and features access
	'MsgBox("Conf triggered")
	Dim currentDoc As AssemblyDocument = ThisDoc.Document
    Dim patternsList As OccurrencePatterns = currentDoc.ComponentDefinition.OccurrencePatterns
	
	'Expressions to add if required
    'Expression = "(kepWidth + noSigns * kepExcessTop + ( noSigns - 1 ul ) * kepGapBetSigns)/2"
	
	'Configuration list
	configuration1 = "URP75"
	configuration2 = "URP552"
	
	
	'Changes configuration
	

    If parameter("currentScreenModel") = configuration1 Then
	' Listar Supresiones y cambios en kep parameters
'        Parameter("kepWidth") = 84
'		featureList("MirrLidAtTop").Suppressed = True
        patternsList("PatternScreensURP75").Unsuppress
		patternsList("PatternScreensURP552").Suppress
		MsgBox("URP75 conf activa")
		
    Else If parameter("currentScreenModel") = configuration2 Then
		
        patternsList("PatternScreensURP75").Suppress
		patternsList("PatternScreensURP552").Unsuppress
		MsgBox("URP552 conf activa")
		
	End If


End Sub
