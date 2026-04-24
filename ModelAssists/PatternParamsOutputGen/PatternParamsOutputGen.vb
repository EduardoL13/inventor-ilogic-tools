Sub Main ()
	
	Dim esteDoc As PartDocument = ThisDoc.Document
	Dim partDef As PartComponentDefinition = esteDoc.ComponentDefinition

	Dim designTrackPropSet As PropertySet = esteDoc.PropertySets.Item("Design Tracking Properties")
	Dim descProp As Object = designTrackPropSet.Item("Description")
	Dim sufix As String = descProp.Value
	
	Dim rawNoElements As String = "rawNoElements" & sufix
	'MsgBox("lengthPattern" & sufix & "/spacingValue" & sufix)
	partDef.Parameters.UserParameters.AddByExpression(rawNoElements, "lengthPattern" & sufix & "/spacingValue" & sufix, "ul")
	
	Dim NoElements As String = "NoElements" & sufix
	partDef.Parameters.UserParameters.AddByExpression(NoElements, "floor(" & rawNoElements & ")+1", "ul")
	
	Dim widthEffective As String = "widthEffective" & sufix
	partDef.Parameters.UserParameters.AddByExpression(widthEffective,"floor(rawNoElements" & sufix &")" & "* spacingValue" & sufix,"in")
	
	Dim distEnds As String = "distEnds" & sufix
	partDef.Parameters.UserParameters.AddByExpression(distEnds, "(lengthPattern" & sufix & "-widthEffective" & sufix & ")/2", "in")

End Sub
