Sub Main
	
	Dim oDoc As PartDocument = ThisDoc.Document
	Dim uParams As UserParameters = oDoc.ComponentDefinition.Parameters.UserParameters
	Dim thkParam As UserParameter = uParams.Item("pThk")
	
	Dim nameStyle As String = "test"
	Dim nameMaterial As String = "Aluminum 6061" ' Existing material in the library
	
	Dim thkStyle As String = thkParam.Name ' 0.125' En centímetros (ej. 0.05 cm = 0.5 mm)
	'convFac = 2.54 ' in to cm
	'thkStyle = thkStyle * convFac
	
	Dim type2Bend As Object = BendReliefShapeEnum.kRoundBendReliefShape ' 32770 = Round (Redondo), 32769 = Straight, 32771 = Tear
	Dim type3Bend As Object = CornerReliefShapeEnum.kRoundCornerReliefShape
	Dim multiplierRadiusBend As Double = 1.5
	Dim multiplierWidthRelief As Double = 1
	Dim mutiplierDepthRelief As Double = 2.5
	
	Dim radiusBend As String = thkStyle & "*" & multiplierRadiusBend.ToString() 	
	Dim depthRelief As String = thkStyle & "*" & mutiplierDepthRelief.ToString()
	Dim widthRelief As String = thkStyle & "*" & multiplierWidthRelief.ToString() 
    
	If oDoc.DocumentType <> DocumentTypeEnum.kPartDocumentObject OrElse oDoc.SubType <> "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}" Then
		MessageBox.Show("The part is not of the type sheet metal", "Error iLogic")
		Exit Sub
	End If

	Dim oCompDef As SheetMetalComponentDefinition = oDoc.ComponentDefinition
	Dim oStyles As SheetMetalStyles = oCompDef.SheetMetalStyles

	Dim oStyle As SheetMetalStyle
	Try
		oStyle = oStyles.Item(nameStyle)
	Catch
		Dim oDefaultStyle As SheetMetalStyle = oCompDef.ActiveSheetMetalStyle
		oStyle = oDefaultStyle.Copy(nameStyle)
		MessageBox.Show(" New sheet metal default has been added" & estiloNombre, "iLogic")
	End Try

	'Material?
	oStyle.Thickness = thkStyle'thkStyle.ToString()
	
	oStyle.BendRadius = radiusBend.ToString()

	oStyle.BendReliefShape = type2Bend
	oStyle.BendReliefDepth = depthRelief
	oStyle.BendReliefWidth = widthRelief

	oStyle.CornerReliefShape = type3Bend
	'oStyle.CornerReliefDepth = depthRelief.ToString()
	'oStyle.CornerReliefWidth = widthRelief.ToString()
	
	oDoc.Update
	MessageBox.Show("Style has been added", "ilogic")
	
End Sub
