Sub Main ()
Dim esteDoc As PartDocument = Thisdoc.Document
Dim partDef As PartComponentDefinition = esteDoc.ComponentDefinition

Dim designTrackPropSet As PropertySet = esteDoc.PropertySets.Item("Design Tracking Properties")
Dim descProp As Object = designTrackPropSet.Item("Description")

sufix = InputBox("Enter Pattern feature name", "Pattern Feature Name", "ej:Face1holes")

descProp.Expression = sufix

Dim spacingValue As String = "spacingValue"
partDef.Parameters.UserParameters.AddByValue(spacingValue & sufix, 1, "in")

Dim lengthPattern As String = "lengthPattern"
partDef.Parameters.UserParameters.AddByValue(lengthPattern & sufix, 1, "in")

End Sub
