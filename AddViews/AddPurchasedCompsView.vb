Sub Main ()
    ' Verificar que el documento activo sea un ensamblaje
    Dim oDoc As AssemblyDocument
    If ThisApplication.ActiveDocument.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
        MessageBox.Show("Esta regla solo se puede ejecutar desde un archivo de Ensamblaje (.iam).", "Error de Tipo de Archivo", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Exit Sub
    End If
    
    oDoc = ThisApplication.ActiveDocument
    Dim oAsmDef As AssemblyComponentDefinition = oDoc.ComponentDefinition
    
    ' --- GESTIÓN DE LA REPRESENTACIÓN DE VISTA ---
    Dim oViewRep As DesignViewRepresentation
    Dim nombreVista As String = "Purchased"
    
    Try
        ' Intentar activar la vista si ya existe
        oViewRep = oAsmDef.RepresentationsManager.DesignViewRepresentations.Item(nombreVista)
        oViewRep.Activate()
    Catch
        ' Si no existe (da error), la creamos desde cero
        oViewRep = oAsmDef.RepresentationsManager.DesignViewRepresentations.Add(nombreVista)
        oViewRep.Activate()
    End Try
    
    ' Bloquear la vista para que los cambios de visibilidad no se pierdan al guardar
    oViewRep.Locked = False 
    ' ----------------------------------------------
    
    ' Obtener todas las ocurrencias hoja
    Dim oLeafOccs As ComponentOccurrencesEnumerator = oAsmDef.Occurrences.AllLeafOccurrences
    Dim oOcc As ComponentOccurrence
    
    ' Iniciar transacción para unificar el historial de cambios (Undo)
    Dim oTransaction As Transaction = ThisApplication.TransactionManager.StartTransaction(oDoc, "Filtrar por Comprados en Vista")
    
    Try
        ' Recorrer cada ocurrencia hoja
        For Each oOcc In oLeafOccs
            ' Omitir componentes virtuales
            If TypeOf oOcc.Definition Is VirtualComponentDefinition Then Continue For
            
            ' Verificar si la estructura del BOM NO es "Purchased" (Comprado)
            If oOcc.BOMStructure <> BOMStructureEnum.kPurchasedBOMStructure Then
                If oOcc.Visible = True Then
                    oOcc.Visible = False
                End If
            Else
                ' Si ES "Purchased", garantizamos que sea visible
                If oOcc.Visible = False Then
                    oOcc.Visible = True
                End If
            End If
        Next
        
        ' Opcional: Puedes bloquear la vista al final para que nadie la modifique por error
        ' oViewRep.Locked = True
        
        ' Confirmar los cambios de la transacción
        oTransaction.End()
        
        ' Actualizar la pantalla gráfica global
        ThisApplication.ActiveView.Update()
        
    Catch ex As Exception
        ' En caso de error, cancelar los cambios parciales
        oTransaction.Abort()
        MessageBox.Show("Ocurrió un error al procesar las ocurrencias: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
    End Try
End Sub
