Public Sub CrearTablaTarot()

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    
    Set db = CurrentDb
    
    ' Si existe, eliminarla
    On Error Resume Next
    db.TableDefs.Delete "tblTarotCartas"
    On Error GoTo 0
    
    ' Crear tabla
    Set tdf = db.CreateTableDef("tblTarotCartas")
    
    ' Campo clave primaria
    Set fld = tdf.CreateField("NumeroGlobal", dbByte)
    fld.Attributes = dbFixedField
    tdf.Fields.Append fld
    tdf.Fields("NumeroGlobal").Attributes = dbPrimaryKey
    
    ' Otros campos
    tdf.Fields.Append tdf.CreateField("NombreCarta", dbText, 100)
    tdf.Fields.Append tdf.CreateField("Palo", dbText, 20)
    tdf.Fields.Append tdf.CreateField("Elemento", dbText, 20)
    tdf.Fields.Append tdf.CreateField("NumeroInterno", dbByte)
    tdf.Fields.Append tdf.CreateField("Figura", dbText, 20)
    tdf.Fields.Append tdf.CreateField("RutaMarkdown", dbText, 255)
    
    ' Añadir tabla a la BD
    db.TableDefs.Append tdf
    
    MsgBox "Tabla tblTarotCartas creada correctamente.", vbInformation

End Sub
