Attribute VB_Name = "Tabla_Astro_Num"
' ------------------------------------------------------
' Nombre:    Tabla_Astro_Num
' Tipo:      Módulo
' Propósito:
' Autor:     asalv
' Fecha:     15/01/2026
' ------------------------------------------------------

Public Sub CrearTablaNumerologia()

    Dim db As dao.Database
    Dim tdf As dao.TableDef
    Dim fld As dao.Field

    Set db = CurrentDb

    ' Si existe, la borramos para reconstruirla limpia
    On Error Resume Next
    db.TableDefs.Delete "tbmNum_Astro"
    On Error GoTo 0

    ' Crear tabla
    Set tdf = db.CreateTableDef("tbmNum_Astro")

    ' Campos
    tdf.Fields.Append tdf.CreateField("ID", dbLong)
    tdf.Fields("ID").Attributes = dbAutoIncrField

    tdf.Fields.Append tdf.CreateField("Reduccion", dbByte)
    tdf.Fields.Append tdf.CreateField("Arcano", dbText, 50)
    tdf.Fields.Append tdf.CreateField("TipoArcano", dbText, 20)
    tdf.Fields.Append tdf.CreateField("Palo", dbText, 20)
    tdf.Fields.Append tdf.CreateField("Signo", dbText, 20)
    tdf.Fields.Append tdf.CreateField("GradoInicio", dbByte)
    tdf.Fields.Append tdf.CreateField("GradoFin", dbByte)
    tdf.Fields.Append tdf.CreateField("Decanato", dbByte)
    tdf.Fields.Append tdf.CreateField("Elemento", dbText, 10)
    tdf.Fields.Append tdf.CreateField("Modalidad", dbText, 10)
    tdf.Fields.Append tdf.CreateField("Planeta", dbText, 20)
    tdf.Fields.Append tdf.CreateField("Estacion", dbText, 20)
    tdf.Fields.Append tdf.CreateField("Notas", dbMemo)

    ' Guardar tabla
    db.TableDefs.Append tdf

    MsgBox "Tabla creada correctamente.", vbInformation

End Sub

Public Sub InsertarFila( _
    ByVal id As Long, _
    ByVal Reduccion As Byte, _
    ByVal Arcano As String, _
    ByVal TipoArcano As String, _
    ByVal Palo As String, _
    ByVal Signo As String, _
    ByVal GradoInicio As Byte, _
    ByVal GradoFin As Byte, _
    ByVal Decanato As Byte, _
    ByVal Elemento As String, _
    ByVal Modalidad As String, _
    ByVal Planeta As String, _
    ByVal Estacion As String, _
    ByVal Notas As String)

    Dim db As dao.Database
    Dim rs As dao.Recordset

    Set db = CurrentDb
    Set rs = db.OpenRecordset("tbmNum_Astro", dbOpenDynaset)

    rs.AddNew
    rs!id = id
    rs!Reduccion = Reduccion
    rs!Arcano = Arcano
    rs!TipoArcano = TipoArcano
    rs!Palo = Palo
    rs!Signo = Signo
    rs!GradoInicio = GradoInicio
    rs!GradoFin = GradoFin
    rs!Decanato = Decanato
    rs!Elemento = Elemento
    rs!Modalidad = Modalidad
    rs!Planeta = Planeta
    rs!Estacion = Estacion
    rs!Notas = Notas
    rs.Update

    rs.Close
    Set rs = Nothing
    Set db = Nothing

End Sub

Public Sub InsertarArcanosMayores()

    ' Nº, Reduccion, Arcano, TipoArcano, Palo, Signo, GradoInicio, GradoFin, Decanato, Elemento, Modalidad, Planeta, Estacion, Notas

    InsertarFila 0, 0, "El Loco", "Mayor", "", "Plutón", 0, 0, 0, "Éter", "Libre", "Plutón", "", "Arcano fuera de la secuencia; representa el origen, el potencial y el espíritu."
    InsertarFila 1, 1, "El Mago", "Mayor", "", "Mercurio", 0, 0, 0, "Aire", "Mutable", "Mercurio", "", ""
    InsertarFila 2, 2, "La Papisa", "Mayor", "", "Luna", 0, 0, 0, "Agua", "Cardinal", "Luna", "", ""
    InsertarFila 3, 3, "La Emperatriz", "Mayor", "", "Venus", 0, 0, 0, "Tierra", "Cardinal", "Venus", "", ""
    InsertarFila 4, 4, "El Emperador", "Mayor", "", "Aries", 0, 0, 0, "Fuego", "Cardinal", "Marte", "", ""
    InsertarFila 5, 5, "El Hierofante", "Mayor", "", "Tauro", 0, 0, 0, "Tierra", "Fijo", "Venus", "", ""
    InsertarFila 6, 6, "Los Enamorados", "Mayor", "", "Géminis", 0, 0, 0, "Aire", "Mutable", "Mercurio", "", ""
    InsertarFila 7, 7, "El Carro", "Mayor", "", "Cáncer", 0, 0, 0, "Agua", "Cardinal", "Luna", "", ""
    InsertarFila 8, 8, "La Fuerza", "Mayor", "", "Leo", 0, 0, 0, "Fuego", "Fijo", "Sol", "", ""
    InsertarFila 9, 9, "El Ermitaño", "Mayor", "", "Virgo", 0, 0, 0, "Tierra", "Mutable", "Mercurio", "", ""
    InsertarFila 10, 1, "La Rueda de la Fortuna", "Mayor", "", "Júpiter", 0, 0, 0, "Fuego", "Mutable", "Júpiter", "", ""
    InsertarFila 11, 2, "La Justicia", "Mayor", "", "Libra", 0, 0, 0, "Aire", "Cardinal", "Venus", "", ""
    InsertarFila 12, 3, "El Ahorcado", "Mayor", "", "Neptuno", 0, 0, 0, "Agua", "Mutable", "Neptuno", "", ""
    InsertarFila 13, 4, "La Muerte", "Mayor", "", "Escorpión", 0, 0, 0, "Agua", "Fijo", "Plutón", "", ""
    InsertarFila 14, 5, "La Templanza", "Mayor", "", "Sagitario", 0, 0, 0, "Fuego", "Mutable", "Júpiter", "", ""
    InsertarFila 15, 6, "El Diablo", "Mayor", "", "Capricornio", 0, 0, 0, "Tierra", "Cardinal", "Saturno", "", ""
    InsertarFila 16, 7, "La Torre", "Mayor", "", "Marte", 0, 0, 0, "Fuego", "Cardinal", "Marte", "", ""
    InsertarFila 17, 8, "La Estrella", "Mayor", "", "Acuario", 0, 0, 0, "Aire", "Fijo", "Urano", "", ""
    InsertarFila 18, 9, "La Luna", "Mayor", "", "Piscis", 0, 0, 0, "Agua", "Mutable", "Neptuno", "", ""
    InsertarFila 19, 1, "El Sol", "Mayor", "", "Sol", 0, 0, 0, "Fuego", "Fijo", "Sol", "", ""
    InsertarFila 20, 2, "El Juicio", "Mayor", "", "Vulcano", 0, 0, 0, "Fuego", "Cardinal", "Vulcano", "", ""
    InsertarFila 21, 3, "El Mundo", "Mayor", "", "Saturno", 0, 0, 0, "Tierra", "Fijo", "Saturno", "", ""
    InsertarFila 22, 4, "El Loco", "Mayor", "", "Plutón", 0, 0, 0, "Agua", "Fijo", "Plutón", "", ""

    MsgBox "Arcanos Mayores insertados correctamente.", vbInformation

End Sub

Public Sub InsertarBastos()

    ' Nº, Reduccion, Arcano, TipoArcano, Palo, Signo, GradoInicio, GradoFin, Decanato, Elemento, Modalidad, Planeta, Estacion, Notas

    InsertarFila 23, 5, "Rey de Bastos", "Corte", "Bastos", "Aries", 0, 10, 1, "Fuego", "Cardinal", "Marte", "Primavera", ""
    InsertarFila 24, 6, "Reina de Bastos", "Corte", "Bastos", "Aries", 11, 20, 2, "Fuego", "Cardinal", "Marte", "Primavera", ""
    InsertarFila 25, 7, "Caballero de Bastos", "Corte", "Bastos", "Primavera", 0, 0, 0, "Fuego", "", "", "Primavera", ""
    InsertarFila 26, 8, "Sota de Bastos", "Corte", "Bastos", "Aries", 21, 30, 3, "Fuego", "Cardinal", "Marte", "Primavera", ""

    InsertarFila 27, 9, "As de Bastos", "As", "Bastos", "Fuego", 0, 0, 0, "Fuego", "", "", "", ""

    InsertarFila 28, 1, "Dos de Bastos", "Menor", "Bastos", "Aries", 0, 10, 1, "Fuego", "Cardinal", "Marte", "", ""
    InsertarFila 29, 11, "Tres de Bastos", "Menor", "Bastos", "Aries", 11, 20, 2, "Fuego", "Cardinal", "Marte", "", ""
    InsertarFila 30, 3, "Cuatro de Bastos", "Menor", "Bastos", "Aries", 21, 30, 3, "Fuego", "Cardinal", "Marte", "", ""

    InsertarFila 31, 4, "Cinco de Bastos", "Menor", "Bastos", "Leo", 0, 10, 1, "Fuego", "Fijo", "Sol", "", ""
    InsertarFila 32, 5, "Seis de Bastos", "Menor", "Bastos", "Leo", 11, 20, 2, "Fuego", "Fijo", "Sol", "", ""
    InsertarFila 33, 6, "Siete de Bastos", "Menor", "Bastos", "Leo", 21, 30, 3, "Fuego", "Fijo", "Sol", "", ""

    InsertarFila 34, 7, "Ocho de Bastos", "Menor", "Bastos", "Sagitario", 0, 10, 1, "Fuego", "Mutable", "Júpiter", "", ""
    InsertarFila 35, 8, "Nueve de Bastos", "Menor", "Bastos", "Sagitario", 11, 20, 2, "Fuego", "Mutable", "Júpiter", "", ""
    InsertarFila 36, 9, "Diez de Bastos", "Menor", "Bastos", "Sagitario", 21, 30, 3, "Fuego", "Mutable", "Júpiter", "", ""

    MsgBox "Bastos insertados correctamente.", vbInformation

End Sub

Public Sub InsertarCopas()

    ' Nº, Reduccion, Arcano, TipoArcano, Palo, Signo, GradoInicio, GradoFin, Decanato, Elemento, Modalidad, Planeta, Estacion, Notas

    InsertarFila 37, 1, "Rey de Copas", "Corte", "Copas", "Cáncer", 0, 10, 1, "Agua", "Cardinal", "Luna", "Verano", ""
    InsertarFila 38, 11, "Reina de Copas", "Corte", "Copas", "Cáncer", 11, 20, 2, "Agua", "Cardinal", "Luna", "Verano", ""
    InsertarFila 39, 3, "Caballero de Copas", "Corte", "Copas", "Verano", 0, 0, 0, "Agua", "", "", "Verano", ""
    InsertarFila 40, 4, "Sota de Copas", "Corte", "Copas", "Cáncer", 21, 30, 3, "Agua", "Cardinal", "Luna", "Verano", ""

    InsertarFila 41, 5, "As de Copas", "As", "Copas", "Agua", 0, 0, 0, "Agua", "", "", "", ""

    InsertarFila 42, 6, "Dos de Copas", "Menor", "Copas", "Cáncer", 0, 10, 1, "Agua", "Cardinal", "Luna", "", ""
    InsertarFila 43, 7, "Tres de Copas", "Menor", "Copas", "Cáncer", 11, 20, 2, "Agua", "Cardinal", "Luna", "", ""
    InsertarFila 44, 8, "Cuatro de Copas", "Menor", "Copas", "Cáncer", 21, 30, 3, "Agua", "Cardinal", "Luna", "", ""

    InsertarFila 45, 9, "Cinco de Copas", "Menor", "Copas", "Escorpión", 0, 10, 1, "Agua", "Fijo", "Plutón", "", ""
    InsertarFila 46, 1, "Seis de Copas", "Menor", "Copas", "Escorpión", 11, 20, 2, "Agua", "Fijo", "Plutón", "", ""
    InsertarFila 47, 11, "Siete de Copas", "Menor", "Copas", "Escorpión", 21, 30, 3, "Agua", "Fijo", "Plutón", "", ""

    InsertarFila 48, 3, "Ocho de Copas", "Menor", "Copas", "Piscis", 0, 10, 1, "Agua", "Mutable", "Neptuno", "", ""
    InsertarFila 49, 4, "Nueve de Copas", "Menor", "Copas", "Piscis", 11, 20, 2, "Agua", "Mutable", "Neptuno", "", ""
    InsertarFila 50, 5, "Diez de Copas", "Menor", "Copas", "Piscis", 21, 30, 3, "Agua", "Mutable", "Neptuno", "", ""

    MsgBox "Copas insertadas correctamente.", vbInformation

End Sub

Public Sub InsertarEspadas()

    ' Nº, Reduccion, Arcano, TipoArcano, Palo, Signo, GradoInicio, GradoFin, Decanato, Elemento, Modalidad, Planeta, Estacion, Notas

    InsertarFila 51, 6, "Rey de Espadas", "Corte", "Espadas", "Libra", 0, 10, 1, "Aire", "Cardinal", "Venus", "Otoño", ""
    InsertarFila 52, 7, "Reina de Espadas", "Corte", "Espadas", "Libra", 11, 20, 2, "Aire", "Cardinal", "Venus", "Otoño", ""
    InsertarFila 53, 8, "Caballero de Espadas", "Corte", "Espadas", "Otoño", 0, 0, 0, "Aire", "", "", "Otoño", ""
    InsertarFila 54, 9, "Sota de Espadas", "Corte", "Espadas", "Libra", 21, 30, 3, "Aire", "Cardinal", "Venus", "Otoño", ""

    InsertarFila 55, 1, "As de Espadas", "As", "Espadas", "Aire", 0, 0, 0, "Aire", "", "", "", ""

    InsertarFila 56, 11, "Dos de Espadas", "Menor", "Espadas", "Libra", 0, 10, 1, "Aire", "Cardinal", "Venus", "", ""
    InsertarFila 57, 3, "Tres de Espadas", "Menor", "Espadas", "Libra", 11, 20, 2, "Aire", "Cardinal", "Venus", "", ""
    InsertarFila 58, 4, "Cuatro de Espadas", "Menor", "Espadas", "Libra", 21, 30, 3, "Aire", "Cardinal", "Venus", "", ""

    InsertarFila 59, 5, "Cinco de Espadas", "Menor", "Espadas", "Acuario", 0, 10, 1, "Aire", "Fijo", "Urano", "", ""
    InsertarFila 60, 6, "Seis de Espadas", "Menor", "Espadas", "Acuario", 11, 20, 2, "Aire", "Fijo", "Urano", "", ""
    InsertarFila 61, 7, "Siete de Espadas", "Menor", "Espadas", "Acuario", 21, 30, 3, "Aire", "Fijo", "Urano", "", ""

    InsertarFila 62, 8, "Ocho de Espadas", "Menor", "Espadas", "Géminis", 0, 10, 1, "Aire", "Mutable", "Mercurio", "", ""
    InsertarFila 63, 9, "Nueve de Espadas", "Menor", "Espadas", "Géminis", 11, 20, 2, "Aire", "Mutable", "Mercurio", "", ""
    InsertarFila 64, 1, "Diez de Espadas", "Menor", "Espadas", "Géminis", 21, 30, 3, "Aire", "Mutable", "Mercurio", "", ""

    MsgBox "Espadas insertadas correctamente.", vbInformation

End Sub

Public Sub InsertarOros()

    ' Nº, Reduccion, Arcano, TipoArcano, Palo, Signo, GradoInicio, GradoFin, Decanato, Elemento, Modalidad, Planeta, Estacion, Notas

    InsertarFila 65, 11, "Rey de Oros", "Corte", "Oros", "Capricornio", 0, 10, 1, "Tierra", "Cardinal", "Saturno", "Invierno", ""
    InsertarFila 66, 3, "Reina de Oros", "Corte", "Oros", "Capricornio", 11, 20, 2, "Tierra", "Cardinal", "Saturno", "Invierno", ""
    InsertarFila 67, 4, "Caballero de Oros", "Corte", "Oros", "Invierno", 0, 0, 0, "Tierra", "", "", "Invierno", ""
    InsertarFila 68, 5, "Sota de Oros", "Corte", "Oros", "Capricornio", 21, 30, 3, "Tierra", "Cardinal", "Saturno", "Invierno", ""

    InsertarFila 69, 6, "As de Oros", "As", "Oros", "Tierra", 0, 0, 0, "Tierra", "", "", "", ""

    InsertarFila 70, 7, "Dos de Oros", "Menor", "Oros", "Capricornio", 0, 10, 1, "Tierra", "Cardinal", "Saturno", "", ""
    InsertarFila 71, 8, "Tres de Oros", "Menor", "Oros", "Capricornio", 11, 20, 2, "Tierra", "Cardinal", "Saturno", "", ""
    InsertarFila 72, 9, "Cuatro de Oros", "Menor", "Oros", "Capricornio", 21, 30, 3, "Tierra", "Cardinal", "Saturno", "", ""

    InsertarFila 73, 1, "Cinco de Oros", "Menor", "Oros", "Tauro", 0, 10, 1, "Tierra", "Fijo", "Venus", "", ""
    InsertarFila 74, 11, "Seis de Oros", "Menor", "Oros", "Tauro", 11, 20, 2, "Tierra", "Fijo", "Venus", "", ""
    InsertarFila 75, 3, "Siete de Oros", "Menor", "Oros", "Tauro", 21, 30, 3, "Tierra", "Fijo", "Venus", "", ""

    InsertarFila 76, 4, "Ocho de Oros", "Menor", "Oros", "Virgo", 0, 10, 1, "Tierra", "Mutable", "Mercurio", "", ""
    InsertarFila 77, 5, "Nueve de Oros", "Menor", "Oros", "Virgo", 11, 20, 2, "Tierra", "Mutable", "Mercurio", "", ""
    InsertarFila 78, 6, "Diez de Oros", "Menor", "Oros", "Virgo", 21, 30, 3, "Tierra", "Mutable", "Mercurio", "", ""

    MsgBox "Oros insertados correctamente.", vbInformation

End Sub


'Public Sub Actualizar_tbuResultados()
'
'    Dim db As DAO.Database
'    Dim tdf As DAO.TableDef
'    Dim fld As DAO.Field
'
'    Set db = CurrentDb
'    Set tdf = db.TableDefs("tbuResultados")
'
'    '--- Función local para añadir campo si no existe ---
'    Dim AddField As Object
'    Set AddField = CreateObject("Scripting.Dictionary")
'
'    AddField("IDAnalisis") = dbLong
'    AddField("PlanoFisico") = dbText
'    AddField("PlanoEmocional") = dbText
'    AddField("PlanoMental") = dbText
'    AddField("PlanoIntuitivo") = dbText
'    AddField("PiedraAngular") = dbText
'    AddField("PiedraToque") = dbText
'    AddField("PrimeraLetra") = dbText
'    AddField("PrimeraVocal") = dbText
'    AddField("PrimeraConsonante") = dbText
'    AddField("RespuestaSubconsciente") = dbText
'    AddField("Poder") = dbText
'    AddField("DeudaKarmica") = dbText
'
'    Dim key As Variant
'    For Each key In AddField.Keys
'        On Error Resume Next
'        Set fld = tdf.Fields(key)
'        On Error GoTo 0
'
'        If fld Is Nothing Then
'            tdf.Fields.Append tdf.CreateField(key, AddField(key), 50)
'        End If
'
'        Set fld = Nothing
'    Next key
'
'    MsgBox "tbuResultados actualizado correctamente.", vbInformation
'
'End Sub

'Public Sub CrearTabla_tbuInclusion()
'
'    Dim db As DAO.Database
'    Dim tdf As DAO.TableDef
''    Dim fld As DAO.Field
'
'    Set db = CurrentDb
'
'    ' Si la tabla existe, salir sin hacer nada
'    On Error Resume Next
'    Set tdf = db.TableDefs("tbuInclusion")
'    On Error GoTo 0
'
'    If Not tdf Is Nothing Then
'        MsgBox "La tabla tbuInclusion ya existe.", vbInformation
'        Exit Sub
'    End If
'
'    ' Crear la tabla
'    Set tdf = db.CreateTableDef("tbuInclusion")
'
'    ' Campos principales
'    tdf.Fields.Append tdf.CreateField("IDFonetica", dbLong)
'    tdf.Fields("IDFonetica").Attributes = dbAutoIncrField
'
'    tdf.Fields.Append tdf.CreateField("IDResultado", dbLong)
'    tdf.Fields.Append tdf.CreateField("IDPersona", dbLong)
'
'    ' Campos N1 a N9 tipo Byte
'    Dim i As Integer
'    For i = 1 To 9
'        tdf.Fields.Append tdf.CreateField("N" & i, dbByte)
'    Next i
'
'    ' Añadir tabla a la base de datos
'    db.TableDefs.Append tdf
'
'    MsgBox "Tabla tbuFoneticaResumen creada correctamente.", vbInformation
'
'End Sub

'Public Sub CrearRelacion_Inclusion_Resultados()
'
'    Dim db As DAO.Database
'    Dim rel As DAO.Relation
'    Dim fld As DAO.Field
'
'    Set db = CurrentDb
'
'    ' Eliminar relación previa si existe
'    On Error Resume Next
'    db.Relations.Delete "rel_Inclusion_Resultados"
'    On Error GoTo 0
'
'    ' Crear relación
'    Set rel = db.CreateRelation("rel_Inclusion_Resultados", _
'                                "tbuResultados", "tbuInclusion", _
'                                dbRelationUpdateCascade + dbRelationDeleteCascade)
'
'    Set fld = rel.CreateField("IDResultado")
'    fld.ForeignName = "IDResultado"
'    rel.Fields.Append fld
'
'    db.Relations.Append rel
'
'    MsgBox "Relación creada correctamente.", vbInformation
'
'End Sub

