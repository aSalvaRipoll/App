Option Compare Database
Option Explicit

' =============================================================================
' Módulo: modCrearTablasAstrologia
' Descripción: Crea las tablas de correspondencias astrológicas
' Autor: Alba - Sistema de Numerología y Tarot
' =============================================================================

Public Sub CrearTodasTablasAstrologia()
    ' Procedimiento principal que crea todas las tablas
    
    On Error GoTo ErrorHandler
    
    Debug.Print "=========================================="
    Debug.Print "CREANDO TABLAS DE ASTROLOGÍA"
    Debug.Print "=========================================="
    Debug.Print ""
    
    ' Eliminar tablas existentes si las hay (para recrear limpias)
    Call EliminarTablasAstrologiaExistentes
    
    ' Crear tablas en orden correcto (respetando FK)
    Call CrearTablaSignosZodiacales
    Call CrearTablaPlanetas
    Call CrearTablaDecanatos
    Call CrearTablaArcanosMayoresAstrologia
    Call CrearTablaArcanosMenoresNumeradosAstrologia
    Call CrearTablaFigurasCorteAstrologia
    Call CrearTablaNumerosAstrologia
    Call CrearTablaElementos
    Call CrearTablaCualidadesZodiacales
    
    Debug.Print ""
    Debug.Print "=========================================="
    Debug.Print "TODAS LAS TABLAS CREADAS EXITOSAMENTE"
    Debug.Print "=========================================="
    
    MsgBox "Tablas de Astrología creadas exitosamente." & vbCrLf & vbCrLf & _
           "Ahora ejecute: PoblarTodasTablasAstrologia()", _
           vbInformation, "Tablas Creadas"
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "ERROR al crear tablas: " & Err.Description
    MsgBox "Error al crear tablas: " & Err.Description, vbCritical
End Sub

Private Sub EliminarTablasAstrologiaExistentes()
    ' Elimina tablas si existen (para poder recrearlas)
    
    On Error Resume Next
    
    Dim db As DAO.Database
    Set db = CurrentDb
    
    Debug.Print "Eliminando tablas existentes..."
    
    ' Eliminar en orden inverso a creación (por FK)
    db.TableDefs.Delete "tblCualidadesZodiacales"
    db.TableDefs.Delete "tblElementos"
    db.TableDefs.Delete "tblNumerosAstrologia"
    db.TableDefs.Delete "tblFigurasCorteAstrologia"
    db.TableDefs.Delete "tblArcanosMenoresNumeradosAstrologia"
    db.TableDefs.Delete "tblArcanosMayoresAstrologia"
    db.TableDefs.Delete "tblDecanatos"
    db.TableDefs.Delete "tblPlanetas"
    db.TableDefs.Delete "tblSignosZodiacales"
    
    Debug.Print "Tablas anteriores eliminadas (si existían)"
    
    Set db = Nothing
    On Error GoTo 0
End Sub

Private Sub CrearTablaSignosZodiacales()
    ' Crea tabla de los 12 signos zodiacales
    
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim tbl As DAO.TableDef
    Dim fld As DAO.Field
    
    Set db = CurrentDb
    Set tbl = db.CreateTableDef("tblSignosZodiacales")
    
    ' Campos
    Set fld = tbl.CreateField("ID_Signo", dbLong)
    fld.Attributes = dbAutoIncrField
    tbl.Fields.Append fld
    
    tbl.Fields.Append tbl.CreateField("NombreSigno", dbText, 20)
    tbl.Fields.Append tbl.CreateField("SimboloUnicode", dbText, 5)
    tbl.Fields.Append tbl.CreateField("Elemento", dbText, 10)
    tbl.Fields.Append tbl.CreateField("Cualidad", dbText, 15)
    tbl.Fields.Append tbl.CreateField("Polaridad", dbText, 10)
    tbl.Fields.Append tbl.CreateField("PlanetaRegente", dbText, 20)
    tbl.Fields.Append tbl.CreateField("PlanetaExaltacion", dbText, 20)
    tbl.Fields.Append tbl.CreateField("PlanetaDetrimento", dbText, 20)
    tbl.Fields.Append tbl.CreateField("PlanetaCaida", dbText, 20)
    tbl.Fields.Append tbl.CreateField("CasaNatural", dbByte)
    tbl.Fields.Append tbl.CreateField("GradoInicial", dbInteger)
    tbl.Fields.Append tbl.CreateField("GradoFinal", dbInteger)
    tbl.Fields.Append tbl.CreateField("FechaInicioAprox", dbText, 10)
    tbl.Fields.Append tbl.CreateField("FechaFinAprox", dbText, 10)
    tbl.Fields.Append tbl.CreateField("ArcanosAsociados", dbText, 100)
    tbl.Fields.Append tbl.CreateField("CaracteristicasClave", dbMemo)
    tbl.Fields.Append tbl.CreateField("PalabrasClave", dbMemo)
    
    ' Clave primaria
    Dim idx As DAO.Index
    Set idx = tbl.CreateIndex("PrimaryKey")
    idx.Primary = True
    idx.Required = True
    Set fld = idx.CreateField("ID_Signo")
    idx.Fields.Append fld
    tbl.Indexes.Append idx
    
    ' Agregar tabla
    db.TableDefs.Append tbl
    
    Debug.Print "? Tabla tblSignosZodiacales creada"
    
    Set fld = Nothing
    Set idx = Nothing
    Set tbl = Nothing
    Set db = Nothing
    Exit Sub
    
ErrorHandler:
    Debug.Print "ERROR creando tblSignosZodiacales: " & Err.Description
    Resume Next
End Sub

Private Sub CrearTablaPlanetas()
    ' Crea tabla de planetas y cuerpos celestes
    
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim tbl As DAO.TableDef
    Dim fld As DAO.Field
    
    Set db = CurrentDb
    Set tbl = db.CreateTableDef("tblPlanetas")
    
    ' Campos
    Set fld = tbl.CreateField("ID_Planeta", dbLong)
    fld.Attributes = dbAutoIncrField
    tbl.Fields.Append fld
    
    tbl.Fields.Append tbl.CreateField("NombrePlaneta", dbText, 20)
    tbl.Fields.Append tbl.CreateField("SimboloUnicode", dbText, 5)
    tbl.Fields.Append tbl.CreateField("TipoCuerpo", dbText, 30)
    tbl.Fields.Append tbl.CreateField("SignoDomicilio", dbText, 50)
    tbl.Fields.Append tbl.CreateField("SignoExaltacion", dbText, 20)
    tbl.Fields.Append tbl.CreateField("SignoDetrimento", dbText, 50)
    tbl.Fields.Append tbl.CreateField("SignoCaida", dbText, 20)
    tbl.Fields.Append tbl.CreateField("ArcanosAsociados", dbText, 100)
    tbl.Fields.Append tbl.CreateField("NumeroNumerologico", dbByte)
    tbl.Fields.Append tbl.CreateField("ElementoAfin", dbText, 10)
    tbl.Fields.Append tbl.CreateField("CaracteristicasClave", dbMemo)
    tbl.Fields.Append tbl.CreateField("PalabrasClave", dbMemo)
    tbl.Fields.Append tbl.CreateField("CicloOrbital", dbText, 50)
    
    ' Clave primaria
    Dim idx As DAO.Index
    Set idx = tbl.CreateIndex("PrimaryKey")
    idx.Primary = True
    idx.Required = True
    Set fld = idx.CreateField("ID_Planeta")
    idx.Fields.Append fld
    tbl.Indexes.Append idx
    
    db.TableDefs.Append tbl
    
    Debug.Print "? Tabla tblPlanetas creada"
    
    Set fld = Nothing
    Set idx = Nothing
    Set tbl = Nothing
    Set db = Nothing
    Exit Sub
    
ErrorHandler:
    Debug.Print "ERROR creando tblPlanetas: " & Err.Description
    Resume Next
End Sub

Private Sub CrearTablaDecanatos()
    ' Crea tabla de 36 decanatos
    
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim tbl As DAO.TableDef
    Dim fld As DAO.Field
    
    Set db = CurrentDb
    Set tbl = db.CreateTableDef("tblDecanatos")
    
    ' Campos
    Set fld = tbl.CreateField("ID_Decanato", dbLong)
    fld.Attributes = dbAutoIncrField
    tbl.Fields.Append fld
    
    tbl.Fields.Append tbl.CreateField("Signo_ID", dbLong)
    tbl.Fields.Append tbl.CreateField("NumeroDecanato", dbByte)
    tbl.Fields.Append tbl.CreateField("GradoInicial", dbInteger)
    tbl.Fields.Append tbl.CreateField("GradoFinal", dbInteger)
    tbl.Fields.Append tbl.CreateField("PlanetaSubregente", dbText, 20)
    tbl.Fields.Append tbl.CreateField("CartaTarotNumero", dbByte)
    tbl.Fields.Append tbl.CreateField("CartaTarotPalo", dbText, 20)
    tbl.Fields.Append tbl.CreateField("CartaTarotCompleta", dbText, 50)
    tbl.Fields.Append tbl.CreateField("ElementoTriplicidad", dbText, 10)
    tbl.Fields.Append tbl.CreateField("InterpretacionClave", dbMemo)
    tbl.Fields.Append tbl.CreateField("PalabrasClave", dbMemo)
    
    ' Clave primaria
    Dim idx As DAO.Index
    Set idx = tbl.CreateIndex("PrimaryKey")
    idx.Primary = True
    idx.Required = True
    Set fld = idx.CreateField("ID_Decanato")
    idx.Fields.Append fld
    tbl.Indexes.Append idx
    
    db.TableDefs.Append tbl
    
    Debug.Print "? Tabla tblDecanatos creada"
    
    Set fld = Nothing
    Set idx = Nothing
    Set tbl = Nothing
    Set db = Nothing
    Exit Sub
    
ErrorHandler:
    Debug.Print "ERROR creando tblDecanatos: " & Err.Description
    Resume Next
End Sub

Private Sub CrearTablaArcanosMayoresAstrologia()
    ' Crea tabla de correspondencias astrológicas para 22 Arcanos Mayores
    
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim tbl As DAO.TableDef
    Dim fld As DAO.Field
    
    Set db = CurrentDb
    Set tbl = db.CreateTableDef("tblArcanosMayoresAstrologia")
    
    Set fld = tbl.CreateField("ID_Arcano", dbLong)
    fld.Attributes = dbAutoIncrField
    tbl.Fields.Append fld
    
    tbl.Fields.Append tbl.CreateField("NumeroArcano", dbByte)
    tbl.Fields.Append tbl.CreateField("NombreArcano", dbText, 50)
    tbl.Fields.Append tbl.CreateField("LetraHebrea", dbText, 20)
    tbl.Fields.Append tbl.CreateField("ValorLetraHebrea", dbInteger)
    tbl.Fields.Append tbl.CreateField("CorrespondenciaAstrologica", dbText, 50)
    tbl.Fields.Append tbl.CreateField("TipoCorrespondencia", dbText, 20)
    tbl.Fields.Append tbl.CreateField("ElementoAsociado", dbText, 10)
    tbl.Fields.Append tbl.CreateField("Polaridad", dbText, 10)
    tbl.Fields.Append tbl.CreateField("Cualidad", dbText, 15)
    tbl.Fields.Append tbl.CreateField("SenderoArbolVida", dbText, 50)
    tbl.Fields.Append tbl.CreateField("SephirothInicio", dbText, 20)
    tbl.Fields.Append tbl.CreateField("SephirothFin", dbText, 20)
    tbl.Fields.Append tbl.CreateField("InterpretacionAstrologica", dbMemo)
    tbl.Fields.Append tbl.CreateField("PalabrasClave", dbMemo)
    
    ' Clave primaria
    Dim idx As DAO.Index
    Set idx = tbl.CreateIndex("PrimaryKey")
    idx.Primary = True
    idx.Required = True
    Set fld = idx.CreateField("ID_Arcano")
    idx.Fields.Append fld
    tbl.Indexes.Append idx
    
    db.TableDefs.Append tbl
    
    Debug.Print "? Tabla tblArcanosMayoresAstrologia creada"
    
    Set fld = Nothing
    Set idx = Nothing
    Set tbl = Nothing
    Set db = Nothing
    Exit Sub
    
ErrorHandler:
    Debug.Print "ERROR creando tblArcanosMayoresAstrologia: " & Err.Description
    Resume Next
End Sub

Private Sub CrearTablaArcanosMenoresNumeradosAstrologia()
    ' Crea tabla para cartas numeradas (2-10 de cada palo = 36 cartas)
    
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim tbl As DAO.TableDef
    Dim fld As DAO.Field
    
    Set db = CurrentDb
    Set tbl = db.CreateTableDef("tblArcanosMenoresNumeradosAstrologia")
    
    Set fld = tbl.CreateField("ID_Carta", dbLong)
    fld.Attributes = dbAutoIncrField
    tbl.Fields.Append fld
    
    tbl.Fields.Append tbl.CreateField("NumeroCarta", dbByte)
    tbl.Fields.Append tbl.CreateField("Palo", dbText, 20)
    tbl.Fields.Append tbl.CreateField("NombreCompleto", dbText, 50)
    tbl.Fields.Append tbl.CreateField("Decanato_ID", dbLong)
    tbl.Fields.Append tbl.CreateField("SignoZodiacal", dbText, 20)
    tbl.Fields.Append tbl.CreateField("GradosZodiacales", dbText, 20)
    tbl.Fields.Append tbl.CreateField("PlanetaEnSigno", dbText, 50)
    tbl.Fields.Append tbl.CreateField("ElementoPalo", dbText, 10)
    tbl.Fields.Append tbl.CreateField("ElementoSigno", dbText, 10)
    tbl.Fields.Append tbl.CreateField("Cualidad", dbText, 15)
    tbl.Fields.Append tbl.CreateField("TituloGoldenDawn", dbText, 100)
    tbl.Fields.Append tbl.CreateField("InterpretacionAstrologica", dbMemo)
    tbl.Fields.Append tbl.CreateField("PalabrasClave", dbMemo)
    
    ' Clave primaria
    Dim idx As DAO.Index
    Set idx = tbl.CreateIndex("PrimaryKey")
    idx.Primary = True
    idx.Required = True
    Set fld = idx.CreateField("ID_Carta")
    idx.Fields.Append fld
    tbl.Indexes.Append idx
    
    db.TableDefs.Append tbl
    
    Debug.Print "? Tabla tblArcanosMenoresNumeradosAstrologia creada"
    
    Set fld = Nothing
    Set idx = Nothing
    Set tbl = Nothing
    Set db = Nothing
    Exit Sub
    
ErrorHandler:
    Debug.Print "ERROR creando tblArcanosMenoresNumeradosAstrologia: " & Err.Description
    Resume Next
End Sub

Private Sub CrearTablaFigurasCorteAstrologia()
    ' Crea tabla para figuras de corte (16 cartas)
    
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim tbl As DAO.TableDef
    Dim fld As DAO.Field
    
    Set db = CurrentDb
    Set tbl = db.CreateTableDef("tblFigurasCorteAstrologia")
    
    Set fld = tbl.CreateField("ID_Figura", dbLong)
    fld.Attributes = dbAutoIncrField
    tbl.Fields.Append fld
    
    tbl.Fields.Append tbl.CreateField("TipoFigura", dbText, 20)
    tbl.Fields.Append tbl.CreateField("Palo", dbText, 20)
    tbl.Fields.Append tbl.CreateField("NombreCompleto", dbText, 50)
    tbl.Fields.Append tbl.CreateField("ElementoPalo", dbText, 10)
    tbl.Fields.Append tbl.CreateField("ElementoFigura", dbText, 10)
    tbl.Fields.Append tbl.CreateField("CombinacionElemental", dbText, 30)
    tbl.Fields.Append tbl.CreateField("CuadranteZodiacal", dbText, 100)
    tbl.Fields.Append tbl.CreateField("SignosPrincipales", dbText, 100)
    tbl.Fields.Append tbl.CreateField("GradosIniciales", dbText, 20)
    tbl.Fields.Append tbl.CreateField("GradosFinales", dbText, 20)
    tbl.Fields.Append tbl.CreateField("CaracteristicasAstrologicas", dbMemo)
    tbl.Fields.Append tbl.CreateField("PersonalidadTipo", dbMemo)
    tbl.Fields.Append tbl.CreateField("PalabrasClave", dbMemo)
    
    ' Clave primaria
    Dim idx As DAO.Index
    Set idx = tbl.CreateIndex("PrimaryKey")
    idx.Primary = True
    idx.Required = True
    Set fld = idx.CreateField("ID_Figura")
    idx.Fields.Append fld
    tbl.Indexes.Append idx
    
    db.TableDefs.Append tbl
    
    Debug.Print "? Tabla tblFigurasCorteAstrologia creada"
    
    Set fld = Nothing
    Set idx = Nothing
    Set tbl = Nothing
    Set db = Nothing
    Exit Sub
    
ErrorHandler:
    Debug.Print "ERROR creando tblFigurasCorteAstrologia: " & Err.Description
    Resume Next
End Sub

Private Sub CrearTablaNumerosAstrologia()
    ' Crea tabla de correspondencias números-planetas
    
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim tbl As DAO.TableDef
    Dim fld As DAO.Field
    
    Set db = CurrentDb
    Set tbl = db.CreateTableDef("tblNumerosAstrologia")
    
    Set fld = tbl.CreateField("ID_Numero", dbLong)
    fld.Attributes = dbAutoIncrField
    tbl.Fields.Append fld
    
    tbl.Fields.Append tbl.CreateField("Numero", dbByte)
    tbl.Fields.Append tbl.CreateField("EsNumeroMaestro", dbBoolean)
    tbl.Fields.Append tbl.CreateField("PlanetaPrimario", dbText, 20)
    tbl.Fields.Append tbl.CreateField("PlanetaSecundario", dbText, 20)
    tbl.Fields.Append tbl.CreateField("SignosAfines", dbText, 100)
    tbl.Fields.Append tbl.CreateField("ElementoAfin", dbText, 10)
    tbl.Fields.Append tbl.CreateField("ArcanosMayoresRelacionados", dbText, 100)
    tbl.Fields.Append tbl.CreateField("CaracteristicasNumerologicas", dbMemo)
    tbl.Fields.Append tbl.CreateField("CaracteristicasAstrologicas", dbMemo)
    tbl.Fields.Append tbl.CreateField("IntegracionInterpretativa", dbMemo)
    
    ' Clave primaria
    Dim idx As DAO.Index
    Set idx = tbl.CreateIndex("PrimaryKey")
    idx.Primary = True
    idx.Required = True
    Set fld = idx.CreateField("ID_Numero")
    idx.Fields.Append fld
    tbl.Indexes.Append idx
    
    db.TableDefs.Append tbl
    
    Debug.Print "? Tabla tblNumerosAstrologia creada"
    
    Set fld = Nothing
    Set idx = Nothing
    Set tbl = Nothing
    Set db = Nothing
    Exit Sub
    
ErrorHandler:
    Debug.Print "ERROR creando tblNumerosAstrologia: " & Err.Description
    Resume Next
End Sub

Private Sub CrearTablaElementos()
    ' Crea tabla de los 4 elementos
    
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim tbl As DAO.TableDef
    Dim fld As DAO.Field
    
    Set db = CurrentDb
    Set tbl = db.CreateTableDef("tblElementos")
    
    Set fld = tbl.CreateField("ID_Elemento", dbLong)
    fld.Attributes = dbAutoIncrField
    tbl.Fields.Append fld
    
    tbl.Fields.Append tbl.CreateField("NombreElemento", dbText, 10)
    tbl.Fields.Append tbl.CreateField("SimboloElemento", dbText, 5)
    tbl.Fields.Append tbl.CreateField("Polaridad", dbText, 10)
    tbl.Fields.Append tbl.CreateField("Cualidades", dbText, 50)
    tbl.Fields.Append tbl.CreateField("PaloTarotAsociado", dbText, 20)
    tbl.Fields.Append tbl.CreateField("SignosAsociados", dbText, 50)
    tbl.Fields.Append tbl.CreateField("TriplicidadSignos", dbText, 50)
    tbl.Fields.Append tbl.CreateField("CaracteristicasClave", dbMemo)
    tbl.Fields.Append tbl.CreateField("PalabrasClave", dbMemo)
    
    ' Clave primaria
    Dim idx As DAO.Index
    Set idx = tbl.CreateIndex("PrimaryKey")
    idx.Primary = True
    idx.Required = True
    Set fld = idx.CreateField("ID_Elemento")
    idx.Fields.Append fld
    tbl.Indexes.Append idx
    
    db.TableDefs.Append tbl
    
    Debug.Print "? Tabla tblElementos creada"
    
    Set fld = Nothing
    Set idx = Nothing
    Set tbl = Nothing
    Set db = Nothing
    Exit Sub
    
ErrorHandler:
    Debug.Print "ERROR creando tblElementos: " & Err.Description
    Resume Next
End Sub

Private Sub CrearTablaCualidadesZodiacales()
    ' Crea tabla de cualidades (Cardinal, Fijo, Mutable)
    
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim tbl As DAO.TableDef
    Dim fld As DAO.Field
    
    Set db = CurrentDb
    Set tbl = db.CreateTableDef("tblCualidadesZodiacales")
    
    Set fld = tbl.CreateField("ID_Cualidad", dbLong)
    fld.Attributes = dbAutoIncrField
    tbl.Fields.Append fld
    
    tbl.Fields.Append tbl.CreateField("NombreCualidad", dbText, 15)
    tbl.Fields.Append tbl.CreateField("CruzCosmica", dbText, 50)
    tbl.Fields.Append tbl.CreateField("SignosAsociados", dbText, 50)
    tbl.Fields.Append tbl.CreateField("CaracteristicasPrincipales", dbMemo)
    tbl.Fields.Append tbl.CreateField("ModoExpresion", dbMemo)
    tbl.Fields.Append tbl.CreateField("PalabrasClave", dbMemo)
    
    ' Clave primaria
    Dim idx As DAO.Index
    Set idx = tbl.CreateIndex("PrimaryKey")
    idx.Primary = True
    idx.Required = True
    Set fld = idx.CreateField("ID_Cualidad")
    idx.Fields.Append fld
    tbl.Indexes.Append idx
    
    db.TableDefs.Append tbl
    
    Debug.Print "? Tabla tblCualidadesZodiacales creada"
    
    Set fld = Nothing
    Set idx = Nothing
    Set tbl = Nothing
    Set db = Nothing
    Exit Sub
    
ErrorHandler:
    Debug.Print "ERROR creando tblCualidadesZodiacales: " & Err.Description
    Resume Next
End Sub
