Attribute VB_Name = "modCrearTablas"
Option Compare Database
Option Explicit

' ============================================================================
' Proyecto:     Sistema de Numerología Tradicional y Fonético
' Módulo: modCrearTablas
' Descripción: Crea la estructura de tablas para el sistema de numerología
' Autor: Sistema de Numerología
' Fecha: 2024
' ============================================================================

Public Sub CrearTodasLasTablas()
    On Error GoTo ErrorHandler
    
    Debug.Print "=== INICIANDO CREACIÓN DE TABLAS ==="
    Debug.Print ""
    
    ' Crear tablas principales
    Call CrearTablaPersonas
    Call CrearTablaCalculos
    Call CrearTablaInterpretaciones
    Call CrearTablaSinastrias
    Call CrearTablaConfiguracion
    Call CrearTablaTiposCalculo
    Call CrearTablaTiposSinastria
    
    Debug.Print ""
    Debug.Print "=== TABLAS CREADAS EXITOSAMENTE ==="
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "ERROR al crear tablas: " & err.Description
    MsgBox "Error al crear tablas: " & err.Description, vbCritical, "Error"
End Sub

' ============================================================================
' TABLA: tblPersonas
' ============================================================================

Private Sub CrearTablaPersonas()
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim idx As DAO.Index
    
    On Error Resume Next
    CurrentDb.TableDefs.Delete "tblPersonas"
    On Error GoTo ErrorHandler
    
    Set tdf = CurrentDb.CreateTableDef("tblPersonas")
    
    ' ID
    Set fld = tdf.CreateField("PersonaID", dbLong)
    fld.Attributes = dbAutoIncrField
    tdf.Fields.Append fld
    
    ' Datos personales
    Set fld = tdf.CreateField("NombreCompleto", dbText, 255)
    fld.Required = True
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("NombreTradicional", dbText, 255)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("NombreFonetico", dbText, 255)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("FechaNacimiento", dbDate)
    fld.Required = True
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("LugarNacimiento", dbText, 100)
    tdf.Fields.Append fld
    
    ' Números calculados - Básicos
    Set fld = tdf.CreateField("CaminoVida", dbInteger)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("Destino", dbInteger)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("Alma", dbInteger)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("Personalidad", dbInteger)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("Madurez", dbInteger)
    tdf.Fields.Append fld
    
    ' Números calculados - Ciclos
    Set fld = tdf.CreateField("Ciclo1", dbInteger)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("Ciclo1Inicio", dbInteger)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("Ciclo1Fin", dbInteger)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("Ciclo2", dbInteger)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("Ciclo2Inicio", dbInteger)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("Ciclo2Fin", dbInteger)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("Ciclo3", dbInteger)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("Ciclo3Inicio", dbInteger)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("Ciclo4", dbInteger)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("Ciclo4Inicio", dbInteger)
    tdf.Fields.Append fld
    
    ' Números calculados - Pináculos
    Set fld = tdf.CreateField("Pinaculo1", dbInteger)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("Pinaculo1Inicio", dbInteger)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("Pinaculo1Fin", dbInteger)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("Pinaculo2", dbInteger)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("Pinaculo2Inicio", dbInteger)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("Pinaculo2Fin", dbInteger)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("Pinaculo3", dbInteger)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("Pinaculo3Inicio", dbInteger)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("Pinaculo3Fin", dbInteger)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("Pinaculo4", dbInteger)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("Pinaculo4Inicio", dbInteger)
    tdf.Fields.Append fld
    
    ' Números calculados - Desafíos
    Set fld = tdf.CreateField("Desafio1", dbInteger)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("Desafio1Inicio", dbInteger)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("Desafio1Fin", dbInteger)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("Desafio2", dbInteger)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("Desafio2Inicio", dbInteger)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("Desafio2Fin", dbInteger)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("Desafio3", dbInteger)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("Desafio3Inicio", dbInteger)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("Desafio3Fin", dbInteger)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("Desafio4", dbInteger)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("Desafio4Inicio", dbInteger)
    tdf.Fields.Append fld
    
    ' Números especiales
    Set fld = tdf.CreateField("PrimeraLetra", dbText, 1)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("PrimeraVocal", dbText, 1)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("PrimeraConsonante", dbText, 1)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("NumeroPoder", dbInteger)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("NumeroDominante", dbInteger)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("NumerosFaltantes", dbText, 50)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("RespuestaSubconsciente", dbInteger)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("PlanoMental", dbInteger)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("PlanoFisico", dbInteger)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("PlanoEmocional", dbInteger)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("PlanoIntuitivo", dbInteger)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("PlanoDominante", dbText, 20)
    tdf.Fields.Append fld
    
    ' Metadatos
    Set fld = tdf.CreateField("FechaCreacion", dbDate)
    fld.DefaultValue = "Date()"
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("FechaModificacion", dbDate)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("Notas", dbMemo)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("Activo", dbBoolean)
    fld.DefaultValue = "True"
    tdf.Fields.Append fld
    
    ' Crear índice de clave primaria
    Set idx = tdf.CreateIndex("PrimaryKey")
    idx.Primary = True
    idx.Required = True
    Set fld = idx.CreateField("PersonaID")
    idx.Fields.Append fld
    tdf.Indexes.Append idx
    
    ' Crear índice para búsquedas por nombre
    Set idx = tdf.CreateIndex("idxNombre")
    Set fld = idx.CreateField("NombreCompleto")
    idx.Fields.Append fld
    tdf.Indexes.Append idx
    
    ' Agregar tabla a la base de datos
    CurrentDb.TableDefs.Append tdf
    
    Debug.Print "? Tabla tblPersonas creada"
    
    Set fld = Nothing
    Set idx = Nothing
    Set tdf = Nothing
    Exit Sub
    
ErrorHandler:
    Debug.Print "? Error al crear tblPersonas: " & err.Description
End Sub

' ============================================================================
' TABLA: tblCalculos
' ============================================================================

Private Sub CrearTablaCalculos()
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim idx As DAO.Index
    
    On Error Resume Next
    CurrentDb.TableDefs.Delete "tblCalculos"
    On Error GoTo ErrorHandler
    
    Set tdf = CurrentDb.CreateTableDef("tblCalculos")
    
    ' ID
    Set fld = tdf.CreateField("CalculoID", dbLong)
    fld.Attributes = dbAutoIncrField
    tdf.Fields.Append fld
    
    ' Relación con persona
    Set fld = tdf.CreateField("PersonaID", dbLong)
    fld.Required = True
    tdf.Fields.Append fld
    
    ' Tipo de cálculo
    Set fld = tdf.CreateField("TipoCalculoID", dbInteger)
    fld.Required = True
    tdf.Fields.Append fld
    
    ' Valores
    Set fld = tdf.CreateField("Valor", dbInteger)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("ValorTexto", dbText, 50)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("EsMaestro", dbBoolean)
    fld.DefaultValue = "False"
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("EsKarmico", dbBoolean)
    fld.DefaultValue = "False"
    tdf.Fields.Append fld
    
    ' Rango de edades (para ciclos, pináculos, desafíos)
    Set fld = tdf.CreateField("EdadInicio", dbInteger)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("EdadFin", dbInteger)
    tdf.Fields.Append fld
    
    ' Metadatos
    Set fld = tdf.CreateField("FechaCalculo", dbDate)
    fld.DefaultValue = "Date()"
    tdf.Fields.Append fld
    
    ' Crear índice de clave primaria
    Set idx = tdf.CreateIndex("PrimaryKey")
    idx.Primary = True
    idx.Required = True
    Set fld = idx.CreateField("CalculoID")
    idx.Fields.Append fld
    tdf.Indexes.Append idx
    
    ' Crear índice para búsquedas por persona
    Set idx = tdf.CreateIndex("idxPersona")
    Set fld = idx.CreateField("PersonaID")
    idx.Fields.Append fld
    tdf.Indexes.Append idx
    
    ' Agregar tabla a la base de datos
    CurrentDb.TableDefs.Append tdf
    
    Debug.Print "? Tabla tblCalculos creada"
    
    Set fld = Nothing
    Set idx = Nothing
    Set tdf = Nothing
    Exit Sub
    
ErrorHandler:
    Debug.Print "? Error al crear tblCalculos: " & err.Description
End Sub

' ============================================================================
' TABLA: tblInterpretaciones
' ============================================================================

Private Sub CrearTablaInterpretaciones()
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim idx As DAO.Index
    
    On Error Resume Next
    CurrentDb.TableDefs.Delete "tblInterpretaciones"
    On Error GoTo ErrorHandler
    
    Set tdf = CurrentDb.CreateTableDef("tblInterpretaciones")
    
    ' ID
    Set fld = tdf.CreateField("InterpretacionID", dbLong)
    fld.Attributes = dbAutoIncrField
    tdf.Fields.Append fld
    
    ' Tipo de interpretación
    Set fld = tdf.CreateField("TipoCalculoID", dbInteger)
    fld.Required = True
    tdf.Fields.Append fld
    
    ' Número
    Set fld = tdf.CreateField("Numero", dbInteger)
    fld.Required = True
    tdf.Fields.Append fld
    
    ' Ruta del archivo
    Set fld = tdf.CreateField("RutaArchivo", dbText, 255)
    fld.Required = True
    tdf.Fields.Append fld
    
    ' Contenido (opcional, para cache)
    Set fld = tdf.CreateField("Contenido", dbMemo)
    tdf.Fields.Append fld
    
    ' Metadatos
    Set fld = tdf.CreateField("FechaCreacion", dbDate)
    fld.DefaultValue = "Date()"
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("FechaModificacion", dbDate)
    tdf.Fields.Append fld
    
    ' Crear índice de clave primaria
    Set idx = tdf.CreateIndex("PrimaryKey")
    idx.Primary = True
    idx.Required = True
    Set fld = idx.CreateField("InterpretacionID")
    idx.Fields.Append fld
    tdf.Indexes.Append idx
    
    ' Crear índice único por tipo y número
    Set idx = tdf.CreateIndex("idxTipoNumero")
    idx.Unique = True
    Set fld = idx.CreateField("TipoCalculoID")
    idx.Fields.Append fld
    Set fld = idx.CreateField("Numero")
    idx.Fields.Append fld
    tdf.Indexes.Append idx
    
    ' Agregar tabla a la base de datos
    CurrentDb.TableDefs.Append tdf
    
    Debug.Print "? Tabla tblInterpretaciones creada"
    
    Set fld = Nothing
    Set idx = Nothing
    Set tdf = Nothing
    Exit Sub
    
ErrorHandler:
    Debug.Print "? Error al crear tblInterpretaciones: " & err.Description
End Sub

' ============================================================================
' TABLA: tblSinastrias
' ============================================================================

Private Sub CrearTablaSinastrias()
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim idx As DAO.Index
    
    On Error Resume Next
    CurrentDb.TableDefs.Delete "tblSinastrias"
    On Error GoTo ErrorHandler
    
    Set tdf = CurrentDb.CreateTableDef("tblSinastrias")
    
    ' ID
    Set fld = tdf.CreateField("SinastriaID", dbLong)
    fld.Attributes = dbAutoIncrField
    tdf.Fields.Append fld
    
    ' Personas involucradas
    Set fld = tdf.CreateField("Persona1ID", dbLong)
    fld.Required = True
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("Persona2ID", dbLong)
    fld.Required = True
    tdf.Fields.Append fld
    
    ' Tipo de sinastría
    Set fld = tdf.CreateField("TipoSinastriaID", dbInteger)
    fld.Required = True
    tdf.Fields.Append fld
    
    ' Números comparados
    Set fld = tdf.CreateField("CaminoVida1", dbInteger)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("CaminoVida2", dbInteger)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("Destino1", dbInteger)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("Destino2", dbInteger)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("Alma1", dbInteger)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("Alma2", dbInteger)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("Personalidad1", dbInteger)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("Personalidad2", dbInteger)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("Madurez1", dbInteger)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("Madurez2", dbInteger)
    tdf.Fields.Append fld
    
    ' Metadatos
    Set fld = tdf.CreateField("FechaCalculo", dbDate)
    fld.DefaultValue = "Date()"
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("Notas", dbMemo)
    tdf.Fields.Append fld
    
    ' Crear índice de clave primaria
    Set idx = tdf.CreateIndex("PrimaryKey")
    idx.Primary = True
    idx.Required = True
    Set fld = idx.CreateField("SinastriaID")
    idx.Fields.Append fld
    tdf.Indexes.Append idx
    
    ' Crear índice para búsquedas por personas
    Set idx = tdf.CreateIndex("idxPersonas")
    Set fld = idx.CreateField("Persona1ID")
    idx.Fields.Append fld
    Set fld = idx.CreateField("Persona2ID")
    idx.Fields.Append fld
    tdf.Indexes.Append idx
    
    ' Agregar tabla a la base de datos
    CurrentDb.TableDefs.Append tdf
    
    Debug.Print "? Tabla tblSinastrias creada"
    
    Set fld = Nothing
    Set idx = Nothing
    Set tdf = Nothing
    Exit Sub
    
ErrorHandler:
    Debug.Print "? Error al crear tblSinastrias: " & err.Description
End Sub

' ============================================================================
' TABLA: tblConfiguracion
' ============================================================================

Private Sub CrearTablaConfiguracion()
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim idx As DAO.Index
    
    On Error Resume Next
    CurrentDb.TableDefs.Delete "tblConfiguracion"
    On Error GoTo ErrorHandler
    
    Set tdf = CurrentDb.CreateTableDef("tblConfiguracion")
    
    ' ID
    Set fld = tdf.CreateField("ConfigID", dbLong)
    fld.Attributes = dbAutoIncrField
    tdf.Fields.Append fld
    
    ' Clave-Valor
    Set fld = tdf.CreateField("Clave", dbText, 50)
    fld.Required = True
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("Valor", dbText, 255)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("Descripcion", dbText, 255)
    tdf.Fields.Append fld
    
    ' Crear índice de clave primaria
    Set idx = tdf.CreateIndex("PrimaryKey")
    idx.Primary = True
    idx.Required = True
    Set fld = idx.CreateField("ConfigID")
    idx.Fields.Append fld
    tdf.Indexes.Append idx
    
    ' Crear índice único por clave
    Set idx = tdf.CreateIndex("idxClave")
    idx.Unique = True
    Set fld = idx.CreateField("Clave")
    idx.Fields.Append fld
    tdf.Indexes.Append idx
    
    ' Agregar tabla a la base de datos
    CurrentDb.TableDefs.Append tdf
    
    Debug.Print "? Tabla tblConfiguracion creada"
    
    Set fld = Nothing
    Set idx = Nothing
    Set tdf = Nothing
    Exit Sub
    
ErrorHandler:
    Debug.Print "? Error al crear tblConfiguracion: " & err.Description
End Sub

' ============================================================================
' TABLA: tblTiposCalculo
' ============================================================================

Private Sub CrearTablaTiposCalculo()
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim idx As DAO.Index
    
    On Error Resume Next
    CurrentDb.TableDefs.Delete "tblTiposCalculo"
    On Error GoTo ErrorHandler
    
    Set tdf = CurrentDb.CreateTableDef("tblTiposCalculo")
    
    ' ID
    Set fld = tdf.CreateField("TipoCalculoID", dbInteger)
    tdf.Fields.Append fld
    
    ' Nombre
    Set fld = tdf.CreateField("Nombre", dbText, 50)
    fld.Required = True
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("Descripcion", dbText, 255)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("CarpetaInterpretaciones", dbText, 100)
    tdf.Fields.Append fld
    
    ' Crear índice de clave primaria
    Set idx = tdf.CreateIndex("PrimaryKey")
    idx.Primary = True
    idx.Required = True
    Set fld = idx.CreateField("TipoCalculoID")
    idx.Fields.Append fld
    tdf.Indexes.Append idx
    
    ' Agregar tabla a la base de datos
    CurrentDb.TableDefs.Append tdf
    
    Debug.Print "? Tabla tblTiposCalculo creada"
    
    Set fld = Nothing
    Set idx = Nothing
    Set tdf = Nothing
    Exit Sub
    
ErrorHandler:
    Debug.Print "? Error al crear tblTiposCalculo: " & err.Description
End Sub

' ============================================================================
' TABLA: tblTiposSinastria
' ============================================================================

Private Sub CrearTablaTiposSinastria()
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim idx As DAO.Index
    
    On Error Resume Next
    CurrentDb.TableDefs.Delete "tblTiposSinastria"
    On Error GoTo ErrorHandler
    
    Set tdf = CurrentDb.CreateTableDef("tblTiposSinastria")
    
    ' ID
    Set fld = tdf.CreateField("TipoSinastriaID", dbInteger)
    tdf.Fields.Append fld
    
    ' Nombre
    Set fld = tdf.CreateField("Nombre", dbText, 50)
    fld.Required = True
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("Descripcion", dbText, 255)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("CarpetaInterpretaciones", dbText, 100)
    tdf.Fields.Append fld
    
    ' Crear índice de clave primaria
    Set idx = tdf.CreateIndex("PrimaryKey")
    idx.Primary = True
    idx.Required = True
    Set fld = idx.CreateField("TipoSinastriaID")
    idx.Fields.Append fld
    tdf.Indexes.Append idx
    
    ' Agregar tabla a la base de datos
    CurrentDb.TableDefs.Append tdf
    
    Debug.Print "? Tabla tblTiposSinastria creada"
    
    Set fld = Nothing
    Set idx = Nothing
    Set tdf = Nothing
    Exit Sub
    
ErrorHandler:
    Debug.Print "? Error al crear tblTiposSinastria: " & err.Description
End Sub
