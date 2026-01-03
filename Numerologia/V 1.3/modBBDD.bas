Attribute VB_Name = "modBBDD"

Option Compare Database
Option Explicit

Dim db As DAO.Database
    
Public Sub GuardarDatosPersona()
    On Error GoTo ErrHandler
    
    Dim rs As DAO.Recordset
    Dim nuevoID As Long
    Dim nombreCompleto As String
    Dim nombreFonetico As String

    '===============================
    ' VALIDACIÓN BÁSICA
    '===============================
    If Trim$(Me.txtNombreBase & "") = "" Then
        MsgBox "El nombre es obligatorio.", vbInformation, "Guardar datos"
        Me.txtNombreBase.SetFocus
        Exit Sub

    ElseIf Trim$(Me.txtApeBase1 & "") = "" Then
        MsgBox "El primer apellido es obligatorio.", vbInformation, "Guardar datos"
        Me.txtApeBase1.SetFocus
        Exit Sub
    End If

    '===============================
    ' GENERAR CAMPOS COMPUESTOS
    '===============================
    nombreCompleto = Trim$(Me.txtNombreBase & " " & Me.txtApeBase1 & " " & Nz(Me.txtApeBase2, ""))
    nombreFonetico = Trim$(Nz(Me.FonNombre, "") & " " & Nz(Me.FonApe1, "") & " " & Nz(Me.FonApe2, ""))

    '===============================
    ' GUARDAR EN TABLA
    '===============================
    Set db = CurrentDb
    Set rs = db.OpenRecordset("tbuPersonas", dbOpenDynaset)

    rs.AddNew

        '--- Datos base ---
        rs!Nombre = Nz(Me.txtNombreBase, "")
        rs!Ape1 = Nz(Me.txtApeBase1, "")
        rs!Ape2 = Nz(Me.txtApeBase2, "")

        '--- Idiomas por campo (FK numérica) ---
        If Not Me.IdiomaNom Is Nothing Then rs!IdiomaNombre = Me.IdiomaNom.ID
        If Not Me.IdiomaApe1 Is Nothing Then rs!IdiomaApe1 = Me.IdiomaApe1.ID
        If Not Me.IdiomaApe2 Is Nothing Then rs!IdiomaApe2 = Me.IdiomaApe2.ID

        '--- Fecha nacimiento ---
        If Not IsNull(Me.txtFechaNac) Then
            rs!FechaNacimiento = Me.txtFechaNac
        End If

        '--- Sexo legal ---
        If Not IsNull(Me.cboSexoLegal) Then
            rs!Genero = Me.cboGenero
        End If

        '--- Fonéticos por campo ---
        rs!FonNombre = Nz(Me.FonNombre, "")
        rs!FonApe1 = Nz(Me.FonApe1, "")
        rs!FonApe2 = Nz(Me.FonApe2, "")

'----------------------------------------------------------------
'          No hace falta, se puede reconstruir posteriormente
        '--- Campos compuestos ---
'        rs!nombreCompleto = nombreCompleto
'        rs!nombreFonetico = nombreFonetico
'----------------------------------------------------------------

        '--- Sistema de cálculo ---
        rs!SistemaCalculo = Me.fraSistema

        '--- Opciones fonéticas ---
        rs!UsarHmuda = Me.chkHmuda
        rs!UsarUmuda = Me.chkUmuda

        '--- Auditoría ---
        rs!FechaAlta = Now()
        rs!FechaModificacion = Now()

    rs.Update

    nuevoID = rs!IDPersona

    rs.Close
    Set rs = Nothing
    Set db = Nothing

    MsgBox "Datos guardados correctamente. ID asignado: " & nuevoID, _
           vbInformation, "Guardar datos"
    Exit Sub

'===============================
' GESTIÓN DE ERRORES
'===============================
ErrHandler:
    MsgBox "Error al guardar los datos: " & Err.Description, _
           vbExclamation, "Error"

End Sub

Public Sub GuardarPersonaEnTabla()

    Dim rs As DAO.Recordset
    Dim sql As String
    Dim nueva As Boolean

    ' ------------------------------------------------------------
    ' 1. Determinar si es nueva persona o actualización
    ' ------------------------------------------------------------
    If Persona.ID_Persona = 0 Then
        nueva = True
        sql = "SELECT * FROM tbuPersonas WHERE 1=0;"   ' recordset vacío para INSERT
    Else
        nueva = False
        sql = "SELECT * FROM tbuPersonas WHERE ID_Persona = " & Persona.ID_Persona & ";"
    End If

    Set rs = CurrentDb.OpenRecordset(sql, dbOpenDynaset)

    ' ------------------------------------------------------------
    ' 2. Si es nueva, añadir registro
    ' ------------------------------------------------------------
    If nueva Then
        rs.AddNew
        Persona.FechaAlta = Now
    Else
        rs.Edit
    End If

    ' ------------------------------------------------------------
    ' 3. Guardar datos personales
    ' ------------------------------------------------------------
    rs!Nombre = Persona.Nombre
    rs!Ape1 = Persona.Ape1
    rs!Ape2 = Persona.Ape2

    If Persona.FechaNacimiento = 0 Then
        rs!FechaNacimiento = Null
    Else
        rs!FechaNacimiento = Persona.FechaNacimiento
    End If

    rs!ID_Genero = Persona.ID_Genero

    ' ------------------------------------------------------------
    ' 4. Guardar idiomas (OBJETO ? ID)
    ' ------------------------------------------------------------
    If Not Persona.IdiomaNombre Is Nothing Then
        rs!IdiomaNombre = Persona.IdiomaNombre.IDIdioma
    Else
        rs!IdiomaNombre = Null
    End If

    If Not Persona.IdiomaApe1 Is Nothing Then
        rs!IdiomaApe1 = Persona.IdiomaApe1.IDIdioma
    Else
        rs!IdiomaApe1 = Null
    End If

    If Not Persona.IdiomaApe2 Is Nothing Then
        rs!IdiomaApe2 = Persona.IdiomaApe2.IDIdioma
    Else
        rs!IdiomaApe2 = Null
    End If

    ' ------------------------------------------------------------
    ' 5. Guardar fonética
    ' ------------------------------------------------------------
    rs!FonNombre = Persona.FonNombre
    rs!FonApe1 = Persona.FonApe1
    rs!FonApe2 = Persona.FonApe2

    ' ------------------------------------------------------------
    ' 6. Guardar parámetros de análisis
    ' ------------------------------------------------------------
    rs!Sistema = Persona.Sistema
    rs!Tarot = Persona.Tarot
    rs!UsarHmuda = Persona.UsarHmuda
    rs!UsarUmuda = Persona.UsarUmuda

    ' ------------------------------------------------------------
    ' 7. Gestión
    ' ------------------------------------------------------------
    rs!FechaAlta = Persona.FechaAlta
    Persona.FechaModificacion = Now
    rs!FechaModificacion = Persona.FechaModificacion

    ' ------------------------------------------------------------
    ' 8. Confirmar cambios
    ' ------------------------------------------------------------
    rs.Update

    ' ------------------------------------------------------------
    ' 9. Si era nuevo, recuperar el ID asignado
    ' ------------------------------------------------------------
    If nueva Then
        Persona.ID_Persona = rs!ID_Persona
    End If

    rs.Close
    Set rs = Nothing

End Sub


Public Sub CargarClaseDesdeTabla(ByVal ID As Long)

    Dim rs As DAO.Recordset
    Dim sql As String

    sql = "SELECT * FROM tbuPersonas WHERE ID_Persona = " & ID & ";"
    Set rs = CurrentDb.OpenRecordset(sql, dbOpenDynaset)

    If Not rs.EOF Then

        ' --- Identificación ---
        Persona.ID_Persona = Nz(rs!ID_Persona, 0)

        ' --- Datos personales ---
        Persona.Nombre = Nz(rs!Nombre, "")
        Persona.Ape1 = Nz(rs!Ape1, "")
        Persona.Ape2 = Nz(rs!Ape2, "")
        Persona.FechaNacimiento = Nz(rs!FechaNacimiento, 0)
        Persona.ID_Genero = Nz(rs!ID_Genero, 0)

        ' --- Idiomas ---
'        Persona.IdiomaNombre = Nz(rs!IdiomaNombre, 0)
'        Persona.IdiomaApe1 = Nz(rs!IdiomaApe1, 0)
'        Persona.IdiomaApe2 = Nz(rs!IdiomaApe2, 0)

        Set Persona.IdiomaNombre = CargarIdiomaDesdeID(Nz(rs!IdiomaNombre, 0))
        Set Persona.IdiomaApe1 = CargarIdiomaDesdeID(Nz(rs!IdiomaApe1, 0))
        Set Persona.IdiomaApe2 = CargarIdiomaDesdeID(Nz(rs!IdiomaApe2, 0))

        ' --- Fonética ---
        Persona.FonNombre = Nz(rs!FonNombre, "")
        Persona.FonApe1 = Nz(rs!FonApe1, "")
        Persona.FonApe2 = Nz(rs!FonApe2, "")

        ' --- Parámetros de análisis ---
        Persona.Sistema = Nz(rs!Sistema, 0)
        Persona.Tarot = Nz(rs!Tarot, 0)
        Persona.UsarHmuda = Nz(rs!UsarHmuda, False)
        Persona.UsarUmuda = Nz(rs!UsarUmuda, False)

        ' --- Gestión ---
        Persona.FechaAlta = Nz(rs!FechaAlta, 0)
        Persona.FechaModificacion = Nz(rs!FechaModificacion, 0)

    End If

    rs.Close
    Set rs = Nothing

End Sub


Public Function CargarIdiomaDesdeID(ByVal ID As Byte) As clsIdioma

    If ID = 0 Then Exit Function

    Dim rs As DAO.Recordset
    Dim sql As String
    Dim i As New clsIdioma

    sql = "SELECT * FROM tbmIdiomas WHERE IDIdioma = " & ID
    Set rs = CurrentDb.OpenRecordset(sql)

    If Not rs.EOF Then
        i.Init rs!IDIdioma, rs!Abreviado, rs!NomIdioma, Nz(rs!notas, "")
        Set CargarIdiomaDesdeID = i
    End If

    rs.Close
    Set rs = Nothing

End Function

'Public Sub RecuperarPersona(IDPersona As Long)
'
'    Dim p As clsPersona
'    Set p = CargarPersonaDesdeBD(IDPersona)
'
'    If p Is Nothing Then
'        MsgBox "No se encontró la persona.", vbExclamation
'        Exit Sub
'    End If
'
'    PoblarFormulario p
'    MostrarEstadoResultados IDPersona
'
'End Sub


Public Sub MostrarEstadoResultados(IDPersona As Long)

    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT IDPersona FROM tbuResultados WHERE IDPersona=" & IDPersona)

    With Forms!frmTomaDatos
        If rs.EOF Then
            .lblResultados.Caption = "Sin resultados calculados"
            .lblResultados.ForeColor = vbRed
        Else
            .lblResultados.Caption = "Resultados disponibles"
            .lblResultados.ForeColor = vbGreen
        End If
    End With

    rs.Close

End Sub



Public Sub Actualizar_tbuResultados()

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    
    Set db = CurrentDb
    Set tdf = db.TableDefs("tbuResultados")
    
    '--- Función local para añadir campo si no existe ---
    Dim AddField As Object
    Set AddField = CreateObject("Scripting.Dictionary")
    
    AddField("IDAnalisis") = dbLong
    AddField("PlanoFisico") = dbText
    AddField("PlanoEmocional") = dbText
    AddField("PlanoMental") = dbText
    AddField("PlanoIntuitivo") = dbText
    AddField("PiedraAngular") = dbText
    AddField("PiedraToque") = dbText
    AddField("PrimeraLetra") = dbText
    AddField("PrimeraVocal") = dbText
    AddField("PrimeraConsonante") = dbText
    AddField("RespuestaSubconsciente") = dbText
    AddField("Poder") = dbText
    AddField("DeudaKarmica") = dbText
    
    Dim key As Variant
    For Each key In AddField.Keys
        On Error Resume Next
        Set fld = tdf.Fields(key)
        On Error GoTo 0
        
        If fld Is Nothing Then
            tdf.Fields.Append tdf.CreateField(key, AddField(key), 50)
        End If
        
        Set fld = Nothing
    Next key
    
    MsgBox "tbuResultados actualizado correctamente.", vbInformation

End Sub

Public Sub CrearTabla_tbuInclusion()

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
'    Dim fld As DAO.Field
    
    Set db = CurrentDb

    ' Si la tabla existe, salir sin hacer nada
    On Error Resume Next
    Set tdf = db.TableDefs("tbuInclusion")
    On Error GoTo 0
    
    If Not tdf Is Nothing Then
        MsgBox "La tabla tbuInclusion ya existe.", vbInformation
        Exit Sub
    End If

    ' Crear la tabla
    Set tdf = db.CreateTableDef("tbuInclusion")

    ' Campos principales
    tdf.Fields.Append tdf.CreateField("IDFonetica", dbLong)
    tdf.Fields("IDFonetica").Attributes = dbAutoIncrField

    tdf.Fields.Append tdf.CreateField("IDResultado", dbLong)
    tdf.Fields.Append tdf.CreateField("IDPersona", dbLong)

    ' Campos N1 a N9 tipo Byte
    Dim i As Integer
    For i = 1 To 9
        tdf.Fields.Append tdf.CreateField("N" & i, dbByte)
    Next i

    ' Añadir tabla a la base de datos
    db.TableDefs.Append tdf

    MsgBox "Tabla tbuFoneticaResumen creada correctamente.", vbInformation

End Sub

Public Sub CrearRelacion_Inclusion_Resultados()

    Dim db As DAO.Database
    Dim rel As DAO.Relation
    Dim fld As DAO.Field
    
    Set db = CurrentDb

    ' Eliminar relación previa si existe
    On Error Resume Next
    db.Relations.Delete "rel_Inclusion_Resultados"
    On Error GoTo 0

    ' Crear relación
    Set rel = db.CreateRelation("rel_Inclusion_Resultados", _
                                "tbuResultados", "tbuInclusion", _
                                dbRelationUpdateCascade + dbRelationDeleteCascade)

    Set fld = rel.CreateField("IDResultado")
    fld.ForeignName = "IDResultado"
    rel.Fields.Append fld

    db.Relations.Append rel

    MsgBox "Relación creada correctamente.", vbInformation

End Sub


