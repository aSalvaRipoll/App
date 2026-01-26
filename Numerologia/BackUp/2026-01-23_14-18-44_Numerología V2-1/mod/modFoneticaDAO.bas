Attribute VB_Name = "modFoneticaDAO"
' ------------------------------------------------------
' Nombre:    modFoneticaDAO
' Tipo:      Módulo
' Propósito:
' Autor:     asalv
' Fecha:     15/01/2026
' ------------------------------------------------------

' ============================================================
'   modFoneticaDAO — Acceso a datos para clsFonetica
' ============================================================

Option Compare Database
Option Explicit


' ============================================================
'   CARGAR FONÉTICA POR ID
' ============================================================
Public Function CargarFonetica(ByVal IDFonetica As Long) As clsFonetica
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim f As clsFonetica
    Dim sql As String

    sql = "SELECT * FROM Fonetica WHERE IDFonetica = " & IDFonetica

    Set db = CurrentDb
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)

    If Not rs.EOF Then
        Set f = New clsFonetica

        f.IDFonetica = rs!IDFonetica
        f.IDPersona = rs!IDPersona

        f.ModoFonetico = rs!ModoFonetico
'        f.UsarHmuda = rs!UsarHmuda
'        f.UsarUmuda = rs!UsarUmuda

        ' Idiomas (no se cargan como objetos)
'        Set f.IdiomaNombre = rs!IDIdiomaNombre 'CargarIdiomaDesdeID(rs!IDIdiomaNombre)
'        Set f.IdiomaApe1 = rs!IDIdiomaApe1 'CargarIdiomaDesdeID(rs!IDIdiomaApe1)
'        Set f.IdiomaApe2 = rs!IDIdiomaApe2 'CargarIdiomaDesdeID(rs!IDIdiomaApe2)
        f.IdiomaNombre = rs!IDIdiomaNombre 'CargarIdiomaDesdeID(rs!IDIdiomaNombre)
        f.IdiomaApe1 = rs!IDIdiomaApe1 'CargarIdiomaDesdeID(rs!IDIdiomaApe1)
        f.IdiomaApe2 = rs!IDIdiomaApe2 'CargarIdiomaDesdeID(rs!IDIdiomaApe2)

        ' Resultados fonéticos
        f.FonNombre = Nz(rs!FonNombre, "")
        f.FonApe1 = Nz(rs!FonApe1, "")
        f.FonApe2 = Nz(rs!FonApe2, "")

        ' Datos originales
'        f.NombreOriginal = Nz(rs!NombreOriginal, "")
'        f.Ape1Original = Nz(rs!Ape1Original, "")
'        f.Ape2Original = Nz(rs!Ape2Original, "")

        f.FechaCalculo = rs!FechaCalculo
    End If

    rs.Close
    Set rs = Nothing
    Set db = Nothing

    Set CargarFonetica = f
End Function


' ============================================================
'   GUARDAR FONÉTICA (INSERTAR O ACTUALIZAR)
' ============================================================
Public Function GuardarFonetica(ByVal f As clsFonetica) As Long
    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    Set db = CurrentDb

    If f.IDFonetica = 0 Then
        ' --- INSERTAR ---
        Set rs = db.OpenRecordset("Fonetica", dbOpenDynaset)
        rs.AddNew
    Else
        ' --- ACTUALIZAR ---
        Set rs = db.OpenRecordset("SELECT * FROM Fonetica WHERE IDFonetica=" & f.IDFonetica, dbOpenDynaset)
        If rs.EOF Then
            rs.Close
            Set rs = db.OpenRecordset("Fonetica", dbOpenDynaset)
            rs.AddNew
        Else
            rs.Edit
        End If
    End If

    ' --- Campos ---
    rs!IDPersona = f.IDPersona
    rs!ModoFonetico = f.ModoFonetico
'    rs!UsarHmuda = f.UsarHmuda
'    rs!UsarUmuda = f.UsarUmuda

    rs!IDIdiomaNombre = f.IdiomaNombre '.IDIdioma
    rs!IDIdiomaApe1 = f.IdiomaApe1 '.IDIdioma
    rs!IDIdiomaApe2 = f.IdiomaApe2 '.IDIdioma

    rs!FonNombre = f.FonNombre
    rs!FonApe1 = f.FonApe1
    rs!FonApe2 = f.FonApe2

'    rs!NombreOriginal = f.NombreOriginal
'    rs!Ape1Original = f.Ape1Original
'    rs!Ape2Original = f.Ape2Original

    rs!FechaCalculo = f.FechaCalculo

    rs.Update

    ' Devolver ID
    If f.IDFonetica = 0 Then
        rs.Bookmark = rs.LastModified
        GuardarFonetica = rs!IDFonetica
    Else
        GuardarFonetica = f.IDFonetica
    End If

    rs.Close
    Set rs = Nothing
    Set db = Nothing
End Function


' ============================================================
'   ELIMINAR FONÉTICA
' ============================================================
Public Sub EliminarFonetica(ByVal IDFonetica As Long)
    Dim db As DAO.Database
    Set db = CurrentDb

    db.Execute "DELETE FROM Fonetica WHERE IDFonetica = " & IDFonetica, dbFailOnError

    Set db = Nothing
End Sub

'-----------------------------------------------------------------------------------------------------

Public Function GuardarFoneticaSmart(f As clsFonetica) As Long
    Dim FDB As clsFonetica

    ' ¿Existe ya una configuración para esta persona?
    Set FDB = CargarFoneticaPorPersona(f.IDPersona)

    If Not FDB Is Nothing Then
        ' Si la configuración es igual ? actualizar
        If ConfiguracionIgual(f, FDB) Then
            f.IDFonetica = FDB.IDFonetica
            GuardarFoneticaSmart = GuardarFonetica(f)
            Exit Function
        End If
    End If

    ' Si no existe o ha cambiado ? insertar nuevo
    f.IDFonetica = 0
    GuardarFoneticaSmart = GuardarFonetica(f)
End Function

Public Function CargarFoneticaPorPersona(IDPersona As Long) As clsFonetica
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String

    sql = "SELECT TOP 1 * FROM Fonetica WHERE IDPersona = " & IDPersona & " ORDER BY FechaCalculo DESC"

    Set db = CurrentDb
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)

    If Not rs.EOF Then
        Set CargarFoneticaPorPersona = CargarFonetica(rs!IDFonetica)
    End If

    rs.Close
    Set rs = Nothing
    Set db = Nothing
End Function


Public Function ConfiguracionIgual(f As clsFonetica, FDB As clsFonetica) As Boolean
    If f.ModoFonetico <> FDB.ModoFonetico Then GoTo Diferente
'    If f.UsarHmuda <> FDB.UsarHmuda Then GoTo Diferente
'    If f.UsarUmuda <> FDB.UsarUmuda Then GoTo Diferente

'    If f.IdiomaNombre.IDIdioma <> FDB.IdiomaNombre.IDIdioma Then GoTo Diferente
'    If f.IdiomaApe1.IDIdioma <> FDB.IdiomaApe1.IDIdioma Then GoTo Diferente
'    If f.IdiomaApe2.IDIdioma <> FDB.IdiomaApe2.IDIdioma Then GoTo Diferente

    If f.IdiomaNombre <> FDB.IdiomaNombre Then GoTo Diferente
    If f.IdiomaApe1 <> FDB.IdiomaApe1 Then GoTo Diferente
    If f.IdiomaApe2 <> FDB.IdiomaApe2 Then GoTo Diferente

    ConfiguracionIgual = True
    Exit Function

Diferente:
    ConfiguracionIgual = False
End Function

'-----------------------------------------------------------------------------------------------------

'Public Sub CargarClaseFoneticaDesdeFormulario(frm As Form)
'    ' Asegurar instancia
'    If Fonetica Is Nothing Then Set Fonetica = New clsFonetica
'
'    ' --- Identificación ---
'    Fonetica.IDPersona = frm!IDPersona
'
'    ' --- Modo ---
'    Fonetica.ModoFonetico = IIf(frm!chkModoFon, 1, 0)
'
'    ' --- Reglas ---
''    Fonetica.UsarHmuda = frm!chkHmuda
''    Fonetica.UsarUmuda = frm!chkUmuda
'
'    ' --- Idiomas ---
'    Set Fonetica.IdiomaNombre = CargarIdioma(frm!cboIdiomaNombre)
'    Set Fonetica.IdiomaApe1 = CargarIdioma(frm!cboIdiomaApe1)
'    Set Fonetica.IdiomaApe2 = CargarIdioma(frm!cboIdiomaApe2)
'
'    ' --- Datos originales ---
''    Fonetica.NombreOriginal = frm!txtNombre
''    Fonetica.Ape1Original = frm!txtApe1
''    Fonetica.Ape2Original = frm!txtApe2
'
'    ' --- Resultados fonéticos ---
'    Fonetica.FonNombre = MotorFonetico_Convertir(frm!txtNombre, Fonetica.IdiomaNombre, Fonetica.UsarHmuda, Fonetica.UsarUmuda, Fonetica.ModoFonetico)
'    Fonetica.FonApe1 = MotorFonetico_Convertir(frm!txtApe1, Fonetica.IdiomaApe1, Fonetica.UsarHmuda, Fonetica.UsarUmuda, Fonetica.ModoFonetico)
'    Fonetica.FonApe2 = MotorFonetico_Convertir(frm!txtApe2, Fonetica.IdiomaApe2, Fonetica.UsarHmuda, Fonetica.UsarUmuda, Fonetica.ModoFonetico)
'
'    ' --- Gestión ---
'    Fonetica.FechaCalculo = Now
'End Sub

'Public Function CrearObjetoFoneticaDesdeFormulario(frm As Form) As clsFonetica
'    Dim F As New clsFonetica
'
'    ' --- Identificación ---
'    F.IDPersona = frm!IDPersona
'
'    ' --- Modo ---
'    F.ModoFonetico = IIf(frm!chkModoFon, 1, 0)
'
'    ' --- Reglas ---
'    F.UsarHmuda = frm!chkHmuda
'    F.UsarUmuda = frm!chkUmuda
'
'    ' --- Idiomas ---
'    Set F.IdiomaNombre = CargarIdioma(frm!cboIdiomaNombre)
'    Set F.IdiomaApe1 = CargarIdioma(frm!cboIdiomaApe1)
'    Set F.IdiomaApe2 = CargarIdioma(frm!cboIdiomaApe2)
'
'    ' --- Datos originales ---
'    F.NombreOriginal = frm!txtNombre
'    F.Ape1Original = frm!txtApe1
'    F.Ape2Original = frm!txtApe2
'
'    ' --- Resultados fonéticos (desde el motor) ---
'    F.FonNombre = ObtenerFonNombre(frm)
'    F.FonApe1 = ObtenerFonApe1(frm)
'    F.FonApe2 = ObtenerFonApe2(frm)
'
'    ' --- Gestión ---
'    F.FechaCalculo = Now
'
'    Set CrearObjetoFoneticaDesdeFormulario = F
'End Function

'Public Sub CargarFormularioDesdeFonetica(frm As Form, f As clsFonetica)
'
'    ' --- Modo ---
'    frm!chkModoFon = (f.ModoFonetico = 1)
'
'    ' --- Reglas ---
''    frm!chkHmuda = f.UsarHmuda
''    frm!chkUmuda = f.UsarUmuda
'
'    ' --- Idiomas ---
'    frm!cboIdiomaNombre = f.IdiomaNombre.IDIdioma
'    frm!cboIdiomaApe1 = f.IdiomaApe1.IDIdioma
'    frm!cboIdiomaApe2 = f.IdiomaApe2.IDIdioma
'
'    ' --- Datos originales ---
''    frm!txtNombre = f.NombreOriginal
''    frm!txtApe1 = f.Ape1Original
''    frm!txtApe2 = f.Ape2Original
'
'    ' --- Resultados fonéticos ---
'    frm!txtFonNombre = f.FonNombre
'    frm!txtFonApe1 = f.FonApe1
'    frm!txtFonApe2 = f.FonApe2
'
'    ' --- Gestión ---
'    frm!txtFechaCalculo = f.FechaCalculo
'
'End Sub
'
'Public Sub CargarFoneticaEnFormulario(frm As Form, IDPersona As Long)
'
'    Dim f As clsFonetica
'
'    ' Cargar la última configuración fonética de la persona
'    Set f = CargarFoneticaPorPersona(IDPersona)
'
'    If Not f Is Nothing Then
'        ' Guardar en el objeto global
'        Set Fonetica = f
'
'        ' Volcar al formulario
'        CargarFormularioDesdeFonetica frm, f
'    Else
'        ' Si no hay fonética previa, limpiar campos
'        LimpiarFormularioFonetica frm
'    End If
'
'End Sub

'Public Function CargarFonetica(ByVal IDFonetica As Long) As clsFonetica
'    Dim db As DAO.Database
'    Dim rs As DAO.Recordset
'    Dim F As clsFonetica
'    Dim sql As String
'
'    sql = "SELECT * FROM Fonetica WHERE IDFonetica = " & IDFonetica
'
'    Set db = CurrentDb
'    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
'
'    If Not rs.EOF Then
'        Set F = New clsFonetica
'
'        ' --- Identificación ---
'        F.IDFonetica = rs!IDFonetica
'        F.IDPersona = rs!IDPersona
'
'        ' --- Modo ---
'        F.ModoFonetico = rs!ModoFonetico
'
'        ' --- Reglas ---
'        F.UsarHmuda = rs!UsarHmuda
'        F.UsarUmuda = rs!UsarUmuda
'
'        ' --- Idiomas (objetos) ---
'        Set F.IdiomaNombre = CargarIdioma(rs!IDIdiomaNombre)
'        Set F.IdiomaApe1 = CargarIdioma(rs!IDIdiomaApe1)
'        Set F.IdiomaApe2 = CargarIdioma(rs!IDIdiomaApe2)
'
'        ' --- Resultados fonéticos ---
'        F.FonNombre = Nz(rs!FonNombre, "")
'        F.FonApe1 = Nz(rs!FonApe1, "")
'        F.FonApe2 = Nz(rs!FonApe2, "")
'
'        ' --- Datos originales ---
'        F.NombreOriginal = Nz(rs!NombreOriginal, "")
'        F.Ape1Original = Nz(rs!Ape1Original, "")
'        F.Ape2Original = Nz(rs!Ape2Original, "")
'
'        ' --- Gestión ---
'        F.FechaCalculo = rs!FechaCalculo
'    End If
'
'    rs.Close
'    Set rs = Nothing
'    Set db = Nothing
'
'    Set CargarFonetica = F
'End Function

