Attribute VB_Name = "modFoneticaDAO"

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
    Dim F As clsFonetica
    Dim sql As String

    sql = "SELECT * FROM Fonetica WHERE IDFonetica = " & IDFonetica

    Set db = CurrentDb
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)

    If Not rs.EOF Then
        Set F = New clsFonetica

        F.IDFonetica = rs!IDFonetica
        F.IDPersona = rs!IDPersona

        F.ModoFonetico = rs!ModoFonetico
        F.UsarHmuda = rs!UsarHmuda
        F.UsarUmuda = rs!UsarUmuda

        ' Idiomas (se cargan como objetos)
        Set F.IdiomaNombre = CargarIdioma(rs!IDIdiomaNombre)
        Set F.IdiomaApe1 = CargarIdioma(rs!IDIdiomaApe1)
        Set F.IdiomaApe2 = CargarIdioma(rs!IDIdiomaApe2)

        ' Resultados fonéticos
        F.FonNombre = Nz(rs!FonNombre, "")
        F.FonApe1 = Nz(rs!FonApe1, "")
        F.FonApe2 = Nz(rs!FonApe2, "")

        ' Datos originales
        F.NombreOriginal = Nz(rs!NombreOriginal, "")
        F.Ape1Original = Nz(rs!Ape1Original, "")
        F.Ape2Original = Nz(rs!Ape2Original, "")

        F.FechaCalculo = rs!FechaCalculo
    End If

    rs.Close
    Set rs = Nothing
    Set db = Nothing

    Set CargarFonetica = F
End Function


' ============================================================
'   GUARDAR FONÉTICA (INSERTAR O ACTUALIZAR)
' ============================================================
Public Function GuardarFonetica(ByVal F As clsFonetica) As Long
    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    Set db = CurrentDb

    If F.IDFonetica = 0 Then
        ' --- INSERTAR ---
        Set rs = db.OpenRecordset("Fonetica", dbOpenDynaset)
        rs.AddNew
    Else
        ' --- ACTUALIZAR ---
        Set rs = db.OpenRecordset("SELECT * FROM Fonetica WHERE IDFonetica=" & F.IDFonetica, dbOpenDynaset)
        If rs.EOF Then
            rs.Close
            Set rs = db.OpenRecordset("Fonetica", dbOpenDynaset)
            rs.AddNew
        Else
            rs.Edit
        End If
    End If

    ' --- Campos ---
    rs!IDPersona = F.IDPersona
    rs!ModoFonetico = F.ModoFonetico
    rs!UsarHmuda = F.UsarHmuda
    rs!UsarUmuda = F.UsarUmuda

    rs!IDIdiomaNombre = F.IdiomaNombre.IDIdioma
    rs!IDIdiomaApe1 = F.IdiomaApe1.IDIdioma
    rs!IDIdiomaApe2 = F.IdiomaApe2.IDIdioma

    rs!FonNombre = F.FonNombre
    rs!FonApe1 = F.FonApe1
    rs!FonApe2 = F.FonApe2

    rs!NombreOriginal = F.NombreOriginal
    rs!Ape1Original = F.Ape1Original
    rs!Ape2Original = F.Ape2Original

    rs!FechaCalculo = F.FechaCalculo

    rs.Update

    ' Devolver ID
    If F.IDFonetica = 0 Then
        rs.Bookmark = rs.LastModified
        GuardarFonetica = rs!IDFonetica
    Else
        GuardarFonetica = F.IDFonetica
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

Public Function GuardarFoneticaSmart(F As clsFonetica) As Long
    Dim FDB As clsFonetica

    ' ¿Existe ya una configuración para esta persona?
    Set FDB = CargarFoneticaPorPersona(F.IDPersona)

    If Not FDB Is Nothing Then
        ' Si la configuración es igual ? actualizar
        If ConfiguracionIgual(F, FDB) Then
            F.IDFonetica = FDB.IDFonetica
            GuardarFoneticaSmart = GuardarFonetica(F)
            Exit Function
        End If
    End If

    ' Si no existe o ha cambiado ? insertar nuevo
    F.IDFonetica = 0
    GuardarFoneticaSmart = GuardarFonetica(F)
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


Public Function ConfiguracionIgual(F As clsFonetica, FDB As clsFonetica) As Boolean
    If F.ModoFonetico <> FDB.ModoFonetico Then GoTo Diferente
    If F.UsarHmuda <> FDB.UsarHmuda Then GoTo Diferente
    If F.UsarUmuda <> FDB.UsarUmuda Then GoTo Diferente

    If F.IdiomaNombre.IDIdioma <> FDB.IdiomaNombre.IDIdioma Then GoTo Diferente
    If F.IdiomaApe1.IDIdioma <> FDB.IdiomaApe1.IDIdioma Then GoTo Diferente
    If F.IdiomaApe2.IDIdioma <> FDB.IdiomaApe2.IDIdioma Then GoTo Diferente

    ConfiguracionIgual = True
    Exit Function

Diferente:
    ConfiguracionIgual = False
End Function

'-----------------------------------------------------------------------------------------------------

Public Sub CargarClaseFoneticaDesdeFormulario(frm As Form)
    ' Asegurar instancia
    If Fonetica Is Nothing Then Set Fonetica = New clsFonetica

    ' --- Identificación ---
    Fonetica.IDPersona = frm!IDPersona

    ' --- Modo ---
    Fonetica.ModoFonetico = IIf(frm!chkModoFon, 1, 0)

    ' --- Reglas ---
    Fonetica.UsarHmuda = frm!chkHmuda
    Fonetica.UsarUmuda = frm!chkUmuda

    ' --- Idiomas ---
    Set Fonetica.IdiomaNombre = CargarIdioma(frm!cboIdiomaNombre)
    Set Fonetica.IdiomaApe1 = CargarIdioma(frm!cboIdiomaApe1)
    Set Fonetica.IdiomaApe2 = CargarIdioma(frm!cboIdiomaApe2)

    ' --- Datos originales ---
    Fonetica.NombreOriginal = frm!txtNombre
    Fonetica.Ape1Original = frm!txtApe1
    Fonetica.Ape2Original = frm!txtApe2

    ' --- Resultados fonéticos ---
    Fonetica.FonNombre = MotorFonetico_Convertir(frm!txtNombre, Fonetica.IdiomaNombre, Fonetica.UsarHmuda, Fonetica.UsarUmuda, Fonetica.ModoFonetico)
    Fonetica.FonApe1 = MotorFonetico_Convertir(frm!txtApe1, Fonetica.IdiomaApe1, Fonetica.UsarHmuda, Fonetica.UsarUmuda, Fonetica.ModoFonetico)
    Fonetica.FonApe2 = MotorFonetico_Convertir(frm!txtApe2, Fonetica.IdiomaApe2, Fonetica.UsarHmuda, Fonetica.UsarUmuda, Fonetica.ModoFonetico)

    ' --- Gestión ---
    Fonetica.FechaCalculo = Now
End Sub

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

Public Sub CargarFormularioDesdeFonetica(frm As Form, F As clsFonetica)

    ' --- Modo ---
    frm!chkModoFon = (F.ModoFonetico = 1)

    ' --- Reglas ---
    frm!chkHmuda = F.UsarHmuda
    frm!chkUmuda = F.UsarUmuda

    ' --- Idiomas ---
    frm!cboIdiomaNombre = F.IdiomaNombre.IDIdioma
    frm!cboIdiomaApe1 = F.IdiomaApe1.IDIdioma
    frm!cboIdiomaApe2 = F.IdiomaApe2.IDIdioma

    ' --- Datos originales ---
    frm!txtNombre = F.NombreOriginal
    frm!txtApe1 = F.Ape1Original
    frm!txtApe2 = F.Ape2Original

    ' --- Resultados fonéticos ---
    frm!txtFonNombre = F.FonNombre
    frm!txtFonApe1 = F.FonApe1
    frm!txtFonApe2 = F.FonApe2

    ' --- Gestión ---
    frm!txtFechaCalculo = F.FechaCalculo

End Sub

Public Sub CargarFoneticaEnFormulario(frm As Form, IDPersona As Long)

    Dim F As clsFonetica

    ' Cargar la última configuración fonética de la persona
    Set F = CargarFoneticaPorPersona(IDPersona)

    If Not F Is Nothing Then
        ' Guardar en el objeto global
        Set Fonetica = F

        ' Volcar al formulario
        CargarFormularioDesdeFonetica frm, F
    Else
        ' Si no hay fonética previa, limpiar campos
        LimpiarFormularioFonetica frm
    End If

End Sub

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

