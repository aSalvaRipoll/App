Attribute VB_Name = "modBBDD"
' ------------------------------------------------------
' Nombre:    modBBDD
' Tipo:      Módulo
' Propósito:
' Autor:     asalv
' Fecha:     15/01/2026
' ------------------------------------------------------

Option Compare Database
Option Explicit

Dim db As DAO.Database

Sub AgregaRaiz()

    Dim rsOri As DAO.Recordset
    Dim rsDest As DAO.Recordset
    Dim C1, C2
    
    Set rsOri = CurrentDb.OpenRecordset("Consulta3")
    
    While Not rsOri.EOF
        Set rsDest = CurrentDb.OpenRecordset("SELECT * FROM tbmEquivNombre WHERE Raiz is Null and NombreOriginal = '" & rsOri!NombreOriginal & "'")
                
        Debug.Print rsOri!NombreOriginal; "--> "; rsOri!raiz
        
        While Not rsDest.EOF
        
            C1 = C1 + 1
        
            DoEvents
            rsDest.Edit
            rsDest!raiz = rsOri!raiz
            rsDest.Update
        
            rsDest.MoveNext
        Wend
        
        Set rsDest = CurrentDb.OpenRecordset("SELECT * FROM tbmEquivNombre WHERE Raiz is Null and NombreEquivalente = '" & rsOri!NombreOriginal & "'")
        While Not rsDest.EOF
        
            C2 = C2 + 1
        
            DoEvents
            rsDest.Edit
            rsDest!raiz = rsOri!raiz
            rsDest.Update
        
            rsDest.MoveNext
        Wend
        rsOri.MoveNext
    Wend
    
Debug.Print "Originales:"; C1
Debug.Print "Equivalentes:"; C2


End Sub
    
    
'Sub SacaRaiz()
'
'    Dim rs As DAO.Recordset
'    Dim cont
'
'    Set rs = CurrentDb.OpenRecordset("tbmEquivNombre")
'
'    cont = 0
'    While Not rs.EOF
'    DoEvents
'    cont = cont + 1
'    Debug.Print cont; " de "; rs.RecordCount; " ";
'        If InStr(rs("Notas"), "Raíz") Then
'            Debug.Print rs!NombreOriginal; " "; rs!IdiomaOriginal; " ->"; rs!Notas;
'            rs.Edit
'
'            rs("Raiz") = rs("Notas")
'
'            rs.Update
'        ElseIf InStr(rs("Notas"), "Origen") Then
'            Debug.Print rs!NombreOriginal; " "; rs!IdiomaOriginal; " ->"; rs!Notas;
'            rs.Edit
'
'            rs("Raiz") = rs("Notas")
'
'            rs.Update
'        End If
'        Debug.Print
'        rs.MoveNext
'
'    Wend
'
'Debug.Print
'Debug.Print "fin"
'End Sub

Sub ReparaApellidos()

    Dim rs As DAO.Recordset
    Dim tx As String
    
    Set rs = CurrentDb.OpenRecordset("Apellidos")
    
    While Not rs.EOF
        DoEvents
        tx = LimpiarApellido(rs("Apellido"))
        rs.Edit
        rs("Apellido") = tx 'StrConv(rs("palabra"), vbProperCase)
    
        rs.Update
    
        rs.MoveNext
    Wend
    Debug.Print "fin"
End Sub




Function LimpiarApellido(ByVal txt As String) As String
    Dim i As Long
    Dim ch As String
    
    For i = 1 To Len(txt)
        ch = Mid$(txt, i, 1)
        
        ' Si encuentra un número, corta antes de él
        If ch Like "#" Then
            LimpiarApellido = Left$(txt, i - 1)
            Exit Function
        End If
    Next i
    
    ' Si no hay números, devuelve el texto original
    LimpiarApellido = txt
End Function


'Sub CargarTablaIdiomas()
'
'CurrentDb.Execute "DELETE FROM tbmIdiomas"
'
'CurrentDb.Execute "INSERT INTO tbmIdiomas (IDIdioma, Abreviado, NomIdioma, Notas) VALUES " & vbCrLf & _
'                                         "(1,'es','Castellano','')"
'
'CurrentDb.Execute "INSERT INTO tbmIdiomas (IDIdioma, Abreviado, NomIdioma, Notas) VALUES " & vbCrLf & _
'                                         "(2,  'ca',     'Català',                    '')"
'CurrentDb.Execute "INSERT INTO tbmIdiomas (IDIdioma, Abreviado, NomIdioma, Notas) VALUES " & vbCrLf & _
'    "(3,  'ca-ib',  'Mallorquí',                 'Idioma balear')"
'CurrentDb.Execute "INSERT INTO tbmIdiomas (IDIdioma, Abreviado, NomIdioma, Notas) VALUES " & vbCrLf & _
'"(4,  'ca-va',  'Valencià',                  'Norma valenciana')"
'CurrentDb.Execute "INSERT INTO tbmIdiomas (IDIdioma, Abreviado, NomIdioma, Notas) VALUES " & vbCrLf & _
'"(5,  'eu',     'Euskara',                   '')"
'CurrentDb.Execute "INSERT INTO tbmIdiomas (IDIdioma, Abreviado, NomIdioma, Notas) VALUES " & vbCrLf & _
'"(6,  'gl',     'Galego',                    '')"
'CurrentDb.Execute "INSERT INTO tbmIdiomas (IDIdioma, Abreviado, NomIdioma, Notas) VALUES " & vbCrLf & _
'"(7,  'pt',     'Português',                 'Compatibilidad; usa PT-EU')"
'CurrentDb.Execute "INSERT INTO tbmIdiomas (IDIdioma, Abreviado, NomIdioma, Notas) VALUES " & vbCrLf & _
'"(8,  'pt-eu',  'Português (Portugal)',      'Variante europea')"
'CurrentDb.Execute "INSERT INTO tbmIdiomas (IDIdioma, Abreviado, NomIdioma, Notas) VALUES " & vbCrLf & _
'"(9,  'pt-br',  'Português (Brasil)',        'Variante brasileira')"
'CurrentDb.Execute "INSERT INTO tbmIdiomas (IDIdioma, Abreviado, NomIdioma, Notas) VALUES " & vbCrLf & _
'"(10, 'fr',     'Français',                  '')"
'CurrentDb.Execute "INSERT INTO tbmIdiomas (IDIdioma, Abreviado, NomIdioma, Notas) VALUES " & vbCrLf & _
'"(11, 'en',     'English',                   '')"
'
'End Sub
    
    
'Sub arreglaTabla()
'
'    Dim rs As DAO.Recordset
'    Dim entrada, n
'
'
'    Set rs = CurrentDb.OpenRecordset("tbmFoneticaCompleta")
'
'
'    While Not rs.EOF
'        rs.Edit
'        entrada = rs!idFonemaOri
'
'        For n = 1 To Len(entrada)
'            If Not Mid(entrada, n, 1) Like "[1234567890]" Then
'                Stop
'            End If
'
'            Debug.Print ">"; Mid(entrada, n, 1); ">";
'            Debug.Print Asc(Mid(entrada, n, 1))
'        Next
'
'        rs!idFonema = rs!idFonemaOri
'
'        rs.Update
'        rs.MoveNext
'    Wend
'
'
'
'End Sub
    
'Public Sub GuardarDatosPersona()
'    On Error GoTo ErrHandler
'
'    Dim rs As DAO.Recordset
'    Dim nuevoID As Long
'    Dim nombreCompleto As String
'    Dim nombreFonetico As String
'
'    '===============================
'    ' VALIDACIÓN BÁSICA
'    '===============================
'    If Trim$(Me.txtNombreBase & "") = "" Then
'        MsgBox "El nombre es obligatorio.", vbInformation, "Guardar datos"
'        Me.txtNombreBase.SetFocus
'        Exit Sub
'
'    ElseIf Trim$(Me.txtApeBase1 & "") = "" Then
'        MsgBox "El primer apellido es obligatorio.", vbInformation, "Guardar datos"
'        Me.txtApeBase1.SetFocus
'        Exit Sub
'    End If
'
'    '===============================
'    ' GENERAR CAMPOS COMPUESTOS
'    '===============================
'    nombreCompleto = Trim$(Me.txtNombreBase & " " & Me.txtApeBase1 & " " & Nz(Me.txtApeBase2, ""))
'    nombreFonetico = Trim$(Nz(Me.FonNombre, "") & " " & Nz(Me.FonApe1, "") & " " & Nz(Me.FonApe2, ""))
'
'    '===============================
'    ' GUARDAR EN TABLA
'    '===============================
'    Set db = CurrentDb
'    Set rs = db.OpenRecordset("tbuPersonas", dbOpenDynaset)
'
'    rs.AddNew
'
'        '--- Datos base ---
'        rs!Nombre = Nz(Me.txtNombreBase, "")
'        rs!Ape1 = Nz(Me.txtApeBase1, "")
'        rs!Ape2 = Nz(Me.txtApeBase2, "")
'
'        '--- Idiomas por campo (FK numérica) ---
'        If Not Me.IdiomaNom Is Nothing Then rs!IdiomaNombre = Me.IdiomaNom.ID
'        If Not Me.IdiomaApe1 Is Nothing Then rs!IdiomaApe1 = Me.IdiomaApe1.ID
'        If Not Me.IdiomaApe2 Is Nothing Then rs!IdiomaApe2 = Me.IdiomaApe2.ID
'
'        '--- Fecha nacimiento ---
'        If Not IsNull(Me.txtFechaNac) Then
'            rs!FechaNacimiento = Me.txtFechaNac
'        End If
'
'        '--- Sexo legal ---
'        If Not IsNull(Me.cboSexoLegal) Then
'            rs!Genero = Me.cboGenero
'        End If
'
'        '--- Fonéticos por campo ---
'        rs!FonNombre = Nz(Me.FonNombre, "")
'        rs!FonApe1 = Nz(Me.FonApe1, "")
'        rs!FonApe2 = Nz(Me.FonApe2, "")
'
''----------------------------------------------------------------
''          No hace falta, se puede reconstruir posteriormente
'        '--- Campos compuestos ---
''        rs!nombreCompleto = nombreCompleto
''        rs!nombreFonetico = nombreFonetico
''----------------------------------------------------------------
'
'        '--- Sistema de cálculo ---
'        rs!SistemaCalculo = Me.fraSistema
'
'        '--- Opciones fonéticas ---
'        rs!UsarHmuda = Me.chkHmuda
'        rs!UsarUmuda = Me.chkUmuda
'
'        '--- Auditoría ---
'        rs!FechaAlta = Now()
'        rs!FechaModificacion = Now()
'
'    rs.Update
'
'    nuevoID = rs!IDPersona
'
'    rs.Close
'    Set rs = Nothing
'    Set db = Nothing
'
'    MsgBox "Datos guardados correctamente. ID asignado: " & nuevoID, _
'           vbInformation, "Guardar datos"
'    Exit Sub
'
''===============================
'' GESTIÓN DE ERRORES
''===============================
'ErrHandler:
'    MsgBox "Error al guardar los datos: " & Err.Description, _
'           vbExclamation, "Error"
'
'End Sub

Public Function GuardarPersonaEnTabla() As Integer

    Dim rs As DAO.Recordset
    Dim sql As String
    Dim nueva As Boolean

    GuardarPersonaEnTabla = 0
    ' ------------------------------------------------------------
    ' 1. Determinar si es nueva persona o actualización
    ' ------------------------------------------------------------
    If Persona.ID_Persona > 0 Then
        Select Case MsgBoxEx2("¿Desea actualizar el registro" & vbCrLf & _
                    "o desea crear uno nuevo?", _
                    vbQuestion + vbYesNoCancel + vbDefaultButton2, _
                    "Datos persona", , "&Cancelar", , , , "C&rear", "&Actualizar")
            
            Case vbYes
                Persona.ID_Persona = 0
                GuardarPersonaEnTabla = 1
            Case vbNo
                GuardarPersonaEnTabla = 2
            Case Else
                Exit Function
        End Select
        
    End If
    
    
    
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
        Persona.ID_Persona = AutoNext("ID_Persona", "tbuPersonas")
        rs!ID_Persona = Persona.ID_Persona
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
    ' 4. Guardar idiomas (OBJETO --> ID)
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
'    rs!FonNombre = Persona.FonNombre
'    rs!FonApe1 = Persona.FonApe1
'    rs!FonApe2 = Persona.FonApe2

    ' ------------------------------------------------------------
    ' 6. Guardar parámetros de análisis
    ' ------------------------------------------------------------
'    rs!Sistema = Persona.Sistema
'    rs!Tarot = Persona.Tarot
'    rs!UsarHmuda = Persona.UsarHmuda
'    rs!UsarUmuda = Persona.UsarUmuda

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
'    If nueva Then
'        Persona.ID_Persona = rs!ID_Persona
'    End If

    rs.Close
    Set rs = Nothing
    
    'GuardarPersonaEnTabla = GuardarPersonaEnTabla + 10
    
End Function


Public Sub CargaPersonaDesdeTabla(ByVal id As Long)

    Dim rs As DAO.Recordset
    Dim sql As String

    If Persona Is Nothing Then Set Persona = New clsPersona
    
    sql = "SELECT * FROM tbuPersonas WHERE ID_Persona = " & id & ";"
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

'        Set Persona.IdiomaNombre = CargarIdiomaDesdeID(Nz(rs!IdiomaNombre, 0))
'        Set Persona.IdiomaApe1 = CargarIdiomaDesdeID(Nz(rs!IdiomaApe1, 0))
'        Set Persona.IdiomaApe2 = CargarIdiomaDesdeID(Nz(rs!IdiomaApe2, 0))

        Set Persona.IdiomaNombre = colIdiomas(CStr(Nz(rs!IdiomaNombre, 0)))
        Set Persona.IdiomaApe1 = colIdiomas(CStr(Nz(rs!IdiomaApe1, 0)))
        Set Persona.IdiomaApe2 = colIdiomas(CStr(Nz(rs!IdiomaApe2, 0)))

        ' --- Fonética ---
'        Persona.FonNombre = Nz(rs!FonNombre, "")
'        Persona.FonApe1 = Nz(rs!FonApe1, "")
'        Persona.FonApe2 = Nz(rs!FonApe2, "")

        ' --- Parámetros de análisis ---
'        Persona.Sistema = Nz(rs!Sistema, 0)
'        Persona.Tarot = Nz(rs!Tarot, 0)
'        Persona.UsarHmuda = Nz(rs!UsarHmuda, False)
'        Persona.UsarUmuda = Nz(rs!UsarUmuda, False)

        ' --- Gestión ---
        Persona.FechaAlta = Nz(rs!FechaAlta, 0)
        Persona.FechaModificacion = Nz(rs!FechaModificacion, 0)

    End If

    rs.Close
    Set rs = Nothing

End Sub

'===================================================================================================

' ============================================================
'   Guarda o actualiza una conversión fonética en la tabla
' ============================================================
Public Sub GuardarConversionFonetica(f As clsFonetica)

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String

    Set db = CurrentDb

    ' Si IDFonetica = 0 --> es un registro nuevo
    If f.IDFonetica = 0 Then
        sql = "SELECT * FROM tbuFonetica WHERE 1=0"
        Set rs = db.OpenRecordset(sql, dbOpenDynaset)
        rs.AddNew
        f.IDFonetica = AutoNext("IDFonetica", "tbuFonetica", "idPersona = " & f.IDPersona)
        rs!IDFonetica = f.IDFonetica
    Else
        sql = "SELECT * FROM tbuFonetica WHERE IDFonetica = " & f.IDFonetica & " And IDPersona = " & f.IDPersona
        Set rs = db.OpenRecordset(sql, dbOpenDynaset)

        If rs.EOF Then
            ' Si no existe, lo tratamos como nuevo
            rs.AddNew
            f.IDFonetica = AutoNext("IDFonetica", "tbuFonetica", "IDPersona = " & f.IDPersona)
            rs!IDFonetica = f.IDFonetica
        Else
            rs.Edit
        End If
    End If

    ' --- Volcado de datos ---
    rs!IDPersona = f.IDPersona
    rs!SistemaFonetico = f.SistemaFonetico
    rs!VersionMotor = f.VersionMotor

    rs!IdiomaNombre = f.IdiomaNombre
    rs!IdiomaApe1 = f.IdiomaApe1
    rs!IdiomaApe2 = f.IdiomaApe2

    rs!FonNombre = f.FonNombre
    rs!FonApe1 = f.FonApe1
    rs!FonApe2 = f.FonApe2

    rs!FechaCalculo = f.FechaCalculo
    rs!Activo = f.Activo

    rs.Update

    ' Si era nuevo, asignamos el ID generado
'    If f.IDFonetica = 0 Then
'        rs.Bookmark = rs.LastModified
'        f.IDFonetica = rs!IDFonetica
'    End If

    rs.Close
    Set rs = Nothing
    Set db = Nothing

End Sub

' ============================================================
'   CargaConversionFonetica — Carga datos en el objeto público Fonetica
' ============================================================
Public Sub CargarConversionFonetica(ByVal IDFonetica As Long, ByVal IDPersona As Long)

    Dim rs As DAO.Recordset
    Dim sql As String

    sql = "SELECT * FROM tbuFonetica WHERE IDFonetica = " & IDFonetica & " AND IDPersona = " & IDPersona ';"
    Set rs = CurrentDb.OpenRecordset(sql, dbOpenDynaset)

    If Not rs.EOF Then

        ' --- Identificación ---
        Fonetica.IDFonetica = Nz(rs!IDFonetica, 0)
        Fonetica.IDPersona = Nz(rs!IDPersona, 0)

        ' --- Sistema ---
        Fonetica.SistemaFonetico = Nz(rs!SistemaFonetico, 0)

        ' --- Idiomas ---
        Fonetica.IdiomaNombre = Nz(rs!IdiomaNombre, "")
        Fonetica.IdiomaApe1 = Nz(rs!IdiomaApe1, "")
        Fonetica.IdiomaApe2 = Nz(rs!IdiomaApe2, "")

        ' --- Resultados fonéticos ---
        Fonetica.FonNombre = Nz(rs!FonNombre, "")
        Fonetica.FonApe1 = Nz(rs!FonApe1, "")
        Fonetica.FonApe2 = Nz(rs!FonApe2, "")

        ' --- Gestión ---
        Fonetica.FechaCalculo = Nz(rs!FechaCalculo, 0)
        Fonetica.Activo = Nz(rs!Activo, True)

    End If

    rs.Close
    Set rs = Nothing

End Sub
'===================================================================================================

'' ============================================================
''   Guarda resultados en la tabla
'' ============================================================
'
'Public Sub GuardarResultado(r As clsResultado)
'
'    Dim db As DAO.Database
'    Dim rs As DAO.Recordset
'    Dim SQL As String
'
'    Set db = CurrentDb
'
'    ' Si IDResultado = 0 --> es un registro nuevo
'    If r.idResultado = 0 Then
'        SQL = "SELECT * FROM tbuResultados WHERE 1=0"
'        Set rs = db.OpenRecordset(SQL, dbOpenDynaset)
'        rs.AddNew
'    Else
'        SQL = "SELECT * FROM tbuResultados WHERE IDResultado = " & r.idResultado
'        Set rs = db.OpenRecordset(SQL, dbOpenDynaset)
'
'        If rs.EOF Then
'            ' Si no existe, lo tratamos como nuevo
'            rs.AddNew
'            ' Si quieres asignar ID manual:
'             'r.idResultado = AutoNext("IDResultado", "tbuResultados")
'             'rs!idResultado = r.idResultado
'        Else
'            rs.Edit
'        End If
'    End If
'
'    ' --- Volcado de datos ---
'    rs!IDPersona = r.IDPersona
'    rs!IDFonetica = r.IDFonetica
'    rs!FechaCalculo = Now
'
'    ' Números principales
'    rs!NumeroDestino = r.NumeroDestino
'    rs!NumeroAlma = r.NumeroAlma
'    rs!NumeroPersonalidad = r.NumeroPersonalidad
'    rs!NumeroCaminoVida = r.NumeroCaminoVida
'    rs!NumeroMadurez = r.NumeroMadurez
'
'    ' Temporales
'    rs!AnioPersonal = r.AnioPersonal
''    rs!MesPersonal = r.MesPersonal
''    rs!DiaPersonal = r.DiaPersonal
'    rs!EdadPersonal = r.NumeroEdadPersonal
'
''    ' Ciclos, pináculos, desafíos
''    rs!CicloActual = r.CicloActual
''    rs!PinaculoActual = r.PinaculoActual
''    rs!DesafioActual = r.DesafíoActual
'
'    ' Planos de expresión
'    rs!PlanoFisico = r.PlanoFisico
'    rs!PlanoEmocional = r.PlanoEmocional
'    rs!PlanoMental = r.PlanoMental
'    rs!PlanoIntuitivo = r.PlanoIntuitivo
'
'    ' Letras y símbolos
'    rs!PiedraAngular = r.PiedraAngular
'    rs!PiedraToque = r.PiedraToque
'    rs!primeraLetra = r.primeraLetra
'    rs!primeraVocal = r.primeraVocal
'    rs!primeraConsonante = r.primeraConsonante
'
'    ' Otros
'    rs!RespuestaSubconsciente = r.NumeroRespuestaSubconsciente
'    rs!Poder = r.NumeroPoder
''    rs!DeudaKarmica = r.DeudaKarmica
'
'    rs.Update
'
'    ' Si era nuevo, asignamos el ID generado por Access
'    If r.idResultado = 0 Then
'        r.idResultado = rs!idResultado
'    End If
'
'    rs.Close
'    Set rs = Nothing
'    Set db = Nothing
'
'End Sub
'
'Public Function CargarResultado(idResultado As Long) As clsResultado
'
'    Dim r As New clsResultado
'    Dim rs As DAO.Recordset
'    Dim SQL As String
'
'    SQL = "SELECT * FROM tbuResultados WHERE IDResultado = " & idResultado
'    Set rs = CurrentDb.OpenRecordset(SQL, dbOpenDynaset)
'
'    If rs.EOF Then
'        Set CargarResultado = Nothing
'        Exit Function
'    End If
'
'    r.idResultado = rs!idResultado
'    r.IDPersona = rs!IDPersona
'    r.IDFonetica = rs!IDFonetica
'
'    r.NumeroDestino = rs!NumeroDestino
'    r.NumeroAlma = rs!NumeroAlma
'    r.NumeroPersonalidad = rs!NumeroPersonalidad
'    r.NumeroCaminoVida = rs!NumeroCaminoVida
'    r.NumeroMadurez = rs!NumeroMadurez
'
'    r.AnioPersonal = rs!AnioPersonal
''    r.MesPersonal = rs!MesPersonal
''    r.DiaPersonal = rs!DiaPersonal
'    r.NumeroEdadPersonal = rs!EdadPersonal
'
''    r.CicloActual = rs!CicloActual
''    r.PinaculoActual = rs!PinaculoActual
''    r.DesafíoActual = rs!DesafioActual
'
'    r.PlanoFisico = rs!PlanoFisico
'    r.PlanoEmocional = rs!PlanoEmocional
'    r.PlanoMental = rs!PlanoMental
'    r.PlanoIntuitivo = rs!PlanoIntuitivo
'
'    r.PiedraAngular = rs!PiedraAngular
'    r.PiedraToque = rs!PiedraToque
'    r.primeraLetra = rs!primeraLetra
'    r.primeraVocal = rs!primeraVocal
'    r.primeraConsonante = rs!primeraConsonante
'
'    r.NumeroRespuestaSubconsciente = rs!RespuestaSubconsciente
'    r.NumeroPoder = rs!Poder
''    r.DeudaKarmica = rs!DeudaKarmica
'
'    rs.Close
'    Set CargarResultado = r
'
'End Function
'
'' ============================================================
''   Guarda ciclos en la tabla
'' ============================================================
'
'Public Sub GuardarCiclos(c As clsCiclos)
'
'    Dim db As DAO.Database
'    Dim rs As DAO.Recordset
'    Dim SQL As String
'
'    Set db = CurrentDb
'
'    ' Si IDCiclo = 0 --> es un registro nuevo
'    If c.idCiclo = 0 Then
'        SQL = "SELECT * FROM tbuCiclos WHERE 1=0"
'        Set rs = db.OpenRecordset(SQL, dbOpenDynaset)
'        rs.AddNew
'    Else
'        SQL = "SELECT * FROM tbuCiclos WHERE IDCiclo = " & c.idCiclo
'        Set rs = db.OpenRecordset(SQL, dbOpenDynaset)
'
'        If rs.EOF Then
'            ' Si no existe, lo tratamos como nuevo
'            rs.AddNew
'            ' Si quieres asignar ID manual:
'            ' c.IDCiclo = AutoNext("IDCiclo", "tbuCiclos")
'            ' rs!IDCiclo = c.IDCiclo
'        Else
'            rs.Edit
'        End If
'    End If
'
'    ' --- Volcado de datos ---
'    rs!idResultado = c.idResultado
'    rs!IDPersona = c.IDPersona
'
'    rs!NumCiclos = c.NumCiclos
'    rs!MetodoCiclos = c.MetodoCiclos
'
'    rs!Ciclo1 = c.Ciclo1
'    rs!EdadIni1 = c.EdadIni1
'    rs!EdadFin1 = c.EdadFin1
'
'    rs!Ciclo2 = c.Ciclo2
'    rs!EdadIni2 = c.EdadIni2
'    rs!EdadFin2 = c.EdadFin2
'
'    rs!Ciclo3 = c.Ciclo3
'    rs!EdadIni3 = c.EdadIni3
'    rs!EdadFin3 = c.EdadFin3
'
'    rs!Ciclo4 = c.Ciclo4
'    rs!EdadIni4 = c.EdadIni4
'    rs!EdadFin4 = c.EdadFin4
'
'    rs.Update
'
'    ' Si era nuevo, asignamos el ID generado por Access
'    'If c.idCiclo = 0 Then
'    '    c.idCiclo = rs!idCiclo
'    'End If
'
'    rs.Close
'    Set rs = Nothing
'    Set db = Nothing
'
'End Sub
'
'Public Function CargarCiclos(idCiclo As Long) As clsCiclos
'
'    Dim c As New clsCiclos
'    Dim rs As DAO.Recordset
'    Dim SQL As String
'
'    SQL = "SELECT * FROM tbuCiclos WHERE IDCiclo = " & idCiclo
'    Set rs = CurrentDb.OpenRecordset(SQL, dbOpenDynaset)
'
'    If rs.EOF Then
'        Set CargarCiclos = Nothing
'        Exit Function
'    End If
'
'    c.idCiclo = rs!idCiclo
'    c.idResultado = rs!idResultado
'    c.IDPersona = rs!IDPersona
'
'    c.NumCiclos = rs!NumCiclos
'    c.MetodoCiclos = rs!MetodoCiclos
'
'    c.Ciclo1 = rs!Ciclo1
'    c.EdadIni1 = rs!EdadIni1
'    c.EdadFin1 = rs!EdadFin1
'
'    c.Ciclo2 = rs!Ciclo2
'    c.EdadIni2 = rs!EdadIni2
'    c.EdadFin2 = rs!EdadFin2
'
'    c.Ciclo3 = rs!Ciclo3
'    c.EdadIni3 = rs!EdadIni3
'    c.EdadFin3 = rs!EdadFin3
'
'    c.Ciclo4 = rs!Ciclo4
'    c.EdadIni4 = rs!EdadIni4
'    c.EdadFin4 = rs!EdadFin4
'
'    rs.Close
'    Set CargarCiclos = c
'
'End Function
'
'
'' ============================================================
''   Guarda pináculos y desafíos en la tabla
'' ============================================================
'
'Public Sub GuardarPinaDes(pd As clsPinaDes)
'
'    Dim db As DAO.Database
'    Dim rs As DAO.Recordset
'    Dim SQL As String
'
'    Set db = CurrentDb
'
'    ' Si IDPinaDes = 0 --> es un registro nuevo
'    If pd.idPinaDes = 0 Then
'        SQL = "SELECT * FROM tbuPinaDes WHERE 1=0"
'        Set rs = db.OpenRecordset(SQL, dbOpenDynaset)
'        rs.AddNew
'    Else
'        SQL = "SELECT * FROM tbuPinaDes WHERE IDPinaDes = " & pd.idPinaDes
'        Set rs = db.OpenRecordset(SQL, dbOpenDynaset)
'
'        If rs.EOF Then
'            ' Si no existe, lo tratamos como nuevo
'            rs.AddNew
'            ' Si quieres asignar ID manual:
'            ' pd.IDPinaDes = AutoNext("IDPinaDes", "tbuPinaDes")
'            ' rs!IDPinaDes = pd.IDPinaDes
'        Else
'            rs.Edit
'        End If
'    End If
'
'    ' --- Volcado de datos ---
'    rs!idResultado = pd.idResultado
'    rs!IDPersona = pd.IDPersona
'
'    rs!Pina1 = pd.Pina1
'    rs!Pina2 = pd.Pina2
'    rs!Pina3 = pd.Pina3
'    rs!Pina4 = pd.Pina4
'
'    rs!Desa1 = pd.Desa1
'    rs!Desa2 = pd.Desa2
'    rs!Desa3 = pd.Desa3
'    rs!Desa4 = pd.Desa4
'
'    rs!fIni1 = pd.EdadIni1
'    rs!fIni2 = pd.EdadIni2
'    rs!fIni3 = pd.EdadIni3
'    rs!fIni4 = pd.EdadIni4
'
'    rs!fFin1 = pd.EdadFin1
'    rs!fFin2 = pd.EdadFin2
'    rs!fFin3 = pd.EdadFin3
'    rs!fFin4 = pd.EdadFin4
'
'    rs.Update
'
'    ' Si era nuevo, asignamos el ID generado por Access
'    'If pd.idPinaDes = 0 Then
'    '    pd.idPinaDes = rs!idPinaDes
'    'End If
'
'    rs.Close
'    Set rs = Nothing
'    Set db = Nothing
'
'End Sub
'
'Public Function CargarPinaDes(idPinaDes As Long) As clsPinaDes
'
'    Dim pd As New clsPinaDes
'    Dim rs As DAO.Recordset
'    Dim SQL As String
'
'    SQL = "SELECT * FROM tbuPinaDes WHERE IDPinaDes = " & idPinaDes
'    Set rs = CurrentDb.OpenRecordset(SQL, dbOpenDynaset)
'
'    If rs.EOF Then
'        Set CargarPinaDes = Nothing
'        Exit Function
'    End If
'
'    pd.idPinaDes = rs!idPinaDes
'    pd.idResultado = rs!idResultado
'    pd.IDPersona = rs!IDPersona
'
'    pd.Pina1 = rs!Pina1
'    pd.Pina2 = rs!Pina2
'    pd.Pina3 = rs!Pina3
'    pd.Pina4 = rs!Pina4
'
'    pd.Desa1 = rs!Desa1
'    pd.Desa2 = rs!Desa2
'    pd.Desa3 = rs!Desa3
'    pd.Desa4 = rs!Desa4
'
'    pd.EdadIni1 = rs!fIni1
'    pd.EdadIni2 = rs!fIni2
'    pd.EdadIni3 = rs!fIni3
'    pd.EdadIni4 = rs!fIni4
'
'    pd.EdadFin1 = rs!fFin1
'    pd.EdadFin2 = rs!fFin2
'    pd.EdadFin3 = rs!fFin3
'    pd.EdadFin4 = rs!fFin4
'
'    rs.Close
'    Set CargarPinaDes = pd
'
'End Function
'
'' ============================================================
''   Guarda Transitos en la tabla
'' ============================================================
'
'Public Sub GuardarTransitosEnLote(col As Collection)
'
'    Dim i As Integer
'    Dim tr As clsTransitos
'    Dim idTransito As Long
'
'    idTransito = AutoNext("IDTransito", "tbuTransitos")
'
'    For i = 1 To col.Count
'        Set tr = col(i)
'        tr.idTransito = idTransito
'        Call GuardarTransito(tr)
'    Next i
'
'End Sub
'
'Public Function GuardarTransito(t As clsTransito) As Long
'
'    Dim db As DAO.Database
'    Dim rs As DAO.Recordset
'    Dim SQL As String
'
'    Set db = CurrentDb
'
'    ' Si IDTransito = 0 --> es un registro nuevo
'    If t.idTransito = 0 Then
'        SQL = "SELECT * FROM tbuTransitos WHERE 1=0"
'        Set rs = db.OpenRecordset(SQL, dbOpenDynaset)
'        t.idTransito = AutoNext("IDTransito", "tbuTransitos")
'        rs.AddNew
'
'    Else
'        SQL = "SELECT * FROM tbuTransitos WHERE IDTransito = " & t.idTransito & " AND Orden = " & t.Orden
'        Set rs = db.OpenRecordset(SQL, dbOpenDynaset)
'
'        If rs.EOF Then
'            ' No existe ? lo tratamos como nuevo
'            rs.AddNew
'            ' Si quieres asignar ID manual:
'             rs!idTransito = t.idTransito
'        Else
'            rs.Edit
'        End If
'    End If
'
'    ' --- Volcado de datos ---
'    rs!IDPersona = t.IDPersona
'    rs!idResultado = t.idResultado
'    rs!Orden = t.Orden
'
'    rs!anio = t.anio
'    rs!Edad = t.Edad
'
'    rs!Fisico = t.Fisico
'    rs!LetraFisico = t.LetraFisico
'
'    rs!Mental = t.Mental
'    rs!LetraMental = t.LetraMental
'
'    rs!Emocional = t.Emocional
'    rs!LetraEmocional = t.LetraEmocional
'
'    rs!Espiritual = t.Espiritual
'    rs!LetraEspiritual = t.LetraEspiritual
'
'    rs!Esencia = t.Esencia
'    rs!AnioPersonal = t.AnioPersonal
'
'    rs.Update
'
'    ' Si era nuevo, asignamos el ID generado por Access
''    If t.idTransito = 0 Then
''        t.idTransito = rs!idTransito
''    End If
'    GuardarTransito = t.idTransito
'    rs.Close
'    Set rs = Nothing
'    Set db = Nothing
'
'End Function
'
'Public Function CargarTransito(idTransito As Long) As clsTransitos
'
'    Dim t As New clsTransitos
'    Dim rs As DAO.Recordset
'    Dim SQL As String
'
'    SQL = "SELECT * FROM tbuTransitos WHERE IDTransito = " & idTransito
'    Set rs = CurrentDb.OpenRecordset(SQL, dbOpenDynaset)
'
'    If rs.EOF Then
'        Set CargarTransito = Nothing
'        Exit Function
'    End If
'
'    t.idTransito = rs!idTransito
'    t.IDPersona = rs!IDPersona
'    t.idResultado = rs!idResultado
'    t.Orden = rs!Orden
'
'    t.anio = rs!anio
'    t.Edad = rs!Edad
'
'    t.Fisico = rs!Fisico
'    t.LetraFisico = rs!LetraFisico
'
'    t.Mental = rs!Mental
'    t.LetraMental = rs!LetraMental
'
'    t.Emocional = rs!Emocional
'    t.LetraEmocional = rs!LetraEmocional
'
'    t.Espiritual = rs!Espiritual
'    t.LetraEspiritual = rs!LetraEspiritual
'
'    t.Esencia = rs!Esencia
'    t.AnioPersonal = rs!AnioPersonal
'
'    rs.Close
'    Set CargarTransito = t
'
'End Function
'
'Public Function CargarTodosLosTransitos(idResultado As Long) As Collection
'
'    Dim col As New Collection
'    Dim rs As DAO.Recordset
'    Dim SQL As String
'    Dim t As clsTransitos
'
'    SQL = "SELECT * FROM tbuTransitos WHERE IDResultado = " & idResultado & " ORDER BY Orden"
'    Set rs = CurrentDb.OpenRecordset(SQL, dbOpenDynaset)
'
'    Do While Not rs.EOF
'        Set t = New clsTransitos
'
'        t.idTransito = rs!idTransito
'        t.IDPersona = rs!IDPersona
'        t.idResultado = rs!idResultado
'        t.Orden = rs!Orden
'
'        t.anio = rs!anio
'        t.Edad = rs!Edad
'
'        t.Fisico = rs!Fisico
'        t.LetraFisico = rs!LetraFisico
'
'        t.Mental = rs!Mental
'        t.LetraMental = rs!LetraMental
'
'        t.Emocional = rs!Emocional
'        t.LetraEmocional = rs!LetraEmocional
'
'        t.Espiritual = rs!Espiritual
'        t.LetraEspiritual = rs!LetraEspiritual
'
'        t.Esencia = rs!Esencia
'        t.AnioPersonal = rs!AnioPersonal
'
'        col.Add t
'        rs.MoveNext
'    Loop
'
'    rs.Close
'    Set CargarTodosLosTransitos = col
'
'End Function
'
'' ============================================================
''   Guarda Datos Inclusion en la tabla
'' ============================================================
'
'Public Sub GuardarInclusion(i As clsInclusion)
'
'    Dim db As DAO.Database
'    Dim rs As DAO.Recordset
'    Dim SQL As String
'
'    Set db = CurrentDb
'
'    ' Intentamos localizar el registro existente
'    SQL = "SELECT * FROM tbuInclusion WHERE IDResultado = " & i.idResultado & _
'          " AND IDPersona = " & i.IDPersona
'
'    Set rs = db.OpenRecordset(SQL, dbOpenDynaset)
'
'    If rs.EOF Then
'        ' No existe ? crear nuevo
'        rs.AddNew
'        rs!idResultado = i.idResultado
'        rs!IDPersona = i.IDPersona
'        rs!IDFonetica = i.IDFonetica
'    Else
'        ' Existe ? editar
'        rs.Edit
'    End If
'
'    ' --- Volcado de datos ---
'    rs!N1 = i.N1
'    rs!N2 = i.N2
'    rs!N3 = i.N3
'    rs!N4 = i.N4
'    rs!N5 = i.N5
'    rs!N6 = i.N6
'    rs!N7 = i.N7
'    rs!N8 = i.N8
'    rs!N9 = i.N9
'
'    rs.Update
'
'    rs.Close
'    Set rs = Nothing
'    Set db = Nothing
'
'End Sub
'
'Public Function CargarInclusion(idResultado As Long, IDPersona As Long) As clsInclusion
'
'    Dim i As New clsInclusion
'    Dim rs As DAO.Recordset
'    Dim SQL As String
'
'    SQL = "SELECT * FROM tbuInclusion WHERE IDResultado = " & idResultado & _
'          " AND IDPersona = " & IDPersona
'
'    Set rs = CurrentDb.OpenRecordset(SQL, dbOpenDynaset)
'
'    If rs.EOF Then
'        Set CargarInclusion = Nothing
'        Exit Function
'    End If
'
'    i.idResultado = rs!idResultado
'    i.IDPersona = rs!IDPersona
'    i.IDFonetica = rs!IDFonetica
'
'    i.N1 = rs!N1
'    i.N2 = rs!N2
'    i.N3 = rs!N3
'    i.N4 = rs!N4
'    i.N5 = rs!N5
'    i.N6 = rs!N6
'    i.N7 = rs!N7
'    i.N8 = rs!N8
'    i.N9 = rs!N9
'
'    rs.Close
'    Set CargarInclusion = i
'
'End Function


' ============================================================
'   Guarda Datos Inclusion en la tabla
' ============================================================




'Public Function CargarTodosLosTransitos(idResultado As Long) As Collection
'
'    Dim col As New Collection
'    Dim rs As DAO.Recordset
'    Dim SQL As String
'    Dim t As clsTransitos
'
'    ' Seleccionamos todos los tránsitos del resultado, ordenados por Orden
'    SQL = "SELECT * FROM tbuTransitos " & _
'          "WHERE IDResultado = " & idResultado & " " & _
'          "ORDER BY Orden"
'
'    Set rs = CurrentDb.OpenRecordset(SQL, dbOpenDynaset)
'
'    Do While Not rs.EOF
'
'        Set t = New clsTransitos
'
'        ' Identificadores
'        t.idTransito = rs!idTransito
'        t.IDPersona = rs!IDPersona
'        t.idResultado = rs!idResultado
'        t.Orden = rs!Orden
'
'        ' Datos base
'        t.anio = rs!anio
'        t.edad = rs!edad
'
'        ' Valores numéricos
'        t.Fisico = rs!Fisico
'        t.Mental = rs!Mental
'        t.Emocional = rs!Emocional
'        t.Espiritual = rs!Espiritual
'
'        ' Letras asociadas
'        t.LetraFisico = rs!LetraFisico
'        t.LetraMental = rs!LetraMental
'        t.LetraEmocional = rs!LetraEmocional
'        t.LetraEspiritual = rs!LetraEspiritual
'
'        ' Otros valores
'        t.Esencia = rs!Esencia
'        t.AñoPersonal = rs!AñoPersonal
'
'        ' Añadir a la colección
'        col.Add t
'
'        rs.MoveNext
'    Loop
'
'    rs.Close
'    Set CargarTodosLosTransitos = col
'
'End Function


'===================================================================================================
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




'Public Function AutoNumber_1(ByVal NombreTabla As String, _
'                                    ByVal NombreCampo As String) As Long
'    Dim rs As DAO.Recordset
'    Dim SQL As String
''    Dim ultimo As Long
'    Dim esperado As Long
'    Dim SiguienteSecuencial As Long
'
'
'    SQL = "SELECT [" & NombreCampo & "] FROM [" & NombreTabla & "] " & _
'          "WHERE [" & NombreCampo & "] Is Not Null " & _
'          "ORDER BY [" & NombreCampo & "]"
'
'    Set rs = CurrentDb.OpenRecordset(SQL, dbOpenSnapshot)
'
'    esperado = 1
'
'    Do While Not rs.EOF
'        If rs.Fields(0).Value > esperado Then
'            ' Encontrado hueco
'            SiguienteSecuencial = esperado
'            rs.Close
'            Exit Function
'        End If
'
'        esperado = esperado + 1
'        rs.MoveNext
'    Loop
'
'    rs.Close
'
'    ' Si no hay huecos, el siguiente es el esperado
'    AutoNumber_1 = esperado
'End Function
'
'Public Function AutoNumber(ByVal Campo As String, _
'                               ByVal Tabla As String, _
'                               Optional ByVal MiWhere As String = "", _
'                               Optional ByVal DbPath As String = "") As Long
'    Dim db As DAO.Database
'    Dim rs As DAO.Recordset
'    Dim SQL As String
'    Dim FieldName As String, TableName As String
'    Dim SiguienteHueco As Long
'
'    FieldName = "[" & Campo & "]"
'    TableName = "[" & Tabla & "]"
'
'    On Error GoTo ErrHandler
'
'    ' Base actual o externa
'    If DbPath = "" Then
'        Set db = CurrentDb
'    Else
'        Set db = DBEngine.OpenDatabase(DbPath)
'    End If
'
'    ' 1) ¿Existe el 1?
'    SQL = "SELECT " & FieldName & " FROM " & TableName & _
'          " WHERE " & FieldName & " = 1"
'    If MiWhere <> "" Then SQL = SQL & " AND " & MiWhere
'
'    Set rs = db.OpenRecordset(SQL, dbOpenSnapshot)
'
'    If rs.EOF Then
'        AutoNumber = 1
'        GoTo Salir
'    End If
'
'    rs.Close
'
'    ' 2) Buscar el primer hueco
'    SQL = "SELECT MIN(" & FieldName & " + 1) " & _
'          "FROM " & TableName & " " & _
'          "WHERE NOT (" & FieldName & " + 1) IN " & _
'          " (SELECT " & FieldName & " FROM " & TableName
'
'    If MiWhere <> "" Then SQL = SQL & " WHERE " & MiWhere
'    SQL = SQL & ")"
'
'    If MiWhere <> "" Then SQL = SQL & " AND " & MiWhere
'
'    Set rs = db.OpenRecordset(SQL, dbOpenSnapshot)
'
'    If Not rs.EOF Then
'        SiguienteHueco = rs(0)
'    Else
'        AutoNumber = 1
'    End If
'
'Salir:
'    If Not rs Is Nothing Then rs.Close
'    Set rs = Nothing
'
'    ' Cerrar solo si es externa
'    If DbPath <> "" Then
'        db.Close
'    End If
'
'    Set db = Nothing
'    Exit Function
'
'ErrHandler:
'    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
'    Resume Salir
'End Function


'Public Sub CrearTabla_tbuTransitos()
'
'    Dim db As DAO.Database
'    Dim t As DAO.TableDef
'    Dim fld As DAO.Field
'    Dim idx As DAO.Index
'
'    Set db = CurrentDb
'
'    ' Si la tabla existe, la borramos (opcional)
'    On Error Resume Next
'    db.TableDefs.Delete "tbuTransitos"
'    On Error GoTo 0
'
'    ' Crear tabla
'    Set t = db.CreateTableDef("tbuTransitos")
'
'    ' ============================
'    '   CAMPOS
'    ' ============================
'
'    ' Clave primaria autonumérica
'    Set fld = t.CreateField("IDTransito", dbLong)
'    fld.Attributes = fld.Attributes Or dbAutoIncrField
'    t.Fields.Append fld
'
'    ' Identificadores
'    t.Fields.Append t.CreateField("IDPersona", dbLong)
'    t.Fields.Append t.CreateField("IDResultado", dbLong)
'    t.Fields.Append t.CreateField("Orden", dbInteger)
'
'    ' Datos base
'    t.Fields.Append t.CreateField("Anio", dbInteger)
'    t.Fields.Append t.CreateField("Edad", dbInteger)
'
'    ' Valores numéricos (texto porque pueden ser K, 11, 22…)
'    t.Fields.Append t.CreateField("Fisico", dbText, 10)
'    t.Fields.Append t.CreateField("Mental", dbText, 10)
'    t.Fields.Append t.CreateField("Emocional", dbText, 10)
'    t.Fields.Append t.CreateField("Espiritual", dbText, 10)
'
'    ' Letras asociadas
'    t.Fields.Append t.CreateField("LetraFisico", dbText, 5)
'    t.Fields.Append t.CreateField("LetraMental", dbText, 5)
'    t.Fields.Append t.CreateField("LetraEmocional", dbText, 5)
'    t.Fields.Append t.CreateField("LetraEspiritual", dbText, 5)
'
'    ' Otros valores
'    t.Fields.Append t.CreateField("Esencia", dbInteger)
'    t.Fields.Append t.CreateField("AñoPersonal", dbInteger)
'
'    ' Añadir tabla a la base de datos
'    db.TableDefs.Append t
'
'    ' ============================
'    '   ÍNDICES
'    ' ============================
'
'    ' Índice primario
'    Set idx = t.CreateIndex("PrimaryKey")
'    idx.Primary = True
'    idx.Unique = True
'    idx.Fields.Append idx.CreateField("IDTransito")
'    t.Indexes.Append idx
'
'    ' Índice para cargar rápido los tránsitos por resultado
'    Set idx = t.CreateIndex("idxResultadoOrden")
'    idx.Fields.Append idx.CreateField("IDResultado")
'    idx.Fields.Append idx.CreateField("Orden")
'    t.Indexes.Append idx
'
'    ' Índice por persona (opcional)
'    Set idx = t.CreateIndex("idxPersona")
'    idx.Fields.Append idx.CreateField("IDPersona")
'    t.Indexes.Append idx
'
'    MsgBox "Tabla tbuTransitos creada correctamente.", vbInformation
'
'End Sub

