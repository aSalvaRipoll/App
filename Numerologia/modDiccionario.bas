Option Compare Database
Option Explicit

' ============================================================
'   DICCIONARIOS GLOBALES
'   (Lookup rápido: PALABRA → IDEntrada)
' ============================================================
Public DicNombres As Scripting.Dictionary
Public DicApellidos As Scripting.Dictionary

' ============================================================
'   COLECCIONES GLOBALES
'   (Objetos completos: IDEntrada → clsEntradaDiccionario)
' ============================================================
Public ColNombres As Scripting.Dictionary
Public ColApellidos As Scripting.Dictionary

' ============================================================
'   INICIALIZACIÓN
' ============================================================
Public Sub InicializarDiccionarios()

    ' Diccionarios por idioma
    Set DicNombres = New Scripting.Dictionary
    Set DicApellidos = New Scripting.Dictionary

    ' Colecciones por idioma
    Public ColNombres As Collection 
	Public ColApellidos As Collection

End Sub


Private Sub CargarDiccionarioIdioma( _
        ByVal idioma As String, _
        ByVal esNombres As Boolean)

    Dim rs As DAO.Recordset
    Dim entrada As clsEntradaDiccionario
    Dim sql As String

    ' Seleccionar tabla según tipo
    If esNombres Then
        sql = "SELECT * FROM tbmDicNombres WHERE Idioma='" & idioma & "' AND Activo=True"
    Else
        sql = "SELECT * FROM tbmDicApellidos WHERE Idioma='" & idioma & "' AND Activo=True"
    End If

    Set rs = CurrentDb.OpenRecordset(sql)

    Do While Not rs.EOF

        Set entrada = New clsEntradaDiccionario

        entrada.ID = rs!ID
        entrada.Texto = UCase(rs!Palabra)
        entrada.Fonema = rs!FonemaCompleto
        entrada.Idioma = idioma
        entrada.Tipo = rs!TipoEntrada
        entrada.Notas = Nz(rs!Notas, "")

        ' Lookup rápido: PALABRA → ID
        If esNombres Then
            If Not DicNombres.Exists(entrada.Texto) Then
                DicNombres.Add entrada.Texto, entrada.ID
            End If
        Else
            If Not DicApellidos.Exists(entrada.Texto) Then
                DicApellidos.Add entrada.Texto, entrada.ID
            End If
        End If

        ' Colección completa: ID → objeto
        If esNombres Then
            ColNombres.Add entrada, CStr(entrada.ID)
        Else
            ColApellidos.Add entrada, CStr(entrada.ID)
        End If

        rs.MoveNext
    Loop

    rs.Close

End Sub




Public Sub CargarTodosLosDiccionarios(idiomaPreferido As String)

    Dim idiomas As Variant
    Dim i As Long

    idiomas = Array("ES", "CA", "EU", "GL")

    ' 1. Cargar primero el idioma preferido
    CargarDiccionarioIdioma idiomaPreferido, True
    CargarDiccionarioIdioma idiomaPreferido, False

    ' 2. Cargar el resto
    For i = LBound(idiomas) To UBound(idiomas)
        If idiomas(i) <> idiomaPreferido Then
            CargarDiccionarioIdioma idiomas(i), True
            CargarDiccionarioIdioma idiomas(i), False
        End If
    Next i

End Sub
