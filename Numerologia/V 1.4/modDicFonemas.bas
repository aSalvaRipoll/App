Attribute VB_Name = "modDicFonemas"
' ------------------------------------------------------
' Name: modDicFonemas
' Date: 30/12/2025
' ------------------------------------------------------

Option Compare Database
Option Explicit

'Public DicNombres As Scripting.Dictionary
'Public DicApellidos As Scripting.Dictionary
'
'Public ColNombres As Collection
'Public ColApellidos As Collection


Public Sub InicializarDiccionarios()
    Set DicNombres = New Scripting.Dictionary
    Set DicApellidos = New Scripting.Dictionary

    Set ColNombres = New Collection
    Set ColApellidos = New Collection
End Sub

Public Sub CargarDiccionarioFonemas(ByVal idioma As String)

    Dim rs As DAO.Recordset
    'Set rs = New Recordset

    Set DicFonemas = New Scripting.Dictionary
    
    Set rs = CurrentDb.OpenRecordset("SELECT Palabra, FonemaCompleto FROM tbmDicFonemas WHERE Idioma = '" & idioma & "' AND Activo = True")
    'rs.Open "SELECT Palabra, FonemaCompleto FROM tbmDicFonemas WHERE Idioma = '" & Idioma & "' AND Activo = True", cn, adOpenForwardOnly, adLockReadOnly

    Do While Not rs.EOF
        If Not DicFonemas.Exists(UCase(rs!palabra)) Then
            DicFonemas.Add UCase(rs!palabra), rs!FonemaCompleto
        End If
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing

End Sub



'Public Sub CargarDiccionarioIdioma(ByVal idioma As String)
'
'    Dim entrada As clsEntradaDiccionario
'    Dim Texto As String, fonema As String, tipo As String
'
'    ' Aquí llamas a tus funciones AgregarEntradaDiccionario
'    ' que ya tienes generadas para EU, CA, GA, ES
'
'    ' Ejemplo conceptual:
'    ' For Each registro In DiccionarioIdioma(idioma)
'    '     Set entrada = New clsEntradaDiccionario
'    '     entrada.Texto = registro.Texto
'    '     entrada.Fonema = registro.Fonema
'    '     entrada.Idioma = idioma
'    '     entrada.Tipo = registro.Tipo
'    '     entrada.Fuente = idioma
'    '
'    '     Diccionario(entrada.Texto) = entrada.Fonema
'    '     ColeccionEntradas.Add entrada
'    ' Next
'
'End Sub

' ====================================================================================

Public Sub AgregarEntradaDiccionario( _
        ByVal idioma As String, _
        ByVal palabra As String, _
        ByVal FonemaCompleto As String, _
        Optional ByVal TipoEntrada As String = "", _
        Optional ByVal notas As String = "", _
        Optional ByVal Origen As String = "", _
        Optional ByVal FonemaIPA As String = "")

    Dim sql As String
    FonemaIPA = ""
    
    palabra = Replace(palabra, "'", "''")
    FonemaCompleto = Replace(FonemaCompleto, "'", "''")
    notas = Replace(notas, "'", "''")
    Origen = Replace(Origen, "'", "''")
    
    If TipoEntrada = "NOMBRE" Then
        sql = "INSERT INTO tbmDicFonemasNom (Idioma, Palabra, FonemaCompleto, FonemaIPA, TipoEntrada, Notas, Origen, Activo) " & _
            "VALUES ('" & UCase(idioma) & "', '" & UCase(palabra) & "', '" & FonemaCompleto & "', '" & FonemaIPA & "', '" & TipoEntrada & "', '" & notas & "', '" & Origen & "', True)"
    
    ElseIf TipoEntrada = "APELLIDO" Then
        sql = "INSERT INTO tbmDicFonemasApe (Idioma, Palabra, FonemaCompleto, FonemaIPA, TipoEntrada, Notas, Origen, Activo) " & _
              "VALUES ('" & UCase(idioma) & "', '" & UCase(palabra) & "', '" & FonemaCompleto & "', '" & FonemaIPA & "', '" & TipoEntrada & "', '" & notas & "', '" & Origen & "', True)"
    
    Else
        
        Stop

    End If
    CurrentDb.Execute sql

End Sub




'Public Sub AgregarEntradaDiccionario( _
'        ByVal idioma As String, _
'        ByVal palabra As String, _
'        ByVal FonemaCompleto As String, _
'        Optional ByVal TipoEntrada As String = "NOMBRE", _
'        Optional ByVal notas As String = "", _
'        Optional ByVal FonemaIPA As String = "")
'
'    Dim sql As String
'    FonemaIPA = ""
'
'    'sql = "INSERT INTO tbmDicFonemasNom (Idioma, Palabra, FonemaCompleto, FonemaIPA, TipoEntrada, Notas, Activo) " & _
'          "VALUES ('" & UCase(idioma) & "', '" & UCase(Palabra) & "', '" & FonemaCompleto & "', '" & FonemaIPA & "', '" & TipoEntrada & "', '" & Notas & "', True)"
'
'    sql = "INSERT INTO tbmDicFonemasApe (Idioma, Palabra, FonemaCompleto, FonemaIPA, TipoEntrada, Notas, Activo) " & _
'          "VALUES ('" & UCase(idioma) & "', '" & UCase(palabra) & "', '" & FonemaCompleto & "', '" & FonemaIPA & "', '" & TipoEntrada & "', '" & notas & "', True)"
'
'    'cn.Execute sql
'    CurrentDb.Execute sql
'
'End Sub

'Public Sub AgregarEntradaDiccionario( _
'        ByVal Idioma As String, _
'        ByVal Palabra As String, _
'        ByVal FonemaCompleto As String, _
'        Optional ByVal TipoEntrada As String = "PALABRA", _
'        Optional ByVal Notas As String = "")
'
'    Dim sql As String
'
'    sql = "INSERT INTO tbmDicFonemas (Idioma, Palabra, FonemaCompleto, TipoEntrada, Notas, Activo) " & _
'          "VALUES ('" & UCase(Idioma) & "', '" & UCase(Palabra) & "', '" & FonemaCompleto & "', '" & TipoEntrada & "', '" & Notas & "', True)"
'
'    cn.Execute sql
'
'End Sub


'Public Sub PoblarDiccionarioBase()
'
'    ' GALLEGO
'    AgregarEntradaDiccionario "GL", "SANXENXO", "SANSHENSHO", "PALABRA", "Patrimonial"
'    AgregarEntradaDiccionario "GL", "XOAN", "JOAN", "PALABRA", "Patrimonial"
'    AgregarEntradaDiccionario "GL", "XURXO", "JURJO", "PALABRA", "Patrimonial"
'    AgregarEntradaDiccionario "GL", "XAVIER", "SHAVIER", "PALABRA", "Medieval"
'    AgregarEntradaDiccionario "GL", "XENON", "KSENON", "PALABRA", "Cultismo"
'
'    ' CATALÁN
'    AgregarEntradaDiccionario "CA", "XIRGU", "SHIRGU", "PALABRA", "Patrimonial"
'    AgregarEntradaDiccionario "CA", "XARXA", "SHARSHA", "PALABRA", "Patrimonial"
'    AgregarEntradaDiccionario "CA", "XILOFON", "KSILOFON", "PALABRA", "Cultismo"
'
'    ' CASTELLANO
'    AgregarEntradaDiccionario "ES", "MEXICO", "MEJICO", "PALABRA", "Excepción histórica"
'
'    ' EUSKERA
'    AgregarEntradaDiccionario "EU", "XABIER", "SHABIER", "PALABRA", "Fonema propio"
'
'    MsgBox "Diccionario base cargado correctamente.", vbInformation
'
'End Sub














