Attribute VB_Name = "bas_Comunes_Motores"

Option Compare Database
Option Explicit


Public Type ConfigFon
    ModoSibilantes As Byte
    ModoLateral As Byte
    ModoH As Byte
    ModoX As Byte
    ModoLigadura As Byte
    'DigrafosActivos As Collection
End Type

Public CFG As ConfigFon
'Public PreferirIlLu As Boolean

Public ObjDTO As clsDTO_Motor

' Caché en memoria para acelerar búsquedas
Public IPA_Cache As Scripting.Dictionary


'Public prefijosCargados As Boolean

Public Function CreaSQL(Idioma As String) As String

    CreaSQL = _
        "SELECT ID, [" & Idioma & "]" & vbCrLf & _
        "FROM tbmDigrafosIdioma" & vbCrLf & _
        "WHERE Nz([" & Idioma & "],'') <> '';"

End Function


Sub CargaCacheIPA()

    Dim rs As DAO.Recordset

    Dim strSQL As String
    
    ' Inicializar caché si es la primera vez
    If IPA_Cache Is Nothing Then
        Set IPA_Cache = CreateObject("Scripting.Dictionary")
'    Else
'        Exit Sub
    End If

'    ' Si ya está en caché ? devolverlo directamente
'    If IPA_Cache.Exists(idFonema) Then
'        ObtenerIPA = IPA_Cache(idFonema)
'        Exit Function
'    End If

'    ' ID desconocido o especial
'    If idFonema = 255 Then
'        IPA_Cache.Add idFonema, ""   ' vacío
'        ObtenerIPA = ""
'        Exit Function
'    End If

    'strSQL = "SELECT ID, IPA FROM qryFonemasValor  WHERE ID=" & idFonema & ";"
    strSQL = "SELECT ID, First(IPA) AS PdIPA FROM qryFonemasValor GROUP BY ID ORDER BY ID, First(IPA);"


    ' Buscar en la tabla qryFonemasValor
    Set rs = CurrentDb.OpenRecordset(strSQL, dbOpenSnapshot)

    While Not rs.EOF ' And rs.BOF) Then
        If Not IPA_Cache.Exists(rs!id) Then
            IPA_Cache.Add rs!id, Nz(rs!PdIPA, "")
        End If
        rs.MoveNext
    Wend
'        ObtenerIPA = Nz(rs!ipa, "")
'    Else
'        ' Si no existe el ID ? devolver vacío
'        IPA_Cache.Add idFonema, ""
'        ObtenerIPA = ""
'    End If

    rs.Close
    Set rs = Nothing

End Sub


Function limpia(texto) As String

    limpia = Nz(texto, "")

    If Right(texto, 1) = "-" Then
        limpia = Trim$(Left(texto, Len(texto) - 1))
    End If
End Function
