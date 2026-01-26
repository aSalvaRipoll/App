Attribute VB_Name = "modColeccionesFonemas"

Option Compare Database
Option Explicit

Public Function GenerarColeccionFonemas( _
        ByVal NombreOriginal As String, _
        ByVal TranscripcionASCII As String, _
        ByVal idioma As String _
    ) As Collection

    Dim col As New Collection
    Dim i As Long
    Dim graf As String
    Dim f As clsFonema
    Dim rs As dao.Recordset
    Dim sql As String
    Dim NumOrden As Long
    
    NumOrden = 1
    i = 1
    
    Do While i <= Len(TranscripcionASCII)
        
        ' Intentar trigrafema (3 chars)
        If i <= Len(TranscripcionASCII) - 2 Then
            graf = Mid(TranscripcionASCII, i, 3)
            If ExisteFonema(graf, idioma) Then
                Set f = CrearDTO(NombreOriginal, graf, idioma, NumOrden)
                col.Add f
                i = i + 3
                NumOrden = NumOrden + 1
                GoTo Siguiente
            End If
        End If
        
        ' Intentar dígrafo (2 chars)
        If i <= Len(TranscripcionASCII) - 1 Then
            graf = Mid(TranscripcionASCII, i, 2)
            If ExisteFonema(graf, idioma) Then
                Set f = CrearDTO(NombreOriginal, graf, idioma, NumOrden)
                col.Add f
                i = i + 2
                NumOrden = NumOrden + 1
                GoTo Siguiente
            End If
        End If
        
        ' Monógrafo (1 char)
        graf = Mid(TranscripcionASCII, i, 1)
        If ExisteFonema(graf, idioma) Then
            Set f = CrearDTO(NombreOriginal, graf, idioma, NumOrden)
            col.Add f
            i = i + 1
            NumOrden = NumOrden + 1
        Else
            ' Si no existe, avanzar para evitar bucles
            i = i + 1
        End If
        
Siguiente:
    Loop
    
    Set GenerarColeccionFonemas = col
End Function


Private Function ExisteFonema(ByVal graf As String, ByVal idioma As String) As Boolean
    Dim rs As dao.Recordset
    Dim sql As String
    
    sql = "SELECT idFonema FROM Fonemas " & _
          "WHERE FonemaASCII = '" & graf & "' " & _
          "AND Lengua LIKE '*" & idioma & "*';"
    
    Set rs = CurrentDb.OpenRecordset(sql)
    ExisteFonema = Not rs.EOF
    rs.Close
End Function

Private Function CrearDTO( _
        ByVal NombreOriginal As String, _
        ByVal graf As String, _
        ByVal idioma As String, _
        ByVal NumOrden As Long _
    ) As clsFonema
    
    Dim f As New clsFonema
    Dim rs As dao.Recordset
    Dim sql As String
    
    sql = "SELECT * FROM Fonemas " & _
          "WHERE FonemaASCII = '" & graf & "' " & _
          "AND Lengua LIKE '*" & idioma & "*';"
    
    Set rs = CurrentDb.OpenRecordset(sql)
    
    If Not rs.EOF Then
        f.NumOrden = NumOrden
        f.GrafemaOri = graf
        f.ASCII = graf
        f.idFonema = rs!idFonema
        f.Valor = rs!Valor
'        f.Tipo = rs!Tipo
    End If
    
    rs.Close
    Set CrearDTO = f
End Function



