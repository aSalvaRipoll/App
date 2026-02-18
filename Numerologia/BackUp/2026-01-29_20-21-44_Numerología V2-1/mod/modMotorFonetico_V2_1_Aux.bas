Attribute VB_Name = "modMotorFonetico_V2_1_Aux"
Option Compare Database
Option Explicit


'Public Function Levenshtein(ByVal s1 As String, ByVal s2 As String) As Long
'    Dim len1 As Long, len2 As Long
'    Dim i As Long, j As Long
'    Dim cost As Long
'    Dim v0() As Long, v1() As Long
'    Dim temp() As Long
'
'    ' Normalización opcional para apellidos
'    s1 = Trim$(LCase$(s1))
'    s2 = Trim$(LCase$(s2))
'
'    len1 = Len(s1)
'    len2 = Len(s2)
'
'    ' Casos rápidos
'    If len1 = 0 Then
'        Levenshtein = len2
'        Exit Function
'    End If
'
'    If len2 = 0 Then
'        Levenshtein = len1
'        Exit Function
'    End If
'
'    ' Redimensionar matrices
'    ReDim v0(0 To len2)
'    ReDim v1(0 To len2)
'
'    ' Inicializar primera fila
'    For i = 0 To len2
'        v0(i) = i
'    Next i
'
'    ' Bucle principal
'    For i = 1 To len1
'        v1(0) = i
'
'        For j = 1 To len2
'            If Mid$(s1, i, 1) = Mid$(s2, j, 1) Then
'                cost = 0
'            Else
'                cost = 1
'            End If
'
'            v1(j) = Application.WorksheetFunction.Min( _
'                        v1(j - 1) + 1, _
'                        v0(j) + 1, _
'                        v0(j - 1) + cost)
'        Next j
'
'        ' Intercambiar filas
'        temp = v0
'        v0 = v1
'        v1 = temp
'    Next i
'
'    Levenshtein = v0(len2)
'End Function



Public Function MF_NormalizarVocalesPorIdioma( _
    ByVal Texto As String, _
    ByVal idioma As String _
    ) As String

    Select Case UCase$(idioma)

        Case "ES"
            Texto = MF_NormalizarVocales_ES(Texto)

        Case "CA" ', "CA-IB", "CA-VA"
            Texto = MF_NormalizarVocales_CA(Texto)
            
        Case "CA-IB"
            Texto = MF_NormalizarVocales_CA_IB(Texto)

        Case "CA-VA"
            Texto = MF_NormalizarVocales_CA_VA(Texto)
            
        Case "GL"
            Texto = MF_NormalizarVocales_GL(Texto)

        Case "EU"
            Texto = MF_NormalizarVocales_EU(Texto)

        Case "PT-EU"
            Texto = MF_NormalizarVocales_PT_EU(Texto)
        
        Case "PT-BR"
            Texto = MF_NormalizarVocales_PT_BR(Texto)

        Case "FR"
            Texto = MF_NormalizarVocales_FR(Texto)

        Case "EN-GB"
            Texto = MF_NormalizarVocales_EN_GB(Texto)

'        Case "EN-US"
'            texto = MF_NormalizarVocales_EN_US(texto)

'        Case "EN-US-AF"
'            texto = MF_NormalizarVocales_EN_US_AF(texto)

        Case Else
            Texto = MF_NormalizarVocales_General(Texto)

    End Select

    MF_NormalizarVocalesPorIdioma = Texto

End Function

Private Function MF_NormalizarVocales_ES(ByVal Texto As String) As String

    ' A
    Texto = Replace(Texto, "Á", "A")
    Texto = Replace(Texto, "À", "A")
    Texto = Replace(Texto, "Ä", "A")
    Texto = Replace(Texto, "Â", "A")

    ' E
    Texto = Replace(Texto, "É", "E")
    Texto = Replace(Texto, "È", "E")
    Texto = Replace(Texto, "Ë", "E")
    Texto = Replace(Texto, "Ê", "E")

    ' I
    Texto = Replace(Texto, "Í", "I")
    Texto = Replace(Texto, "Ì", "I")
    Texto = Replace(Texto, "Ï", "I")
    Texto = Replace(Texto, "Î", "I")

    ' O
    Texto = Replace(Texto, "Ó", "O")
    Texto = Replace(Texto, "Ò", "O")
    Texto = Replace(Texto, "Ö", "O")
    Texto = Replace(Texto, "Ô", "O")

    ' U (sin tocar Ü)
    Texto = Replace(Texto, "Ú", "U")
    Texto = Replace(Texto, "Ù", "U")
    Texto = Replace(Texto, "Û", "U")

    MF_NormalizarVocales_ES = Texto

End Function

Private Function MF_NormalizarVocales_CA(ByVal Texto As String) As String

    ' A
    Texto = Replace(Texto, "À", "A")
    Texto = Replace(Texto, "Á", "A")

    ' E
    Texto = Replace(Texto, "È", "E")
    Texto = Replace(Texto, "É", "E")

    ' I  (NO tocar Ï)
    Texto = Replace(Texto, "Í", "I")

    ' O
    Texto = Replace(Texto, "Ò", "O")
    Texto = Replace(Texto, "Ó", "O")

    ' U  (NO tocar Ü)
    Texto = Replace(Texto, "Ú", "U")

    MF_NormalizarVocales_CA = Texto

End Function

Private Function MF_NormalizarVocales_CA_IB(ByVal Texto As String) As String
    MF_NormalizarVocales_CA_IB = MF_NormalizarVocales_CA(Texto)
End Function

Private Function MF_NormalizarVocales_CA_VA(ByVal Texto As String) As String
    MF_NormalizarVocales_CA_VA = MF_NormalizarVocales_CA(Texto)
End Function

Private Function MF_NormalizarVocales_GL(ByVal Texto As String) As String

    ' A
    Texto = Replace(Texto, "Á", "A")
    Texto = Replace(Texto, "À", "A")

    ' E
    Texto = Replace(Texto, "É", "E")
    Texto = Replace(Texto, "È", "E")

    ' I
    Texto = Replace(Texto, "Í", "I")
    Texto = Replace(Texto, "Ì", "I")

    ' O
    Texto = Replace(Texto, "Ó", "O")
    Texto = Replace(Texto, "Ò", "O")

    ' U
    Texto = Replace(Texto, "Ú", "U")
    Texto = Replace(Texto, "Ù", "U")

    MF_NormalizarVocales_GL = Texto

End Function


Private Function MF_NormalizarVocales_EU(ByVal Texto As String) As String

    ' A
    Texto = Replace(Texto, "Á", "A")
    Texto = Replace(Texto, "À", "A")

    ' E
    Texto = Replace(Texto, "É", "E")
    Texto = Replace(Texto, "È", "E")

    ' I
    Texto = Replace(Texto, "Í", "I")
    Texto = Replace(Texto, "Ì", "I")

    ' O
    Texto = Replace(Texto, "Ó", "O")
    Texto = Replace(Texto, "Ò", "O")

    ' U
    Texto = Replace(Texto, "Ú", "U")
    Texto = Replace(Texto, "Ù", "U")

    MF_NormalizarVocales_EU = Texto

End Function

Private Function MF_NormalizarVocales_PT_EU(ByVal Texto As String) As String

    ' Nasales
    Texto = Replace(Texto, "Ã", "A~")
    Texto = Replace(Texto, "Õ", "O~")

    ' Cerradas (circunflejo)
    Texto = Replace(Texto, "Â", "Â")
    Texto = Replace(Texto, "Ê", "Ê")
    Texto = Replace(Texto, "Î", "I") ' no existe en PT, pero por robustez
    Texto = Replace(Texto, "Ô", "Ô")
    Texto = Replace(Texto, "Û", "U") ' no existe en PT, robustez

    ' Abiertas (agudas)
    Texto = Replace(Texto, "Á", "A´")
    Texto = Replace(Texto, "É", "E´")
    Texto = Replace(Texto, "Í", "I´")
    Texto = Replace(Texto, "Ó", "O´")
    Texto = Replace(Texto, "Ú", "U´")

    ' Graves (no existen en PT, pero pueden aparecer en nombres importados)
    Texto = Replace(Texto, "À", "A")
    Texto = Replace(Texto, "È", "E")
    Texto = Replace(Texto, "Ì", "I")
    Texto = Replace(Texto, "Ò", "O")
    Texto = Replace(Texto, "Ù", "U")

    MF_NormalizarVocales_PT_EU = Texto

End Function

Private Function MF_NormalizarVocales_PT_BR(ByVal Texto As String) As String

    ' Nasales (idénticas a PT-EU)
    Texto = Replace(Texto, "Ã", "A~")
    Texto = Replace(Texto, "Õ", "O~")

    ' Cerradas ? se suavizan (PT-BR no las mantiene tan tensas)
    Texto = Replace(Texto, "Â", "A")
    Texto = Replace(Texto, "Ê", "E")
    Texto = Replace(Texto, "Ô", "O")

    ' Abiertas (agudas)
    Texto = Replace(Texto, "Á", "A´")
    Texto = Replace(Texto, "É", "E´")
    Texto = Replace(Texto, "Í", "I´")
    Texto = Replace(Texto, "Ó", "O´")
    Texto = Replace(Texto, "Ú", "U´")

    ' Graves (robustez)
    Texto = Replace(Texto, "À", "A")
    Texto = Replace(Texto, "È", "E")
    Texto = Replace(Texto, "Ì", "I")
    Texto = Replace(Texto, "Ò", "O")
    Texto = Replace(Texto, "Ù", "U")

    MF_NormalizarVocales_PT_BR = Texto

End Function

Private Function MF_NormalizarVocales_FR(ByVal Texto As String) As String

    ' A
    Texto = Replace(Texto, "À", "A")   ' abierta
    Texto = Replace(Texto, "Á", "A")   ' rara, pero robustez
    Texto = Replace(Texto, "Â", "Â")   ' cerrada
    Texto = Replace(Texto, "Ä", "A¨")  ' hiato

    ' E
    Texto = Replace(Texto, "È", "E")   ' abierta
    Texto = Replace(Texto, "É", "E´")  ' cerrada
    Texto = Replace(Texto, "Ê", "Ê")   ' cerrada tensa
    Texto = Replace(Texto, "Ë", "E¨")  ' hiato

    ' I
    Texto = Replace(Texto, "Ì", "I")   ' robustez
    Texto = Replace(Texto, "Í", "I")   ' robustez
    Texto = Replace(Texto, "Î", "Î")   ' cerrada
    Texto = Replace(Texto, "Ï", "I¨")  ' hiato

    ' O
    Texto = Replace(Texto, "Ò", "O")   ' robustez
    Texto = Replace(Texto, "Ó", "O")   ' robustez
    Texto = Replace(Texto, "Ô", "Ô")   ' cerrada
    Texto = Replace(Texto, "Ö", "O¨")  ' hiato

    ' U
    Texto = Replace(Texto, "Ù", "U")   ' abierta
    Texto = Replace(Texto, "Ú", "U")   ' robustez
    Texto = Replace(Texto, "Û", "Û")   ' cerrada
    Texto = Replace(Texto, "Ü", "U¨")  ' hiato

    MF_NormalizarVocales_FR = Texto

End Function

Private Function MF_NormalizarVocales_EN_GB(ByVal Texto As String) As String

    ' Solo por robustez ante nombres importados
    Texto = Replace(Texto, "Á", "A")
    Texto = Replace(Texto, "À", "A")
    Texto = Replace(Texto, "Ä", "A")
    Texto = Replace(Texto, "Â", "A")

    Texto = Replace(Texto, "É", "E")
    Texto = Replace(Texto, "È", "E")
    Texto = Replace(Texto, "Ë", "E")
    Texto = Replace(Texto, "Ê", "E")

    Texto = Replace(Texto, "Í", "I")
    Texto = Replace(Texto, "Ì", "I")
    Texto = Replace(Texto, "Ï", "I")
    Texto = Replace(Texto, "Î", "I")

    Texto = Replace(Texto, "Ó", "O")
    Texto = Replace(Texto, "Ò", "O")
    Texto = Replace(Texto, "Ö", "O")
    Texto = Replace(Texto, "Ô", "O")

    Texto = Replace(Texto, "Ú", "U")
    Texto = Replace(Texto, "Ù", "U")
    Texto = Replace(Texto, "Ü", "U")
    Texto = Replace(Texto, "Û", "U")

    MF_NormalizarVocales_EN_GB = Texto

End Function

'Private Function MF_NormalizarVocales_EN_US(ByVal texto As String) As String
'
'    ' Solo por robustez ante nombres importados
'    texto = Replace(texto, "Á", "A")
'    texto = Replace(texto, "À", "A")
'    texto = Replace(texto, "Ä", "A")
'    texto = Replace(texto, "Â", "A")
'
'    texto = Replace(texto, "É", "E")
'    texto = Replace(texto, "È", "E")
'    texto = Replace(texto, "Ë", "E")
'    texto = Replace(texto, "Ê", "E")
'
'    texto = Replace(texto, "Í", "I")
'    texto = Replace(texto, "Ì", "I")
'    texto = Replace(texto, "Ï", "I")
'    texto = Replace(texto, "Î", "I")
'
'    texto = Replace(texto, "Ó", "O")
'    texto = Replace(texto, "Ò", "O")
'    texto = Replace(texto, "Ö", "O")
'    texto = Replace(texto, "Ô", "O")
'
'    texto = Replace(texto, "Ú", "U")
'    texto = Replace(texto, "Ù", "U")
'    texto = Replace(texto, "Ü", "U")
'    texto = Replace(texto, "Û", "U")
'
'    MF_NormalizarVocales_EN_US = texto
'
'End Function
'
'Private Function MF_NormalizarVocales_EN_US_AF(ByVal texto As String) As String
'
'    ' Solo por robustez ante nombres importados
'    texto = Replace(texto, "Á", "A")
'    texto = Replace(texto, "À", "A")
'    texto = Replace(texto, "Ä", "A")
'    texto = Replace(texto, "Â", "A")
'
'    texto = Replace(texto, "É", "E")
'    texto = Replace(texto, "È", "E")
'    texto = Replace(texto, "Ë", "E")
'    texto = Replace(texto, "Ê", "E")
'
'    texto = Replace(texto, "Í", "I")
'    texto = Replace(texto, "Ì", "I")
'    texto = Replace(texto, "Ï", "I")
'    texto = Replace(texto, "Î", "I")
'
'    texto = Replace(texto, "Ó", "O")
'    texto = Replace(texto, "Ò", "O")
'    texto = Replace(texto, "Ö", "O")
'    texto = Replace(texto, "Ô", "O")
'
'    texto = Replace(texto, "Ú", "U")
'    texto = Replace(texto, "Ù", "U")
'    texto = Replace(texto, "Ü", "U")
'    texto = Replace(texto, "Û", "U")
'
'    MF_NormalizarVocales_EN_US_AF = texto
'
'End Function

Private Function MF_NormalizarVocales_General(ByVal Texto As String) As String
    
    Dim i As Integer
    Dim c As String
    Dim Res As String

    c = ""
    Res = ""
    
    For i = 1 To Len(Texto)
        c = Mid(Texto, i, 1)
        Select Case c
            
            Case "Á", "À", "Ä"
                c = "A"
    
            Case "É", "È", "Ë"
                c = "E"
    
            Case "Í", "Ï"
                c = "I"
            
            Case "Ó", "Ò", "Ö"
                c = "O"
    
            Case "Ú", "Ü"
                c = "U"
        End Select
        
        Res = Res & c
    Next i
    
    MF_NormalizarVocales_General = Res

End Function

Public Function EsVocal(ByVal c As String) As Boolean

'Versión blindada UNICODE

    Dim code As Long
    code = AscW(c)

    ' Vocales básicas A E I O U
    If code = &H41 Or code = &H45 Or code = &H49 Or code = &H4F Or code = &H55 Then
        EsVocal = True: Exit Function
    End If

    ' Vocales minúsculas (por si acaso)
    If code = &H61 Or code = &H65 Or code = &H69 Or code = &H6F Or code = &H75 Then
        EsVocal = True: Exit Function
    End If

    ' Vocales acentuadas (agudas, graves, circunflejos, diéresis)
    ' Rango general: U+00C0 – U+00FF (letras latinas extendidas)
    If code >= &HC0 And code <= &HFF Then
        Select Case code
            ' Á É Í Ó Ú
            Case &HC1, &HC9, &HCD, &HD3, &HDA
                EsVocal = True: Exit Function

            ' À È Ì Ò Ù
            Case &HC0, &HC8, &HCC, &HD2, &HD9
                EsVocal = True: Exit Function

            ' Â Ê Î Ô Û
            Case &HC2, &HCA, &HCE, &HD4, &HDB
                EsVocal = True: Exit Function

            ' Ä Ë Ï Ö Ü
            Case &HC4, &HCB, &HCF, &HD6, &HDC
                EsVocal = True: Exit Function

            ' Nasales portuguesas: Ã Õ
            Case &HC3, &HD5
                EsVocal = True: Exit Function
        End Select
    End If

    ' Si no coincide con nada
    EsVocal = False
    
End Function

'Public Function EsVocal(ByVal C As String) As Boolean
'
'' Versión no blindada
'
'    Select Case UCase$(C)
'
'        ' Vocales simples
'        Case "A", "E", "I", "O", "U"
'
'        ' Agudas
'        Case "Á", "É", "Í", "Ó", "Ú"
'
'        ' Graves
'        Case "À", "È", "Ì", "Ò", "Ù"
'
'        ' Circunflejos
'        Case "Â", "Ê", "Î", "Ô", "Û"
'
'        ' Diéresis
'        Case "Ä", "Ë", "Ï", "Ö", "Ü"
'
'        ' Nasales portuguesas
'        Case "Ã", "Õ"
'
'            EsVocal = True
'
'        Case Else
'            EsVocal = False
'
'    End Select
'
'End Function

Public Function EsConsonante(ByVal c As String) As Boolean
    If c = "" Then
        EsConsonante = False
    Else
        EsConsonante = Not EsVocal(c)
    End If
End Function

Public Function ProcesarY(ByVal ant As String, ByVal sig As String) As String
    If (ant = " ") And (sig = " ") Then
        ProcesarY = "I": Exit Function
    End If

    If ant = "" And EsVocal(sig) Then
        ProcesarY = "Y": Exit Function
    End If

    If ant = "" And EsConsonante(sig) Then
        ProcesarY = "I": Exit Function
    End If

    If sig = "" Then
        ProcesarY = "I": Exit Function
    End If

    If EsVocal(ant) And EsVocal(sig) Then
        ProcesarY = "Y": Exit Function
    End If

    If EsConsonante(ant) And EsConsonante(sig) Then
        ProcesarY = "I": Exit Function
    End If

    If EsVocal(sig) Then
        ProcesarY = "Y": Exit Function
    End If

    If EsConsonante(sig) Then
        ProcesarY = "I": Exit Function
    End If

    ProcesarY = "Y"
End Function

Public Function ProcesarW() As String
    ProcesarW = "GÜ"
End Function

Public Sub CargarFonemas()
    Dim rs As DAO.Recordset
    Dim f As clsFonema

    If Not colFonemas Is Nothing Then Exit Sub   ' Ya cargada

    Set colFonemas = New Collection
    'Set rs = CurrentDb.OpenRecordset("SELECT * FROM tbmFonemas ORDER BY idFonema")
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tbmFoneticaCompleta ORDER BY idFonema")

    Do While Not rs.EOF
        Set f = New clsFonema
        f.idFonema = rs!idFonema
        f.GrafemaOri = rs!fonema
        f.EsVocal = rs!EsVocal
        'f.Valor = rs!Valor
        ' ... cualquier otra propiedad

        colFonemas.Add f, CStr(f.idFonema)
        rs.MoveNext
    Loop

    rs.Close
End Sub

'Public Function BuscarExcepcion(ByVal palabra As String, ByVal idioma As String) As String
'
'    Dim rs As DAO.Recordset
'    Set rs = CurrentDb.OpenRecordset( _
'        "SELECT FonemaCompleto FROM tbmDicExcepciones " & _
'        "WHERE Idioma = '" & idioma & "' AND Palabra = '" & UCase(palabra) & "' AND Activo = True")
'
'    If Not rs.EOF Then
'        BuscarExcepcion = rs!FonemaCompleto
'    Else
'        BuscarExcepcion = ""
'    End If
'
'    rs.Close
'    Set rs = Nothing
'
'End Function

Public Function BuscarExcepcion(ByVal graf As String, ByVal idioma As String) As Byte

    Dim rs As DAO.Recordset
    Dim sql As String

    sql = "SELECT idFonema FROM tbmDicExcepciones " & _
          "WHERE Tipo = 'GRAFEMA' " & _
          "AND Idioma = '" & idioma & "' " & _
          "AND Grafema = '" & graf & "' " & _
          "AND Activo = True"

    Set rs = CurrentDb.OpenRecordset(sql)

    If Not rs.EOF Then
        BuscarExcepcion = rs!idFonema
    Else
        BuscarExcepcion = 0
    End If

    rs.Close
    Set rs = Nothing

End Function

Public Function BuscarExcepcionPalabra(ByVal palabra As String, ByVal idioma As String) As String

    Dim rs As DAO.Recordset
    Dim sql As String

    sql = "SELECT FonemaCompleto FROM tbmDicExcepciones " & _
          "WHERE Tipo = 'PALABRA' " & _
          "AND Idioma = '" & idioma & "' " & _
          "AND Palabra = '" & UCase(palabra) & "' " & _
          "AND Activo = True"

    Set rs = CurrentDb.OpenRecordset(sql)

    If Not rs.EOF Then
        BuscarExcepcionPalabra = rs!FonemaCompleto
    Else
        BuscarExcepcionPalabra = ""
    End If

    rs.Close
    Set rs = Nothing

End Function


