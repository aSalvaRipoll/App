Attribute VB_Name = "modMotorFonetico_V2_1_Aux"
Option Compare Database
Option Explicit


Public Function Levenshtein(ByVal s1 As String, ByVal s2 As String) As Long
    Dim len1 As Long, len2 As Long
    Dim i As Long, j As Long
    Dim cost As Long
    Dim v0() As Long, v1() As Long
    Dim temp() As Long
    
    ' Normalización opcional para apellidos
    s1 = Trim$(LCase$(s1))
    s2 = Trim$(LCase$(s2))
    
    len1 = Len(s1)
    len2 = Len(s2)
    
    ' Casos rápidos
    If len1 = 0 Then
        Levenshtein = len2
        Exit Function
    End If
    
    If len2 = 0 Then
        Levenshtein = len1
        Exit Function
    End If
    
    ' Redimensionar matrices
    ReDim v0(0 To len2)
    ReDim v1(0 To len2)
    
    ' Inicializar primera fila
    For i = 0 To len2
        v0(i) = i
    Next i
    
    ' Bucle principal
    For i = 1 To len1
        v1(0) = i
        
        For j = 1 To len2
            If Mid$(s1, i, 1) = Mid$(s2, j, 1) Then
                cost = 0
            Else
                cost = 1
            End If
            
            v1(j) = Application.WorksheetFunction.Min( _
                        v1(j - 1) + 1, _
                        v0(j) + 1, _
                        v0(j - 1) + cost)
        Next j
        
        ' Intercambiar filas
        temp = v0
        v0 = v1
        v1 = temp
    Next i
    
    Levenshtein = v0(len2)
End Function



Public Function MF_NormalizarVocalesPorIdioma( _
    ByVal texto As String, _
    ByVal idioma As String _
    ) As String

    Select Case UCase$(idioma)

        Case "ES"
            texto = MF_NormalizarVocales_ES(texto)

        Case "CA" ', "CA-IB", "CA-VA"
            texto = MF_NormalizarVocales_CA(texto)
            
        Case "CA-IB"
            texto = MF_NormalizarVocales_CA_IB(texto)

        Case "CA-VA"
            texto = MF_NormalizarVocales_CA_VA(texto)
            
        Case "GL"
            texto = MF_NormalizarVocales_GL(texto)

        Case "EU"
            texto = MF_NormalizarVocales_EU(texto)

        Case "PT-EU"
            texto = MF_NormalizarVocales_PT_EU(texto)
        
        Case "PT-BR"
            texto = MF_NormalizarVocales_PT_BR(texto)

        Case "FR"
            texto = MF_NormalizarVocales_FR(texto)

        Case "EN-GB"
            texto = MF_NormalizarVocales_EN_GB(texto)

'        Case "EN-US"
'            texto = MF_NormalizarVocales_EN_US(texto)

'        Case "EN-US-AF"
'            texto = MF_NormalizarVocales_EN_US_AF(texto)

        Case Else
            texto = MF_NormalizarVocales_General(texto)

    End Select

    MF_NormalizarVocalesPorIdioma = texto

End Function

Private Function MF_NormalizarVocales_ES(ByVal texto As String) As String

    ' A
    texto = Replace(texto, "Á", "A")
    texto = Replace(texto, "À", "A")
    texto = Replace(texto, "Ä", "A")
    texto = Replace(texto, "Â", "A")

    ' E
    texto = Replace(texto, "É", "E")
    texto = Replace(texto, "È", "E")
    texto = Replace(texto, "Ë", "E")
    texto = Replace(texto, "Ê", "E")

    ' I
    texto = Replace(texto, "Í", "I")
    texto = Replace(texto, "Ì", "I")
    texto = Replace(texto, "Ï", "I")
    texto = Replace(texto, "Î", "I")

    ' O
    texto = Replace(texto, "Ó", "O")
    texto = Replace(texto, "Ò", "O")
    texto = Replace(texto, "Ö", "O")
    texto = Replace(texto, "Ô", "O")

    ' U (sin tocar Ü)
    texto = Replace(texto, "Ú", "U")
    texto = Replace(texto, "Ù", "U")
    texto = Replace(texto, "Û", "U")

    MF_NormalizarVocales_ES = texto

End Function

Private Function MF_NormalizarVocales_CA(ByVal texto As String) As String

    ' A
    texto = Replace(texto, "À", "A")
    texto = Replace(texto, "Á", "A")

    ' E
    texto = Replace(texto, "È", "E")
    texto = Replace(texto, "É", "E")

    ' I  (NO tocar Ï)
    texto = Replace(texto, "Í", "I")

    ' O
    texto = Replace(texto, "Ò", "O")
    texto = Replace(texto, "Ó", "O")

    ' U  (NO tocar Ü)
    texto = Replace(texto, "Ú", "U")

    MF_NormalizarVocales_CA = texto

End Function

Private Function MF_NormalizarVocales_CA_IB(ByVal texto As String) As String
    MF_NormalizarVocales_CA_IB = MF_NormalizarVocales_CA(texto)
End Function

Private Function MF_NormalizarVocales_CA_VA(ByVal texto As String) As String
    MF_NormalizarVocales_CA_VA = MF_NormalizarVocales_CA(texto)
End Function

Private Function MF_NormalizarVocales_GL(ByVal texto As String) As String

    ' A
    texto = Replace(texto, "Á", "A")
    texto = Replace(texto, "À", "A")

    ' E
    texto = Replace(texto, "É", "E")
    texto = Replace(texto, "È", "E")

    ' I
    texto = Replace(texto, "Í", "I")
    texto = Replace(texto, "Ì", "I")

    ' O
    texto = Replace(texto, "Ó", "O")
    texto = Replace(texto, "Ò", "O")

    ' U
    texto = Replace(texto, "Ú", "U")
    texto = Replace(texto, "Ù", "U")

    MF_NormalizarVocales_GL = texto

End Function


Private Function MF_NormalizarVocales_EU(ByVal texto As String) As String

    ' A
    texto = Replace(texto, "Á", "A")
    texto = Replace(texto, "À", "A")

    ' E
    texto = Replace(texto, "É", "E")
    texto = Replace(texto, "È", "E")

    ' I
    texto = Replace(texto, "Í", "I")
    texto = Replace(texto, "Ì", "I")

    ' O
    texto = Replace(texto, "Ó", "O")
    texto = Replace(texto, "Ò", "O")

    ' U
    texto = Replace(texto, "Ú", "U")
    texto = Replace(texto, "Ù", "U")

    MF_NormalizarVocales_EU = texto

End Function

Private Function MF_NormalizarVocales_PT_EU(ByVal texto As String) As String

    ' Nasales
    texto = Replace(texto, "Ã", "A~")
    texto = Replace(texto, "Õ", "O~")

    ' Cerradas (circunflejo)
    texto = Replace(texto, "Â", "Â")
    texto = Replace(texto, "Ê", "Ê")
    texto = Replace(texto, "Î", "I") ' no existe en PT, pero por robustez
    texto = Replace(texto, "Ô", "Ô")
    texto = Replace(texto, "Û", "U") ' no existe en PT, robustez

    ' Abiertas (agudas)
    texto = Replace(texto, "Á", "A´")
    texto = Replace(texto, "É", "E´")
    texto = Replace(texto, "Í", "I´")
    texto = Replace(texto, "Ó", "O´")
    texto = Replace(texto, "Ú", "U´")

    ' Graves (no existen en PT, pero pueden aparecer en nombres importados)
    texto = Replace(texto, "À", "A")
    texto = Replace(texto, "È", "E")
    texto = Replace(texto, "Ì", "I")
    texto = Replace(texto, "Ò", "O")
    texto = Replace(texto, "Ù", "U")

    MF_NormalizarVocales_PT_EU = texto

End Function

Private Function MF_NormalizarVocales_PT_BR(ByVal texto As String) As String

    ' Nasales (idénticas a PT-EU)
    texto = Replace(texto, "Ã", "A~")
    texto = Replace(texto, "Õ", "O~")

    ' Cerradas ? se suavizan (PT-BR no las mantiene tan tensas)
    texto = Replace(texto, "Â", "A")
    texto = Replace(texto, "Ê", "E")
    texto = Replace(texto, "Ô", "O")

    ' Abiertas (agudas)
    texto = Replace(texto, "Á", "A´")
    texto = Replace(texto, "É", "E´")
    texto = Replace(texto, "Í", "I´")
    texto = Replace(texto, "Ó", "O´")
    texto = Replace(texto, "Ú", "U´")

    ' Graves (robustez)
    texto = Replace(texto, "À", "A")
    texto = Replace(texto, "È", "E")
    texto = Replace(texto, "Ì", "I")
    texto = Replace(texto, "Ò", "O")
    texto = Replace(texto, "Ù", "U")

    MF_NormalizarVocales_PT_BR = texto

End Function

Private Function MF_NormalizarVocales_FR(ByVal texto As String) As String

    ' A
    texto = Replace(texto, "À", "A")   ' abierta
    texto = Replace(texto, "Á", "A")   ' rara, pero robustez
    texto = Replace(texto, "Â", "Â")   ' cerrada
    texto = Replace(texto, "Ä", "A¨")  ' hiato

    ' E
    texto = Replace(texto, "È", "E")   ' abierta
    texto = Replace(texto, "É", "E´")  ' cerrada
    texto = Replace(texto, "Ê", "Ê")   ' cerrada tensa
    texto = Replace(texto, "Ë", "E¨")  ' hiato

    ' I
    texto = Replace(texto, "Ì", "I")   ' robustez
    texto = Replace(texto, "Í", "I")   ' robustez
    texto = Replace(texto, "Î", "Î")   ' cerrada
    texto = Replace(texto, "Ï", "I¨")  ' hiato

    ' O
    texto = Replace(texto, "Ò", "O")   ' robustez
    texto = Replace(texto, "Ó", "O")   ' robustez
    texto = Replace(texto, "Ô", "Ô")   ' cerrada
    texto = Replace(texto, "Ö", "O¨")  ' hiato

    ' U
    texto = Replace(texto, "Ù", "U")   ' abierta
    texto = Replace(texto, "Ú", "U")   ' robustez
    texto = Replace(texto, "Û", "Û")   ' cerrada
    texto = Replace(texto, "Ü", "U¨")  ' hiato

    MF_NormalizarVocales_FR = texto

End Function

Private Function MF_NormalizarVocales_EN_GB(ByVal texto As String) As String

    ' Solo por robustez ante nombres importados
    texto = Replace(texto, "Á", "A")
    texto = Replace(texto, "À", "A")
    texto = Replace(texto, "Ä", "A")
    texto = Replace(texto, "Â", "A")

    texto = Replace(texto, "É", "E")
    texto = Replace(texto, "È", "E")
    texto = Replace(texto, "Ë", "E")
    texto = Replace(texto, "Ê", "E")

    texto = Replace(texto, "Í", "I")
    texto = Replace(texto, "Ì", "I")
    texto = Replace(texto, "Ï", "I")
    texto = Replace(texto, "Î", "I")

    texto = Replace(texto, "Ó", "O")
    texto = Replace(texto, "Ò", "O")
    texto = Replace(texto, "Ö", "O")
    texto = Replace(texto, "Ô", "O")

    texto = Replace(texto, "Ú", "U")
    texto = Replace(texto, "Ù", "U")
    texto = Replace(texto, "Ü", "U")
    texto = Replace(texto, "Û", "U")

    MF_NormalizarVocales_EN_GB = texto

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

Private Function MF_NormalizarVocales_General(ByVal texto As String) As String
    
    Dim i As Integer
    Dim c As String
    Dim Res As String

    c = ""
    Res = ""
    
    For i = 1 To Len(texto)
        c = Mid(texto, i, 1)
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


