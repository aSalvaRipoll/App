Attribute VB_Name = "modMotorFonetico_V2_1_Aux"
Option Compare Database
Option Explicit


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

        Case "EN"
            texto = MF_NormalizarVocales_EN(texto)

        Case Else
            texto = MF_NormalizarVocales_General(texto)

    End Select

    MF_NormalizarVocalesPorIdioma = texto

End Function


'Public Function MF_NormalizarVocalesPorIdioma( _
'    ByVal texto As String, _
'    ByVal idioma As String _
'    ) As String
'
'    Select Case UCase$(idioma)
'
'        Case "ES"
'            texto = MF_NormalizarVocales_ES(texto)
'
'        Case "CA", "CA-IB", "CA-VA"
'            texto = MF_NormalizarVocales_CA(texto)
'
''        Case "CA-IB"
''            texto = MF_NormalizarVocales_CA_IB(texto)
''
''        Case "CA-VA"
''            texto = MF_NormalizarVocales_CA_VA(texto)
'
'        Case "GL"
'            texto = MF_NormalizarVocales_GL(texto)
'
'        Case "EU"
'            texto = MF_NormalizarVocales_EU(texto)
'
'        Case Else
'            texto = MF_NormalizarVocales_General(texto)
'
'    End Select
'
'    MF_NormalizarVocalesPorIdioma = texto
'
'End Function

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

Private Function MF_NormalizarVocales_EN(ByVal texto As String) As String

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

    MF_NormalizarVocales_EN = texto

End Function


Private Function MF_NormalizarVocales_General(ByVal texto As String) As String
    
    Dim i As Integer
    Dim C As String
    Dim Res As String

    C = ""
    Res = ""
    
    For i = 1 To Len(texto)
        C = Mid(texto, i, 1)
        Select Case C
            
            Case "Á", "À", "Ä"
                C = "A"
    
            Case "É", "È", "Ë"
                C = "E"
    
            Case "Í", "Ï"
                C = "I"
            
            Case "Ó", "Ò", "Ö"
                C = "O"
    
            Case "Ú", "Ü"
                C = "U"
        End Select
        
        Res = Res & C
    Next i
    
    MF_NormalizarVocales_General = Res

End Function


Public Function EsVocal(ByVal C As String) As Boolean
    Select Case UCase$(C)
        Case "A", "Á", "À", _
             "E", "É", "È", _
             "I", "Í", "Ï", _
             "O", "Ó", "Ò", _
             "U", "Ú", "Ü"
            EsVocal = True
        Case Else
            EsVocal = False
    End Select
End Function

Public Function EsConsonante(ByVal C As String) As Boolean
    If C = "" Then
        EsConsonante = False
    Else
        EsConsonante = Not EsVocal(C)
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
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tbmFonemas ORDER BY idFonema")

    Do While Not rs.EOF
        Set f = New clsFonema
        f.idFonema = rs!idFonema
        f.GrafemaOri = rs!fonema
        f.EsVocal = rs!EsVocal
        f.Valor = rs!Valor
        ' ... cualquier otra propiedad

        colFonemas.Add f, CStr(f.idFonema)
        rs.MoveNext
    Loop

    rs.Close
End Sub

Public Function BuscarExcepcion(ByVal palabra As String, ByVal idioma As String) As String

    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT FonemaCompleto FROM tbmDicExcepciones " & _
        "WHERE Idioma = '" & idioma & "' AND Palabra = '" & UCase(palabra) & "' AND Activo = True")

    If Not rs.EOF Then
        BuscarExcepcion = rs!FonemaCompleto
    Else
        BuscarExcepcion = ""
    End If

    rs.Close
    Set rs = Nothing

End Function


