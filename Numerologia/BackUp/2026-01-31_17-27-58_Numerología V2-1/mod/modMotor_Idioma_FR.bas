Attribute VB_Name = "modMotor_Idioma_FR"

Option Compare Database
Option Explicit

'=============
'== Francés ==
'=============
Public Sub MF_SilabearAjustesFR( _
        ByVal Texto As String, _
        ByRef silabas As Collection _
    )

    Dim i As Long

    ' ============================================================
    ' 1. ELIMINAR LA E MUDA FINAL
    ' ============================================================
    If Right$(Texto, 1) = "E" Then
        Call MF_EliminarVocalFinal(silabas, Len(Texto))
    End If

    ' ============================================================
    ' 2. AGRUPAR NASALES (an, am, en, em, in, im, ain, ein, un, um, on, om)
    ' ============================================================
    Dim nasales As Variant
    nasales = Array("AN", "AM", "EN", "EM", "IN", "IM", "AIN", "EIN", "UN", "UM", "ON", "OM")

    For i = 2 To Len(Texto)
        Dim seq2 As String, seq3 As String
        seq2 = Mid$(Texto, i - 1, 2)
        seq3 = Mid$(Texto, i - 2, 3)

        If EsMiembro(seq3, nasales) Then
            Call MF_UnirVocalesEnDiptongo(silabas, i - 1)
        ElseIf EsMiembro(seq2, nasales) Then
            Call MF_UnirVocalesEnDiptongo(silabas, i)
        End If
    Next i

    ' ============================================================
    ' 3. ATAQUES CONSONÁNTICOS PERMITIDOS EN FRANCÉS
    ' ============================================================
    Dim ataques As Variant
    ataques = Array("TR", "DR", "PR", "BR", "CR", "GR", "FR", "FL", "CL", "GL", "PL")

    For i = 2 To Len(Texto) - 1
        Dim par As String
        par = Mid$(Texto, i, 2)

        If EsMiembro(par, ataques) Then
            Call MF_UnirConsonantesEnAtaque(silabas, i)
        End If
    Next i

    ' ============================================================
    ' 4. HIATOS POR DIÉRESIS (ï, ë, ü)
    ' ============================================================
    Dim dieresis As String
    dieresis = "ÏËÜ"

    For i = 2 To Len(Texto)
        If InStr(dieresis, Mid$(Texto, i, 1)) > 0 Then
            Call MF_ForzarDivisionSilabica(silabas, i)
        End If
    Next i

End Sub

Public Sub MF_MarcarTonicaFrances( _
        ByVal Texto As String, _
        ByRef esTonica() As Boolean _
    )

    Dim silabas As Collection
    Dim idxTonica As Long
    Dim inicio As Long, fin As Long
    Dim i As Long

    ' --------------------------------------------------------
    ' 1. Silabear palabra
    ' --------------------------------------------------------
    'Set silabas = MF_SilabearCastellano(Texto)
    Set silabas = MF_Silabear(Texto, "fr")


    If silabas Is Nothing Then Exit Sub
    If silabas.Count = 0 Then Exit Sub

    ' --------------------------------------------------------
    ' 2. Acento francés: SIEMPRE en la última sílaba fonética
    ' --------------------------------------------------------
    idxTonica = silabas.Count

    ' --------------------------------------------------------
    ' 3. Marcar índices tónicos
    ' --------------------------------------------------------
    inicio = silabas(idxTonica)(1)
    fin = silabas(idxTonica)(2)

    For i = inicio To fin
        esTonica(i) = True
    Next i

End Sub


' ============================================================
'   ReglasFrances (FR)
'   Devuelve idFonema según la fonética del francés.
'   Si no aplica, devuelve 0 para que el motor siga probando.
' ============================================================

Public Function ReglasFrances( _
        ByVal graf As String, _
        ByVal ant As String, _
        ByVal sig As String, _
        ByVal esTonica As Boolean _
    ) As Byte

    Dim g As String
    g = UCase$(graf)

    ' ============================================================
    '   TRIGRAFEMAS
    ' ============================================================

    ' GÜE / GÜI ? /gw/ ? id 57
    If g = "GÜE" Or g = "GÜI" Then
        ReglasFrances = 57
        Exit Function
    End If

    ' GUE / GUI ? /g/ ? id 31
    If g = "GUE" Or g = "GUI" Then
        ReglasFrances = 31
        Exit Function
    End If

    ' QUE / QUI ? /k/ ? id 30
    If g = "QUE" Or g = "QUI" Then
        ReglasFrances = 30
        Exit Function
    End If


    ' ============================================================
    '   DÍGRAFOS Y CASOS ESPECIALES
    ' ============================================================

    ' CH ? /?/ ? id 36
    If g = "CH" Then
        ReglasFrances = 36
        Exit Function
    End If

    ' GN ? /?/ ? id 41
    If g = "GN" Then
        ReglasFrances = 41
        Exit Function
    End If

    ' J ? /?/ ? id 37
    If g = "J" Then
        ReglasFrances = 37
        Exit Function
    End If

    ' G + E/I/Y ? /?/ ? id 37
    If g = "G" And (sig = "E" Or sig = "I" Or sig = "Y") Then
        ReglasFrances = 37
        Exit Function
    End If

    ' S entre vocales ? /z/ ? id 35
    If g = "S" And (ant Like "[AEIOU]" And sig Like "[AEIOU]") Then
        ReglasFrances = 35
        Exit Function
    End If

    ' Ç ? /s/ ? id 34
    If g = "Ç" Then
        ReglasFrances = 34
        Exit Function
    End If


    ' ============================================================
    '   NASALIZACIONES
    ' ============================================================

    ' AN / AM / EN / EM ? /?~/ ? id 2
    If g = "AN" Or g = "AM" Or g = "EN" Or g = "EM" Then
        ReglasFrances = 2
        Exit Function
    End If

    ' IN / IM / AIN / EIN / EIM / YN / YM ? /?~/ ? id 3
    If g = "IN" Or g = "IM" Or g = "AIN" Or g = "EIN" Or g = "EIM" Or g = "YN" Or g = "YM" Then
        ReglasFrances = 3
        Exit Function
    End If

    ' ON / OM ? /?~/ ? id 4
    If g = "ON" Or g = "OM" Then
        ReglasFrances = 4
        Exit Function
    End If

    ' UN / UM ? /œ~/ ? id 3 (aproximación razonable)
    If g = "UN" Or g = "UM" Then
        ReglasFrances = 3
        Exit Function
    End If


    ' ============================================================
    '   DÍGRAFOS VOCÁLICOS (diptongos franceses)
    ' ============================================================

    ' OI ? /wa/ ? id 18
    If g = "OI" Then ReglasFrances = 18: Exit Function

    ' AI ? /?/ ? id 6
    If g = "AI" Then ReglasFrances = 6: Exit Function

    ' EI ? /e/ ? id 5
    If g = "EI" Then ReglasFrances = 5: Exit Function

    ' OU ? /u/ ? id 10
    If g = "OU" Then ReglasFrances = 10: Exit Function


    ' ============================================================
    '   MONÓGRAFOS — VOCALES
    ' ============================================================

    If g = "A" Then ReglasFrances = 1: Exit Function
    If g = "E" Then ReglasFrances = 5: Exit Function
    If g = "I" Then ReglasFrances = 9: Exit Function
    If g = "O" Then ReglasFrances = 7: Exit Function
    If g = "U" Then ReglasFrances = 10: Exit Function


    ' ============================================================
    '   MONÓGRAFOS — CONSONANTES
    ' ============================================================

    If g = "P" Then ReglasFrances = 26: Exit Function
    If g = "B" Then ReglasFrances = 27: Exit Function
    If g = "T" Then ReglasFrances = 28: Exit Function
    If g = "D" Then ReglasFrances = 29: Exit Function
    If g = "K" Then ReglasFrances = 30: Exit Function
    If g = "G" Then ReglasFrances = 31: Exit Function

    If g = "F" Then ReglasFrances = 32: Exit Function
    If g = "V" Then ReglasFrances = 33: Exit Function

    ' C + E/I/Y ? /s/
    If g = "C" And (sig = "E" Or sig = "I" Or sig = "Y") Then
        ReglasFrances = 34
        Exit Function
    End If

    ' S ? /s/
    If g = "S" Then ReglasFrances = 34: Exit Function

    ' X ? /ks/ o /gz/ ? simplificamos a /s/ (el motor segmenta la K aparte)
    If g = "X" Then
        ReglasFrances = 34
        Exit Function
    End If

    If g = "M" Then ReglasFrances = 39: Exit Function
    If g = "N" Then ReglasFrances = 40: Exit Function

    ' L ? /l/
    If g = "L" Then ReglasFrances = 43: Exit Function

    ' R ? /?/ ? id 47
    If g = "R" Then
        ReglasFrances = 47
        Exit Function
    End If

    ' H ? muda ? id 38
    If g = "H" Then
        ReglasFrances = 38
        Exit Function
    End If


    ' ============================================================
    '   SI NO APLICA, DEVOLVER 0
    ' ============================================================
    ReglasFrances = 0

End Function

'============================================================================================

Public Function MF_NormalizarVocales_FR(ByVal Texto As String) As String

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

