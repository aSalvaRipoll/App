Attribute VB_Name = "modMotor_Idioma_EN_GB"

Option Compare Database
Option Explicit

'===========================
'== Inglés (Gran Bretaña) ==
'===========================

Public Sub MF_SilabearAjustesEN_GB( _
        ByVal Texto As String, _
        ByRef silabas As Collection _
    )

    Dim i As Long

    ' ============================================================
    ' 1. ATAQUES CONSONÁNTICOS PROPIOS DEL INGLÉS
    ' ============================================================
    Dim ataques As Variant
    ataques = Array( _
        "PL", "PR", "BL", "BR", "CL", "CR", "GL", "GR", _
        "FL", "FR", "TR", "DR", "SK", "SL", "SM", "SN", _
        "SP", "ST", "SW", "SH", "TH", "WH", _
        "STR", "SPR", "SPL", "SCR", "SHR", "THR" _
    )

    ' Primero ataques de 3 letras
    For i = 3 To Len(Texto) - 1
        Dim tri As String
        tri = Mid$(Texto, i - 1, 3)

        If EsMiembro(tri, ataques) Then
            Call MF_UnirConsonantesEnAtaque(silabas, i - 1)
        End If
    Next i

    ' Luego ataques de 2 letras
    For i = 2 To Len(Texto) - 1
        Dim par As String
        par = Mid$(Texto, i, 2)

        If EsMiembro(par, ataques) Then
            Call MF_UnirConsonantesEnAtaque(silabas, i)
        End If
    Next i

    ' ============================================================
    ' 2. DIPTONGOS INGLESES (refuerzo)
    ' ============================================================
    Dim dipt As Variant
    dipt = Array( _
        "AI", "AY", "EI", "EY", "OI", "OY", _
        "AU", "AW", "OU", "OW", _
        "EA", "EE", "IE", "OA" _
    )

    For i = 2 To Len(Texto)
        Dim dv As String
        dv = Mid$(Texto, i - 1, 2)

        If EsMiembro(dv, dipt) Then
            Call MF_UnirVocalesEnDiptongo(silabas, i)
        End If
    Next i

    ' ============================================================
    ' 3. HIATOS OBLIGATORIOS (cooperate, naive, reenter…)
    ' ============================================================
    Dim hiatos As Variant
    hiatos = Array("COO", "OO", "RE", "NAI", "REI")

    For i = 3 To Len(Texto)
        Dim tri2 As String
        tri2 = Mid$(Texto, i - 2, 3)

        If EsMiembro(tri2, hiatos) Then
            Call MF_ForzarDivisionSilabica(silabas, i - 1)
        End If
    Next i

    ' ============================================================
    ' 4. DIÉRESIS (ï, ë) ? rompen diptongo
    ' ============================================================
    Dim dieresis As String
    dieresis = "ÏË"

    For i = 2 To Len(Texto)
        If InStr(dieresis, Mid$(Texto, i, 1)) > 0 Then
            Call MF_ForzarDivisionSilabica(silabas, i)
        End If
    Next i

End Sub
Public Sub MF_MarcarTonicaIngles( _
        ByVal Texto As String, _
        ByRef esTonica() As Boolean _
    )

    Dim silabas As Collection
    Dim idxTonica As Long
    Dim inicio As Long, fin As Long
    Dim i As Long
    Dim ultima As String
    Dim ult2 As String

    ' --------------------------------------------------------
    ' 1. Silabear palabra
    ' --------------------------------------------------------
    'Set silabas = MF_SilabearCastellano(Texto)
    Set silabas = MF_Silabear(Texto, "en-gb")

    If silabas Is Nothing Then Exit Sub
    If silabas.Count = 0 Then Exit Sub

    ultima = Right$(Texto, 1)
    If Len(Texto) >= 2 Then ult2 = Right$(Texto, 2)

    ' --------------------------------------------------------
    ' 2. Heurísticas principales
    ' --------------------------------------------------------

    ' 2.1. Palabras muy largas ? antepenúltima
    If silabas.Count >= 4 Then
        idxTonica = silabas.Count - 2
        GoTo Marcar
    End If

    ' 2.2. Sufijos débiles ? penúltima
    If ult2 = "ER" Or ult2 = "OR" Or ult2 = "AN" Or ult2 = "ON" Or _
       ult2 = "LY" Or ult2 = "EY" Or ult2 = "AY" Or ult2 = "RY" Then
        If silabas.Count >= 2 Then
            idxTonica = silabas.Count - 1
            GoTo Marcar
        End If
    End If

    ' 2.3. Prefijos comunes ? segunda sílaba
    If Left$(Texto, 2) = "MC" Or Left$(Texto, 2) = "DE" Or _
       Left$(Texto, 2) = "LA" Or Left$(Texto, 2) = "LE" Then
        If silabas.Count >= 2 Then
            idxTonica = 2
            GoTo Marcar
        End If
    End If

    ' 2.4. Bisílabas ? primera sílaba
    If silabas.Count = 2 Then
        idxTonica = 1
        GoTo Marcar
    End If

    ' 2.5. Trisílabas ? primera sílaba
    If silabas.Count = 3 Then
        idxTonica = 1
        GoTo Marcar
    End If

    ' 2.6. Monosílabas ? única sílaba
    If silabas.Count = 1 Then
        idxTonica = 1
        GoTo Marcar
    End If

Marcar:
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
'   ReglasIngles (ENG)
'   Devuelve idFonema según la fonética del inglés.
'   Si no aplica, devuelve 0 para que el motor siga probando.
' ============================================================

Public Function ReglasIngles( _
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

    ' GUE / GUI ? /g/ (U muda)
    If g = "GUE" Or g = "GUI" Then
        ReglasIngles = 31
        Exit Function
    End If

    ' QUE / QUI ? /k/
    If g = "QUE" Or g = "QUI" Then
        ReglasIngles = 30
        Exit Function
    End If


    ' ============================================================
    '   DÍGRAFOS ESPECIALES
    ' ============================================================

    ' TH ? /?/
    If g = "TH" Then
        ReglasIngles = 54
        Exit Function
    End If

    ' DH ? /ð/
    If g = "DH" Then
        ReglasIngles = 55
        Exit Function
    End If

    ' SH ? /?/
    If g = "SH" Then
        ReglasIngles = 36
        Exit Function
    End If

    ' CH ? /t?/
    If g = "CH" Then
        ReglasIngles = 50
        Exit Function
    End If

    ' PH ? /f/
    If g = "PH" Then
        ReglasIngles = 32
        Exit Function
    End If

    ' NG ? /?/
    If g = "NG" Then
        ReglasIngles = 42
        Exit Function
    End If

    ' WH ? /w/
    If g = "WH" Then
        ReglasIngles = 49
        Exit Function
    End If


    ' ============================================================
    '   DÍGRAFOS VOCÁLICOS (diptongos ingleses)
    ' ============================================================

    If g = "AI" Or g = "AY" Then ReglasIngles = 13: Exit Function
    If g = "EI" Then ReglasIngles = 13: Exit Function
    If g = "OI" Or g = "OY" Then ReglasIngles = 14: Exit Function
    If g = "OU" Or g = "OW" Then ReglasIngles = 15: Exit Function
    If g = "AU" Or g = "AW" Then ReglasIngles = 16: Exit Function

    ' EA / EE ? /i/
    If g = "EA" Or g = "EE" Then ReglasIngles = 9: Exit Function

    ' IE ? /ai/
    If g = "IE" Then ReglasIngles = 12: Exit Function


    ' ============================================================
    '   MONÓGRAFOS — VOCALES
    ' ============================================================

    If g = "A" Then ReglasIngles = 1: Exit Function
    If g = "E" Then ReglasIngles = 5: Exit Function
    If g = "I" Then ReglasIngles = 9: Exit Function
    If g = "O" Then ReglasIngles = 7: Exit Function
    If g = "U" Then ReglasIngles = 10: Exit Function


    ' ============================================================
    '   MONÓGRAFOS — CONSONANTES
    ' ============================================================

    If g = "P" Then ReglasIngles = 26: Exit Function
    If g = "B" Then ReglasIngles = 27: Exit Function
    If g = "T" Then ReglasIngles = 28: Exit Function
    If g = "D" Then ReglasIngles = 29: Exit Function
    If g = "K" Or g = "C" Then ReglasIngles = 30: Exit Function
    If g = "G" Then ReglasIngles = 31: Exit Function

    If g = "F" Then ReglasIngles = 32: Exit Function
    If g = "V" Then ReglasIngles = 33: Exit Function
    If g = "S" Then ReglasIngles = 34: Exit Function
    If g = "Z" Then ReglasIngles = 35: Exit Function

    ' J ? /d?/
    If g = "J" Then ReglasIngles = 51: Exit Function

    ' Y ? /j/
    If g = "Y" Then ReglasIngles = 48: Exit Function

    ' W ? /w/
    If g = "W" Then ReglasIngles = 49: Exit Function

    ' X ? /ks/ ? devolvemos /s/
    If g = "X" Then ReglasIngles = 34: Exit Function

    If g = "M" Then ReglasIngles = 39: Exit Function
    If g = "N" Then ReglasIngles = 40: Exit Function

    If g = "L" Then ReglasIngles = 43: Exit Function
    If g = "R" Then ReglasIngles = 45: Exit Function

    ' H ? aspiración suave
    If g = "H" Then ReglasIngles = 38: Exit Function


    ' ============================================================
    '   SI NO APLICA, DEVOLVER 0
    ' ============================================================
    ReglasIngles = 0

End Function

'============================================================================================

Public Function MF_NormalizarVocales_EN_GB(ByVal Texto As String) As String

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
