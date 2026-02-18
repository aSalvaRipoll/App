Attribute VB_Name = "modMotor_Idioma_EN_GB"

Option Compare Database
Option Explicit

'===========================
'== Inglés (Gran Bretaña) ==
'===========================

Public Sub MF_MarcarTonica_EN_GB( _
        ByVal texto As String, _
        ByRef esTonica() As Boolean _
    )

    Dim Silabas As Collection
    Dim idxTonica As Long
    Dim inicio As Long, fin As Long
    Dim i As Long
    Dim ultima As String
    Dim ult2 As String

    ' 1. Silabear palabra (motor con revisión)
    Set Silabas = Silabear_EN_ConRevision(texto)

    If Silabas Is Nothing Then Exit Sub
    If Silabas.Count = 0 Then Exit Sub

    ultima = Right$(texto, 1)
    If Len(texto) >= 2 Then ult2 = Right$(texto, 2)

    ' 2. Heurísticas principales

    ' 2.1. Palabras largas --> antepenúltima
    If Silabas.Count >= 4 Then
        idxTonica = Silabas.Count - 2
        GoTo Marcar
    End If

    ' 2.2. Sufijos débiles --> penúltima
    If ult2 = "ER" Or ult2 = "OR" Or ult2 = "AN" Or ult2 = "ON" Or _
       ult2 = "LY" Or ult2 = "EY" Or ult2 = "AY" Or ult2 = "RY" Then
        If Silabas.Count >= 2 Then
            idxTonica = Silabas.Count - 1
            GoTo Marcar
        End If
    End If

    ' 2.3. Prefijos comunes --> segunda sílaba
    If Left$(texto, 2) = "MC" Or Left$(texto, 2) = "DE" Or _
       Left$(texto, 2) = "LA" Or Left$(texto, 2) = "LE" Then
        If Silabas.Count >= 2 Then
            idxTonica = 2
            GoTo Marcar
        End If
    End If

    ' 2.4. Bisílabas --> primera
    If Silabas.Count = 2 Then
        idxTonica = 1
        GoTo Marcar
    End If

    ' 2.5. Trisílabas --> primera
    If Silabas.Count = 3 Then
        idxTonica = 1
        GoTo Marcar
    End If

    ' 2.6. Monosílabas --> única
    If Silabas.Count = 1 Then
        idxTonica = 1
        GoTo Marcar
    End If

Marcar:
    inicio = Silabas(idxTonica)(1)
    fin = Silabas(idxTonica)(2)

    For i = inicio To fin
        esTonica(i) = True
    Next i

End Sub

Public Function Silabear_EN_ConRevision(ByVal texto As String) As Collection

    Dim col As Collection
    Dim resultado As New Collection
    Dim item As Variant
    Dim s As String
    Dim partes As Variant
    Dim p As Variant
    Dim inicio As Long, fin As Long
    Dim valido As Boolean
    Dim msg As String

    ' 1. Silabear automáticamente (motor puro inglés)
    Set col = Silabear_EN(texto)

    ' 2. Convertir a string con "-"
    For Each item In col
        s = s & Mid$(texto, item(0), item(1) - item(0) + 1) & "-"
    Next item
    If Len(s) > 0 Then s = Left$(s, Len(s) - 1)

    ' 3. Bucle de validación con formulario
    Do
        valido = True
        msg = ""

        s = RevisarSilabas_EnFormulario(texto, s)

        If s = "" Then
            Set Silabear_EN_ConRevision = col
            Exit Function
        End If

        If Left$(s, 1) = "-" Or Right$(s, 1) = "-" Then
            valido = False
            msg = "Cannot start or end with '-'."
        End If

        If InStr(s, "--") > 0 Then
            valido = False
            msg = "Cannot contain empty syllables ('--')."
        End If

        Dim reconstruido As String
        Dim textoSinEspacios As String

        reconstruido = Replace(s, "-", "")
        reconstruido = Replace(reconstruido, " ", "")
        textoSinEspacios = Replace(texto, " ", "")

        If UCase$(reconstruido) <> UCase$(textoSinEspacios) Then
            valido = False
            msg = "Syllables do not match the original text."
        End If

        If Not valido Then
            MsgBox msg, vbExclamation, "Syllable error"
        End If

    Loop Until valido

    ' 4. Reconstruir colección final
    partes = Split(s, "-")
    inicio = 1

    For Each p In partes
        fin = inicio + Len(p) - 1
        resultado.Add Array(inicio, fin)
        inicio = fin + 1
    Next p

    Set Silabear_EN_ConRevision = resultado

End Function

Public Function Silabear_EN(ByVal texto As String) As Collection

    Dim col As New Collection
    Dim i As Long, ini As Long
    Dim c1 As String, c2 As String, c3 As String
    Dim par As String

    texto = Trim$(texto)
    If Len(texto) = 0 Then
        Set Silabear_EN = col
        Exit Function
    End If

    ini = 1

    For i = 2 To Len(texto)

        c1 = Mid$(texto, i - 1, 1)
        c2 = Mid$(texto, i, 1)
        par = c1 & c2

        ' ---------------------------------------------------------
        ' 0. Spaces --> separate words
        ' ---------------------------------------------------------
        If c1 = " " Then
            If i - 2 >= ini Then col.Add Array(ini, i - 2)
            ini = i
            GoTo siguiente
        End If

        If c2 = " " Then
            col.Add Array(ini, i - 1)
            ini = i + 1
            GoTo siguiente
        End If

        ' ---------------------------------------------------------
        ' 1. VV --> hiatus (English rarely forms stable VV diphthongs)
        ' ---------------------------------------------------------
        If EsVocal_EN(c1) And EsVocal_EN(c2) Then
            col.Add Array(ini, i - 1)
            ini = i
            GoTo siguiente
        End If

        ' ---------------------------------------------------------
        ' 2. VCV --> V | CV
        ' ---------------------------------------------------------
        If EsVocal_EN(c1) And EsConsonant_EN(c2) Then
            If i < Len(texto) Then
                c3 = Mid$(texto, i + 1, 1)
                If EsVocal_EN(c3) Then
                    col.Add Array(ini, i - 1)
                    ini = i
                    GoTo siguiente
                End If
            End If
        End If

        ' ---------------------------------------------------------
        ' 3. CCV --> C | CV
        ' ---------------------------------------------------------
        If EsConsonant_EN(c1) And EsConsonant_EN(c2) Then
            If i < Len(texto) Then
                c3 = Mid$(texto, i + 1, 1)
                If EsVocal_EN(c3) Then
                    If Not EsGrupInseparable_EN(par) Then
                        col.Add Array(ini, i - 1)
                        ini = i
                        GoTo siguiente
                    End If
                End If
            End If
        End If

siguiente:
    Next i

    ' Última sílaba
    If ini <= Len(texto) Then
        col.Add Array(ini, Len(texto))
    End If

    Set Silabear_EN = col

End Function

Public Function EsVocal_EN(c As String) As Boolean
    EsVocal_EN = InStr("AEIOUYaeiouy", c) > 0
End Function

Public Function EsConsonant_EN(c As String) As Boolean
    EsConsonant_EN = Not EsVocal_EN(c) And c <> " "
End Function

Public Function EsGrupInseparable_EN(par As String) As Boolean
    Select Case UCase$(par)
        Case "BR", "BL", "CR", "CL", "DR", "TR", "PR", "PL", "GR", "GL", "FR", "FL"
            EsGrupInseparable_EN = True
        Case Else
            EsGrupInseparable_EN = False
    End Select
End Function


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

    ' GUE / GUI --> /g/ (U muda)
    If g = "GUE" Or g = "GUI" Then
        ReglasIngles = 31
        Exit Function
    End If

    ' QUE / QUI --> /k/
    If g = "QUE" Or g = "QUI" Then
        ReglasIngles = 30
        Exit Function
    End If


    ' ============================================================
    '   DÍGRAFOS ESPECIALES
    ' ============================================================

    ' TH --> /?/
    If g = "TH" Then
        ReglasIngles = 54
        Exit Function
    End If

    ' DH --> /ð/
    If g = "DH" Then
        ReglasIngles = 55
        Exit Function
    End If

    ' SH --> /?/
    If g = "SH" Then
        ReglasIngles = 36
        Exit Function
    End If

    ' CH --> /t?/
    If g = "CH" Then
        ReglasIngles = 50
        Exit Function
    End If

    ' PH --> /f/
    If g = "PH" Then
        ReglasIngles = 32
        Exit Function
    End If

    ' NG --> /?/
    If g = "NG" Then
        ReglasIngles = 42
        Exit Function
    End If

    ' WH --> /w/
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

    ' EA / EE --> /i/
    If g = "EA" Or g = "EE" Then ReglasIngles = 9: Exit Function

    ' IE --> /ai/
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

    ' J --> /d?/
    If g = "J" Then ReglasIngles = 51: Exit Function

    ' Y --> /j/
    If g = "Y" Then ReglasIngles = 48: Exit Function

    ' W --> /w/
    If g = "W" Then ReglasIngles = 49: Exit Function

    ' X --> /ks/ --> devolvemos /s/
    If g = "X" Then ReglasIngles = 34: Exit Function

    If g = "M" Then ReglasIngles = 39: Exit Function
    If g = "N" Then ReglasIngles = 40: Exit Function

    If g = "L" Then ReglasIngles = 43: Exit Function
    If g = "R" Then ReglasIngles = 45: Exit Function

    ' H --> aspiración suave
    If g = "H" Then ReglasIngles = 38: Exit Function


    ' ============================================================
    '   SI NO APLICA, DEVOLVER 0
    ' ============================================================
    ReglasIngles = 0

End Function

'============================================================================================

Public Function MF_NormalizarVocales_EN_GB(ByVal texto As String) As String

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

