Attribute VB_Name = "modMotor_Idioma_ES"

Option Compare Database
Option Explicit

'=================
'==  Castellano ==
'=================

Public Sub MF_MarcarTonica_ES( _
        ByVal Texto As String, _
        ByRef esTonica() As Boolean _
    )

    Dim i As Long
    Dim Silabas As Collection
    Dim idxTonica As Long
    Dim inicio As Long, fin As Long

    ' --------------------------------------------------------
    ' 1. Silabear palabra (con revisión manual)
    ' --------------------------------------------------------
    Set Silabas = Silabear_ES_ConRevision(Texto)

    If Silabas.Count = 0 Then Exit Sub

    ' --------------------------------------------------------
    ' 2. Determinar sílaba tónica
    ' --------------------------------------------------------
    idxTonica = MF_DetectarTonicaCastellano(Texto, Silabas)

    If idxTonica = 0 Then Exit Sub

    ' --------------------------------------------------------
    ' 3. Marcar índices de la sílaba tónica
    ' --------------------------------------------------------
    inicio = Silabas(idxTonica)(0)
    fin = Silabas(idxTonica)(1)

    For i = inicio To fin
        esTonica(i) = True
    Next i

End Sub


' ============================================================
'   Silabear_ES — Silabeador para nombres y apellidos en español
'   - Respeta espacios entre palabras
'   - No mezcla sílabas entre palabras
'   - Aplica reglas fonéticas del español
'   - Detecta dígrafos (CH, LL, RR)
'   - Detecta grupos consonánticos inseparables (BR, CR, TR…)
'   - Trata la H como consonante muda (no rompe sílabas)
'   - Elimina H final aislada
'   - Devuelve posiciones absolutas (ini, fin)
' ============================================================

' ============================================================
'   REVISIÓN MANUAL MEDIANTE INPUTBOX
' ============================================================

Public Function Silabear_ES_ConRevision(ByVal Texto As String) As Collection

    Dim col As Collection
    Dim resultado As New Collection
    Dim item As Variant
    Dim s As String
    Dim partes As Variant
    Dim p As Variant
    Dim inicio As Long, fin As Long
    Dim valido As Boolean
    Dim msg As String

    ' 1. Silabear automáticamente
    Set col = Silabear_ES(Texto)

    ' 2. Convertir a string con separador visual "-"
    '    (las sílabas no incluyen espacios; los espacios están en Texto)
    For Each item In col
        s = s & Mid$(Texto, item(0), item(1) - item(0) + 1) & "-"
    Next item
    If Len(s) > 0 Then s = Left$(s, Len(s) - 1)

    ' ============================================================
    ' 3. Bucle de validación
    ' ============================================================
    Do
        valido = True
        msg = ""

        's = InputBox("Revisa o corrige las sílabas:" & vbCrLf & _
                     "(usa '-' como separador entre sílabas)", _
                     "Revisión de sílabas", s)


        s = RevisarSilabas_EnFormulario(Texto, s)


        ' Si el usuario cancela --> devolver silabeo automático
        If s = "" Then
            Set Silabear_ES_ConRevision = col
            Exit Function
        End If

        ' No recortamos espacios: pueden ser parte del nombre compuesto
        ' pero sí limpiamos dobles guiones accidentales

        ' Validación 1: no puede empezar ni acabar con "-"
        If Left$(s, 1) = "-" Or Right$(s, 1) = "-" Then
            valido = False
            msg = "No puede empezar ni terminar con '-'."
        End If

        ' Validación 2: no puede contener "--" (sílabas vacías)
        If InStr(s, "--") > 0 Then
            valido = False
            msg = "No puede haber sílabas vacías ('--')."
        End If

        ' Validación 3: comprobar que las sílabas reconstruyen el texto original
        ' Ignorando espacios y separadores
        Dim reconstruido As String
        Dim textoSinEspacios As String

        ' Quitar separadores "-" y espacios del string de sílabas
        reconstruido = Replace(s, "-", "")
        reconstruido = Replace(reconstruido, " ", "")

        ' Quitar espacios del texto original
        textoSinEspacios = Replace(Texto, " ", "")

        If UCase$(reconstruido) <> UCase$(textoSinEspacios) Then
            valido = False
            msg = "Las sílabas no coinciden con el texto original (ignorando espacios)."
        End If

        ' Si no es válido --> mostrar mensaje y repetir
        If Not valido Then
            MsgBox msg, vbExclamation, "Error en las sílabas"
        End If

    Loop Until valido

    ' ============================================================
    ' 4. Reconstruir colección válida
    ' ============================================================
    partes = Split(s, "-")
    inicio = 1

    For Each p In partes
        fin = inicio + Len(p) - 1
        resultado.Add Array(inicio, fin)
        inicio = fin + 1
    Next p

    Set Silabear_ES_ConRevision = resultado

End Function

Private Function EsVocalDebilTilde(c As String) As Boolean
    EsVocalDebilTilde = (c = "Í" Or c = "Ú")
End Function


' ============================================================
'   Silabear_ES — Silabeador para nombres y apellidos en español
' ============================================================

Public Function Silabear_ES(ByVal Texto As String) As Collection

    Dim col As New Collection
    Dim i As Long, ini As Long
    Dim c1 As String, c2 As String, c3 As String
    Dim par As String

    Texto = Trim$(Texto)
    If Len(Texto) = 0 Then
        Set Silabear_ES = col
        Exit Function
    End If

    ini = 1

    For i = 2 To Len(Texto)

        c1 = Mid$(Texto, i - 1, 1)
        c2 = Mid$(Texto, i, 1)
        par = c1 & c2

        ' =====================================================
        ' 0. ESPACIOS
        ' =====================================================
        If c1 = " " Then
            If i - 2 >= ini Then col.Add Array(ini, i - 2)
            ini = i
            GoTo Siguiente
        End If

        If c2 = " " Then
            col.Add Array(ini, i - 1)
            ini = i + 1
            GoTo Siguiente
        End If

        ' =====================================================
        ' 1. DÍGRAFOS INSEPARABLES (CH, LL, RR)
        ' =====================================================
        If par = "CH" Or par = "LL" Or par = "RR" Then
            GoTo Siguiente
        End If

        ' =====================================================
        ' 2. GRUPOS CONSONÁNTICOS INSEPARABLES (BR, CR, TR…)
        ' =====================================================
        If EsConsonante_ES(c1) And EsConsonante_ES(c2) Then
            If EsGrupoInseparable_ES(par) Then
                GoTo Siguiente
            End If
        End If

        ' =====================================================
        ' 3. PATRÓN CCV --> C | CV
        ' =====================================================
        If EsConsonante_ES(c1) And EsConsonante_ES(c2) Then
            If i < Len(Texto) Then
                c3 = Mid$(Texto, i + 1, 1)
                If EsVocal_ES(c3) Then
                    If Not EsGrupoInseparable_ES(par) Then
                        col.Add Array(ini, i - 1)
                        ini = i
                        GoTo Siguiente
                    End If
                End If
            End If
        End If

        ' =====================================================
        ' 4. PATRÓN VCV --> V | CV  (SIN excepciones)
        ' =====================================================
        If EsVocal_ES(c1) And EsConsonante_ES(c2) Then
            If i < Len(Texto) Then
                c3 = Mid$(Texto, i + 1, 1)
                If EsVocal_ES(c3) Then
                    col.Add Array(ini, i - 1)
                    ini = i
                    GoTo Siguiente
                End If
            End If
        End If

        ' =====================================================
        ' 5. PATRÓN VCV --> V | CV
        '     Excepción: A-H-U + consonante/Y --> no cortar (AHU...)
        ' =====================================================
        If EsVocal_ES(c1) And EsConsonante_ES(c2) Then
            If i < Len(Texto) Then
                c3 = Mid$(Texto, i + 1, 1)
                If EsVocal_ES(c3) Then

                    ' Excepción: A-H-U seguido de consonante o Y --> AHU en una sola sílaba
                    If c1 = "A" And c2 = "H" And c3 = "U" Then
                        If i + 1 < Len(Texto) Then
                            Dim c4 As String
                            c4 = Mid$(Texto, i + 2, 1)
                            If EsConsonante_ES(c4) Or c4 = "Y" Then
                                GoTo Siguiente
                            End If
                        End If
                    End If

                    ' Caso general VCV --> V | CV
                    col.Add Array(ini, i - 1)
                    ini = i
                    GoTo Siguiente
                End If
            End If
        End If

        ' =====================================================
        ' 6. VOCAL + VOCAL (VV) --> posible hiato
        '     - Si hay vocal débil tildada (Í, Ú) --> hiato seguro
        '     - MARÍA --> MA-RÍ-A
        '     - RAÚL --> RA-ÚL
        ' =====================================================
        If EsVocal_ES(c1) And EsVocal_ES(c2) Then

            ' Hiato por vocal débil tildada
            If c1 = "Í" Or c1 = "Ú" Or c2 = "Í" Or c2 = "Ú" Then
                col.Add Array(ini, i - 1)
                ini = i
                GoTo Siguiente
            End If

        End If

Siguiente:
    Next i

    ' =====================================================
    ' 6. Última sílaba
    ' =====================================================
    If ini <= Len(Texto) Then
        col.Add Array(ini, Len(Texto))
    End If

    Set Silabear_ES = col

End Function




' ================================
' Auxiliares castellanas
' ================================
Private Function EsVocal_ES(c As String) As Boolean
    EsVocal_ES = (InStr("AEIOUÁÉÍÓÚ", c) > 0)
End Function

Private Function EsVocalFuerte_ES(c As String) As Boolean
    EsVocalFuerte_ES = (InStr("AÁEÉOÓ", c) > 0)
End Function

Private Function EsConsonante_ES(c As String) As Boolean
    EsConsonante_ES = (c <> " " And Not EsVocal_ES(c))
End Function

Private Function EsDiptongo_ES(c1 As String, c2 As String) As Boolean
    Dim d As Variant, lista As Variant
    lista = Array( _
        "AI", "EI", "OI", "UI", _
        "AU", "EU", "OU", _
        "IA", "IE", "IO", "IU", _
        "UA", "UE", "UO" _
    )
    For Each d In lista
        If c1 & c2 = d Then
            EsDiptongo_ES = True
            Exit Function
        End If
    Next d
End Function

Private Function EsGrupoInseparable_ES(par As String) As Boolean
    Dim g As Variant, lista As Variant
    lista = Array("BR", "BL", "CR", "CL", "DR", "FR", "GR", "GL", "PR", "PL", "TR")
    For Each g In lista
        If par = g Then
            EsGrupoInseparable_ES = True
            Exit Function
        End If
    Next g
End Function

Public Function MF_DetectarTonicaCastellano( _
        ByVal Texto As String, _
        ByVal Silabas As Collection _
    ) As Long

    Dim i As Long
    Dim vocalesTilde As String
    vocalesTilde = "ÁÉÍÓÚ"

    ' 1. Si hay tilde --> esa sílaba es tónica
    For i = 1 To Len(Texto)
        If InStr(vocalesTilde, Mid$(Texto, i, 1)) > 0 Then
            MF_DetectarTonicaCastellano = MF_SilabaDeIndice(i, Silabas)
            Exit Function
        End If
    Next i

    ' 2. Si no hay tilde --> reglas generales
    Dim ultima As String
    ultima = Right$(Texto, 1)

    If ultima = "N" Or ultima = "S" Or InStr("AEIOU", ultima) > 0 Then
        ' Palabra llana --> penúltima sílaba
        If Silabas.Count >= 2 Then
            MF_DetectarTonicaCastellano = Silabas.Count - 1
        Else
            MF_DetectarTonicaCastellano = Silabas.Count
        End If
    Else
        ' Palabra aguda --> última sílaba
        MF_DetectarTonicaCastellano = Silabas.Count
    End If

End Function

Public Function MF_SilabaDeIndice( _
        ByVal idx As Long, _
        ByVal Silabas As Collection _
    ) As Long

    Dim i As Long
    For i = 1 To Silabas.Count
        If idx >= Silabas(i)(0) And idx <= Silabas(i)(1) Then
            MF_SilabaDeIndice = i
            Exit Function
        End If
    Next i

End Function



' ============================================================
'   ReglasCastellano (ESP)
'   Devuelve idFonema según la fonética del castellano.
'   Si no aplica, devuelve 0 para que el motor siga probando.
' ============================================================

Public Function ReglasCastellano( _
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

    ' GÜE / GÜI --> /gw/ ? id 57
    If g = "GÜE" Or g = "GÜI" Then
        ReglasCastellano = 57
        Exit Function
    End If

    ' GUE / GUI --> /g/ (U muda) ? id 31
    If g = "GUE" Or g = "GUI" Then
        ReglasCastellano = 31
        Exit Function
    End If

    ' QUE / QUI --> /k/ --> id 30
    If g = "QUE" Or g = "QUI" Then
        ReglasCastellano = 30
        Exit Function
    End If


    ' ============================================================
    '   DÍGRAFOS Y CASOS ESPECIALES
    ' ============================================================

    ' CH --> /t?/ --> id 50
    If g = "CH" Then
        ReglasCastellano = 50
        Exit Function
    End If

    ' LL --> /?/ (fonema histórico; hoy yeísmo --> /?/)
    ' Usamos /?/ --> id 44 para mantener coherencia fonética
    If g = "LL" Then
        ReglasCastellano = 44
        Exit Function
    End If

    ' RR --> /r/ múltiple --> id 46
    If g = "RR" Then
        ReglasCastellano = 46
        Exit Function
    End If

    ' Ñ --> /?/ --> id 41
    If g = "Ñ" Then
        ReglasCastellano = 41
        Exit Function
    End If

    ' GU + vocal --> /g/ --> id 31
    If g = "GU" And (sig = "A" Or sig = "O" Or sig = "U") Then
        ReglasCastellano = 31
        Exit Function
    End If

    ' QU + vocal --> /k/ --> id 30
    If g = "QU" And (sig = "A" Or sig = "O" Or sig = "U") Then
        ReglasCastellano = 30
        Exit Function
    End If


    ' ============================================================
    '   DÍGRAFOS VOCÁLICOS (diptongos castellanos)
    ' ============================================================

    If g = "AI" Then ReglasCastellano = 12: Exit Function
    If g = "EI" Then ReglasCastellano = 13: Exit Function
    If g = "OI" Then ReglasCastellano = 14: Exit Function
    If g = "OU" Then ReglasCastellano = 15: Exit Function
    If g = "AU" Then ReglasCastellano = 16: Exit Function


    ' ============================================================
    '   MONÓGRAFOS — VOCALES (5 vocales)
    ' ============================================================

    If g = "A" Then ReglasCastellano = 1: Exit Function
    If g = "E" Then ReglasCastellano = 5: Exit Function
    If g = "I" Then ReglasCastellano = 9: Exit Function
    If g = "O" Then ReglasCastellano = 7: Exit Function
    If g = "U" Then ReglasCastellano = 10: Exit Function


    ' ============================================================
    '   MONÓGRAFOS — CONSONANTES
    ' ============================================================

    If g = "P" Then ReglasCastellano = 26: Exit Function
    If g = "B" Then ReglasCastellano = 27: Exit Function
    If g = "T" Then ReglasCastellano = 28: Exit Function
    If g = "D" Then ReglasCastellano = 29: Exit Function
    If g = "K" Then ReglasCastellano = 30: Exit Function
    If g = "G" Then ReglasCastellano = 31: Exit Function

    If g = "F" Then ReglasCastellano = 32: Exit Function

    ' C/Z --> /?/ (castellano estándar)
    If g = "C" And (sig = "E" Or sig = "I") Then
        ReglasCastellano = 54   ' /?/
        Exit Function
    End If
    If g = "Z" Then
        ReglasCastellano = 54   ' /?/
        Exit Function
    End If

    ' S --> /s/
    If g = "S" Then ReglasCastellano = 34: Exit Function

    ' J / G + E/I --> /x/ --> id 58
    If g = "J" Then ReglasCastellano = 58: Exit Function
    If g = "G" And (sig = "E" Or sig = "I") Then
        ReglasCastellano = 58
        Exit Function
    End If

    ' M / N
    If g = "M" Then ReglasCastellano = 39: Exit Function
    If g = "N" Then ReglasCastellano = 40: Exit Function

    ' L / R simple
    If g = "L" Then ReglasCastellano = 43: Exit Function
    If g = "R" Then ReglasCastellano = 45: Exit Function

    ' H muda --> /h/ glotal suave --> id 38
    If g = "H" Then ReglasCastellano = 38: Exit Function


    ' ============================================================
    '   SI NO APLICA, DEVOLVER 0
    ' ============================================================
    ReglasCastellano = 0

End Function
'============================================================================================

Public Function MF_NormalizarVocales_ES(ByVal Texto As String) As String

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

