Attribute VB_Name = "modMotor_Idioma_ES_2"

' ------------------------------------------------------
' Nombre:    modMotor_Idioma_ES_2
' Tipo:      Módulo
' Propósito: Motor fonético para el Español (es)
' Autor:     Alba Salvá
' Fecha:     02/02/2026
' ------------------------------------------------------

Option Compare Database
Option Explicit

' ============================================================
'   ConvertirTextoAFonemas_ES
'   Conversor fonético por sílaba (KOSMOS 2.1)
'
'   Flujo:
'       1. MF_Silabear_ES --> colFinal + esTonica()
'       2. Recorrer sílabas finales
'       3. Convertir cada sílaba a fonemas
'       4. Insertar separadores entre palabras (0)
'       5. Devolver estructura fonética completa
' ============================================================

Public Function ConvertirTextoAFonemas_ES( _
        ByVal texto As String, _
        ByVal Abreviado As Boolean _
    ) As Collection

    Dim colSilabas As Collection
    Dim resultado As New Collection
    Dim esTonica() As Boolean
    Dim sRevisado As String

    Dim i As Long
    Dim inicio As Long, fin As Long
    Dim silabaTexto As String
    Dim esTon As Boolean
    Dim fonemasSilaba As Collection

    ' --------------------------------------------------------
    ' 1. Silabear + tónica + revisión (función unificada)
    ' --------------------------------------------------------
    If Not MF_Silabear_ES(texto, colSilabas, esTonica, sRevisado) Then
        Set ConvertirTextoAFonemas_ES = resultado
        Exit Function
    End If

    ' --------------------------------------------------------
    ' 2. Recorrer sílabas finales
    ' --------------------------------------------------------
    For i = 1 To colSilabas.Count

        inicio = colSilabas(i)(0)
        fin = colSilabas(i)(1)

        ' Extraer texto de la sílaba
        silabaTexto = Mid$(texto, inicio, fin - inicio + 1)

        ' Determinar si esta sílaba es tónica
        esTon = esTonica(inicio)

        ' ----------------------------------------------------
        ' 2.1 Convertir sílaba a fonemas
        ' ----------------------------------------------------
        Set fonemasSilaba = ConvertirSilabaAFonemas_ES(silabaTexto, esTon)

        ' Añadir al resultado
        resultado.Add fonemasSilaba

        ' ----------------------------------------------------
        ' 2.2 Detectar separador entre palabras
        ' ----------------------------------------------------
        If i < colSilabas.Count Then
            If colSilabas(i + 1)(0) > fin + 1 Then
                resultado.Add 0   ' separador entre palabras
            End If
        End If

    Next i

    Set ConvertirTextoAFonemas_ES = resultado

End Function



' ============================================================
'   MF_Silabear_ES
'   Motor Fonético – Silabeo unificado con tónica y revisión
'
'   Devuelve:
'       - colFinal: colección de sílabas (rangos [ini, fin])
'       - esTonica(): array booleano marcando la sílaba tónica
'       - sFinal: cadena revisada (opcional)
'
'   Flujo:
'       1. Silabear automáticamente
'       2. Detectar tónica (ortográfica)
'       3. Construir cadena editable
'       4. Revisión manual en formulario
'       5. Validación
'       6. Reconstrucción final de rangos
' ============================================================

Private Function MF_Silabear_ES( _
        ByVal texto As String, _
        ByRef colFinal As Collection, _
        ByRef esTonica() As Boolean, _
        Optional ByRef sFinal As String _
    ) As Boolean

    Dim colAuto As Collection
    Dim s As String
    Dim valido As Boolean
    Dim msg As String

    Dim palabras() As String
    Dim sils() As String
    Dim p As Variant, s2 As Variant
    Dim pos As Long
    Dim ini As Long, fin As Long
    
    Dim partesManual() As String
    Dim idxTonica As Long
    Dim iSil As Long
    Dim posSil As Long


    ' ============================================================
    ' 1. Silabear automáticamente
    ' ============================================================
    Set colAuto = Silabear_ES(texto)
    If colAuto Is Nothing Or colAuto.Count = 0 Then
        MF_Silabear_ES = False
        Exit Function
    End If

    ' Redimensionar array de tónica
    ReDim esTonica(1 To Len(texto))

    ' ============================================================
    ' 2. Detectar tónica (ortográfica)
    ' ============================================================
    'Call MF_MarcarTonica_ES(Texto, esTonica)
    Call MF_MarcarTonica_ES(texto, esTonica, colAuto)

    ' ============================================================
    ' 3. Construir cadena editable (guiones y espacios)
    ' ============================================================
    s = ""
    Dim i As Long
    For i = 1 To colAuto.Count

        ' Añadir sílaba
        s = s & Mid$(texto, colAuto(i)(0), colAuto(i)(1) - colAuto(i)(0) + 1)

        ' Separador
        If i < colAuto.Count Then
            If colAuto(i + 1)(0) > colAuto(i)(1) + 1 Then
                s = s & " "   ' espacio real entre palabras
            Else
                s = s & "-"   ' separador silábico
            End If
        End If

    Next i

    ' ============================================================
    ' 4. Revisión manual en formulario
    ' ============================================================
    s = RevisarSilabas_EnFormulario(texto, s)

    ' Si el usuario cancela ? devolver silabeo automático
    If s = "" Then
        Set colFinal = colAuto
        sFinal = ""
        MF_Silabear_ES = True
        Exit Function
    End If

' ============================================================
' 4.b Detectar tónica manual mediante símbolo "*"
' ============================================================

idxTonica = 0

' Dividir en sílabas ignorando espacios
partesManual = Split(Replace(s, " ", ""), "-")

' Buscar sílaba con "*"
For iSil = LBound(partesManual) To UBound(partesManual)
    If Left$(partesManual(iSil), 1) = "*" Then
        idxTonica = iSil + 1                ' índice 1-based
        partesManual(iSil) = Mid$(partesManual(iSil), 2) ' quitar "*"
        Exit For
    End If
Next iSil

' Si hay tónica manual, sobrescribir esTonica()
If idxTonica > 0 Then

    ' Limpiar array de tónica
    Dim k As Long
    For k = LBound(esTonica) To UBound(esTonica)
        esTonica(k) = False
    Next k

    ' Calcular posición exacta en el texto original
    posSil = 1
    For iSil = 1 To idxTonica - 1
        posSil = posSil + Len(partesManual(iSil - 1))
    Next iSil

    ' Marcar la sílaba tónica en esTonica()
    For k = posSil To posSil + Len(partesManual(idxTonica - 1)) - 1
        esTonica(k) = True
    Next k

    ' Limpiar "*" en la cadena s
    Dim tmp() As String
    tmp = Split(s, "-")

    For iSil = LBound(tmp) To UBound(tmp)
        If Left$(tmp(iSil), 1) = "*" Then
            tmp(iSil) = Mid$(tmp(iSil), 2)
        End If
    Next iSil

    s = Join(tmp, "-")
End If

    ' ============================================================
    ' 5. Validación
    ' ============================================================
    Do
        valido = True
        msg = ""

        ' No puede empezar ni acabar con "-"
        If Left$(s, 1) = "-" Or Right$(s, 1) = "-" Then
            valido = False
            msg = "No puede empezar ni terminar con '-'."
        End If

        ' No puede contener "--"
        If InStr(s, "--") > 0 Then
            valido = False
            msg = "No puede haber sílabas vacías ('--')."
        End If

        ' Comprobar reconstrucción del texto
        Dim reconstruido As String
        Dim textoSinEspacios As String

        reconstruido = Replace(Replace(s, "-", ""), " ", "")
        textoSinEspacios = Replace(texto, " ", "")

        If UCase$(reconstruido) <> UCase$(textoSinEspacios) Then
            valido = False
            msg = "Las sílabas no coinciden con el texto original."
        End If

        If Not valido Then
            MsgBox msg, vbExclamation, "Error en las sílabas"
            s = RevisarSilabas_EnFormulario(texto, s)
            If s = "" Then
                Set colFinal = colAuto
                sFinal = ""
                MF_Silabear_ES = True
                Exit Function
            End If
        End If

    Loop Until valido

    ' ============================================================
    ' 6. Reconstrucción final de rangos (respetando espacios)
    ' ============================================================
    Set colFinal = New Collection
    pos = 1

    palabras = Split(s, " ")

    For Each p In palabras
        sils = Split(p, "-")
        For Each s2 In sils
            If Trim$(s2) <> "" Then
                ' Búsqueda posicional exacta
                ini = pos
                Do While ini <= Len(texto) - Len(s2) + 1
                    If Mid$(texto, ini, Len(s2)) = s2 Then Exit Do
                    ini = ini + 1
                Loop

                If ini > Len(texto) - Len(s2) + 1 Then
                    MsgBox "Error: la sílaba '" & s2 & "' no se encuentra en el texto original.", vbCritical
                    MF_Silabear_ES = False
                    Exit Function
                End If

                fin = ini + Len(s2) - 1
                colFinal.Add Array(ini, fin)

                pos = fin + 1
            End If
        Next s2
        ' Avanzar espacio si existe
        If pos <= Len(texto) Then
            If Mid$(texto, pos, 1) = " " Then pos = pos + 1
        End If
    Next p

    ' ============================================================
    ' 7. Devolver cadena final revisada
    ' ============================================================
    sFinal = s
    MF_Silabear_ES = True

End Function

' ============================================================
'   Silabear_ES — Silabeador para nombres y apellidos en español
' ============================================================

Private Function Silabear_ES(ByVal texto As String) As Collection

    Dim col As New Collection
    Dim i As Long, ini As Long
    Dim c1 As String, c2 As String, c3 As String
    Dim par As String

    texto = Trim$(texto)
    If Len(texto) = 0 Then
        Set Silabear_ES = col
        Exit Function
    End If

    ini = 1

    For i = 2 To Len(texto)

        c1 = Mid$(texto, i - 1, 1)
        c2 = Mid$(texto, i, 1)
        par = c1 & c2

        ' =====================================================
        ' 0. ESPACIOS
        ' =====================================================
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

        ' =====================================================
        ' 1. DÍGRAFOS INSEPARABLES (CH, LL, RR)
        ' =====================================================
        If par = "CH" Or par = "LL" Or par = "RR" Then
            GoTo siguiente
        End If

        ' =====================================================
        ' 2. GRUPOS CONSONÁNTICOS INSEPARABLES (BR, CR, TR…)
        ' =====================================================
        If EsConsonante_ES(c1) And EsConsonante_ES(c2) Then
            If EsGrupoInseparable_ES(par) Then
                GoTo siguiente
            End If
        End If

        ' =====================================================
        ' 3. PATRÓN CCV --> C | CV
        ' =====================================================
        If EsConsonante_ES(c1) And EsConsonante_ES(c2) Then
            If i < Len(texto) Then
                c3 = Mid$(texto, i + 1, 1)
                If EsVocal_ES(c3) Then
                    If Not EsGrupoInseparable_ES(par) Then
                        col.Add Array(ini, i - 1)
                        ini = i
                        GoTo siguiente
                    End If
                End If
            End If
        End If

        ' =====================================================
        ' 4. PATRÓN VCV --> V | CV  (SIN excepciones)
        ' =====================================================
        If EsVocal_ES(c1) And EsConsonante_ES(c2) Then
            If i < Len(texto) Then
                c3 = Mid$(texto, i + 1, 1)
                If EsVocal_ES(c3) Then
                    col.Add Array(ini, i - 1)
                    ini = i
                    GoTo siguiente
                End If
            End If
        End If

        ' =====================================================
        ' 5. PATRÓN VCV --> V | CV
        '     Excepción: A-H-U + consonante/Y --> no cortar (AHU...)
        ' =====================================================
        If EsVocal_ES(c1) And EsConsonante_ES(c2) Then
            If i < Len(texto) Then
                c3 = Mid$(texto, i + 1, 1)
                If EsVocal_ES(c3) Then

                    ' Excepción: A-H-U seguido de consonante o Y --> AHU en una sola sílaba
                    If c1 = "A" And c2 = "H" And c3 = "U" Then
                        If i + 1 < Len(texto) Then
                            Dim c4 As String
                            c4 = Mid$(texto, i + 2, 1)
                            If EsConsonante_ES(c4) Or c4 = "Y" Then
                                GoTo siguiente
                            End If
                        End If
                    End If

                    ' Caso general VCV --> V | CV
                    col.Add Array(ini, i - 1)
                    ini = i
                    GoTo siguiente
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
                GoTo siguiente
            End If

        End If

siguiente:
    Next i

    ' =====================================================
    ' 6. Última sílaba
    ' =====================================================
    If ini <= Len(texto) Then
        col.Add Array(ini, Len(texto))
    End If

    Set Silabear_ES = col

End Function

Private Sub MF_MarcarTonica_ES( _
        ByVal texto As String, _
        ByRef esTonica() As Boolean, _
        ByVal Silabas As Collection _
    )

    Dim i As Long
    Dim idxTonica As Long
    Dim inicio As Long, fin As Long
    Dim yaMarcada As Boolean

    If Silabas Is Nothing Or Silabas.Count = 0 Then Exit Sub

    ' ============================================================
    ' 1. Comprobar si ya existe una tónica marcada (manual con "*")
    ' ============================================================
    yaMarcada = False
    For i = LBound(esTonica) To UBound(esTonica)
        If esTonica(i) = True Then
            yaMarcada = True
            Exit For
        End If
    Next i

    ' Si ya hay tónica manual, no hacemos nada
    If yaMarcada Then Exit Sub

    ' ============================================================
    ' 2. Detectar tónica ortográfica (RAE)
    ' ============================================================
    idxTonica = MF_DetectarTonicaES(texto, Silabas)
    If idxTonica = 0 Then Exit Sub

    ' ============================================================
    ' 3. Marcar índices de la sílaba tónica
    ' ============================================================
    inicio = Silabas(idxTonica)(0)
    fin = Silabas(idxTonica)(1)

    For i = inicio To fin
        esTonica(i) = True
    Next i

End Sub

Private Function MF_DetectarTonicaES( _
        ByVal texto As String, _
        ByVal Silabas As Collection _
    ) As Long

    Dim i As Long
    Dim vocalesTilde As String
    vocalesTilde = "ÁÉÍÓÚáéíóú"

    ' ============================================================
    ' 1. Buscar vocal con tilde en el texto original
    ' ============================================================
    For i = 1 To Len(texto)
        If InStr(vocalesTilde, Mid$(texto, i, 1)) > 0 Then
            MF_DetectarTonicaES = MF_SilabaDeIndice(i, Silabas)
            Exit Function
        End If
    Next i

    ' ============================================================
    ' 2. No hay tilde ? aplicar reglas generales RAE
    ' ============================================================
    Dim ultima As String
    ultima = Right$(Trim$(texto), 1)

    ' Normalizar a mayúsculas para comparar
    ultima = UCase$(ultima)

    If ultima = "N" Or ultima = "S" Or InStr("AEIOU", ultima) > 0 Then
        ' Llana ? penúltima sílaba
        If Silabas.Count >= 2 Then
            MF_DetectarTonicaES = Silabas.Count - 1
        Else
            MF_DetectarTonicaES = Silabas.Count
        End If
    Else
        ' Aguda ? última sílaba
        MF_DetectarTonicaES = Silabas.Count
    End If

End Function

' ============================================================
'   ConvertirSilabaAFonemas_ES
'   Convierte una sílaba ortográfica en una secuencia de fonemas
'   usando las reglas del castellano.
' ============================================================

Public Function ConvertirSilabaAFonemas_ES( _
        ByVal silabaTexto As String, _
        ByVal esTonica As Boolean _
    ) As Collection

    Dim Fonemas As New Collection
    Dim i As Long
    Dim graf As String
    Dim ant As String, sig As String
    Dim idF As Byte

    ' Normalizar vocales
    silabaTexto = MF_NormalizarVocales_ES(silabaTexto)

    ' Insertar acento universal si la sílaba es tónica
    If esTonica Then Fonemas.Add 61

    i = 1
    Do While i <= Len(silabaTexto)

        ' Contexto
        If i > 1 Then ant = Mid$(silabaTexto, i - 1, 1) Else ant = ""
        If i < Len(silabaTexto) Then sig = Mid$(silabaTexto, i + 1, 1) Else sig = ""

        ' Intentar trigrafema
        If i <= Len(silabaTexto) - 2 Then
            graf = Mid$(silabaTexto, i, 3)
            idF = ReglasCastellano(graf, ant, sig, esTonica)
            If idF <> 0 Then
                Fonemas.Add idF
                i = i + 3
                GoTo siguiente
            End If
        End If

        ' Intentar dígrafo
        If i <= Len(silabaTexto) - 1 Then
            graf = Mid$(silabaTexto, i, 2)
            idF = ReglasCastellano(graf, ant, sig, esTonica)
            If idF <> 0 Then
                Fonemas.Add idF
                i = i + 2
                GoTo siguiente
            End If
        End If

        ' Monógrafo
        graf = Mid$(silabaTexto, i, 1)
        idF = ReglasCastellano(graf, ant, sig, esTonica)
        If idF <> 0 Then Fonemas.Add idF

        i = i + 1

siguiente:
    Loop

    Set ConvertirSilabaAFonemas_ES = Fonemas

End Function

' ============================================================
'   ReglasCastellano (ESP)
'   Devuelve idFonema según la fonética del castellano.
'   Si no aplica, devuelve 0 para que el motor siga probando.
' ============================================================

Private Function ReglasCastellano( _
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
    If g = "B" Or _
       g = "V" Then ReglasCastellano = 27: Exit Function
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

' ================================
' Auxiliares ES
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

Private Function EsVocalDebilTilde(c As String) As Boolean
    EsVocalDebilTilde = (c = "Í" Or c = "Ú")
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

Private Function MF_SilabaDeIndice( _
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
    texto = Replace(texto, "Ü", "Ü") 'no tocar

    MF_NormalizarVocales_ES = texto

End Function


''===============================
''  MF_MarcarTonica_ES (CORREGIDA)
''===============================
'Private Sub MF_MarcarTonica_ES( _
'        ByVal Texto As String, _
'        ByRef esTonica() As Boolean _
'    )
'
'    Dim i As Long
'    Dim Silabas As Collection
'    Dim idxTonica As Long
'    Dim inicio As Long, fin As Long
'
'    ' --------------------------------------------------------
'    ' 1. Silabear palabra (AUTOMÁTICO, sin revisión)
'    '    La tónica se determina ortográficamente,
'    '    por tanto NO debe depender de la revisión manual.
'    ' --------------------------------------------------------
'    Set Silabas = Silabear_ES(Texto)
'
'    If Silabas.Count = 0 Then Exit Sub
'
'    ' --------------------------------------------------------
'    ' 2. Determinar sílaba tónica
'    ' --------------------------------------------------------
'    idxTonica = MF_DetectarTonicaCastellano(Texto, Silabas)
'
'    If idxTonica = 0 Then Exit Sub
'
'    ' --------------------------------------------------------
'    ' 3. Marcar índices de la sílaba tónica
'    ' --------------------------------------------------------
'    inicio = Silabas(idxTonica)(0)
'    fin = Silabas(idxTonica)(1)
'
'    For i = inicio To fin
'        esTonica(i) = True
'    Next i
'
'End Sub


'Private Sub MF_MarcarTonica_ES( _
'        ByVal Texto As String, _
'        ByRef esTonica() As Boolean, _
'        ByVal Silabas As Collection _
'    )
'
'    Dim i As Long
'    Dim idxTonica As Long
'    Dim inicio As Long, fin As Long
'
'    If Silabas Is Nothing Or Silabas.Count = 0 Then Exit Sub
'
'    ' Determinar sílaba tónica
'    idxTonica = MF_DetectarTonicaES(Texto, Silabas)
'    If idxTonica = 0 Then Exit Sub
'
'    ' Marcar índices de la sílaba tónica
'    inicio = Silabas(idxTonica)(0)
'    fin = Silabas(idxTonica)(1)
'
'    For i = inicio To fin
'        esTonica(i) = True
'    Next i
'
'End Sub

'Private Function MF_DetectarTonicaES( _
'        ByVal Texto As String, _
'        ByVal Silabas As Collection _
'    ) As Long
'
'    Dim i As Long
'    Dim vocalesTilde As String
'    vocalesTilde = "ÁÉÍÓÚ"
'
'    ' 1. Si hay tilde --> esa sílaba es tónica
'    For i = 1 To Len(Texto)
'        If InStr(vocalesTilde, Mid$(Texto, i, 1)) > 0 Then
'            MF_DetectarTonicaES = MF_SilabaDeIndice(i, Silabas)
'            Exit Function
'        End If
'    Next i
'
'    ' 2. Si no hay tilde --> reglas generales
'    Dim ultima As String
'    ultima = Right$(Texto, 1)
'
'    If ultima = "N" Or ultima = "S" Or InStr("AEIOU", ultima) > 0 Then
'        ' Palabra llana --> penúltima sílaba
'        If Silabas.Count >= 2 Then
'            MF_DetectarTonicaES = Silabas.Count - 1
'        Else
'            MF_DetectarTonicaES = Silabas.Count
'        End If
'    Else
'        ' Palabra aguda --> última sílaba
'        MF_DetectarTonicaES = Silabas.Count
'    End If
'
'End Function

