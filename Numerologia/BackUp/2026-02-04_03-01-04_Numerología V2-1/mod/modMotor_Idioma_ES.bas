Attribute VB_Name = "modMotor_Idioma_ES"

' modMotor_Idioma_ES
Option Compare Database
Option Explicit

'=================
'==  Español ==
'=================

Private IndiceSilabaActual As Long
Private EsTonicaActual As Boolean
Private GrafAnterior As String
Private GrafActual As String
Private GrafSiguiente As String

Private objDTO As clsMotorFonetico

Public Function EntradaMotor_ES(texto As String) As String

    Set objDTO = New clsMotorFonetico
    objDTO.TextoOriginal = texto

    Call NormalizarConReglas
    Call SilabearAuto
    Call DetectarTonicaGeneral
    Call RevisionSilabeo
    Call ReconstruirSilabasFinales
    Call MarcarTonicaEnSilaba
    Call ConvertirSilabasAFonemas
    'Call DescodificarResultado

    EntradaMotor_ES = objDTO.TextoFinal

End Function

'=======================================================================================

Private Sub NormalizarConReglas()

    Dim s As String

    ' 1. Partimos del texto original
    s = objDTO.TextoOriginal

    ' 2. Eliminar espacios duplicados
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop

    ' 3. Eliminar caracteres invisibles o de control
    s = Replace(s, vbTab, " ")
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")

    ' 4. Normalizar comillas y guiones
    s = Replace(s, "–", "-")   ' guion largo ? guion normal
    s = Replace(s, "—", "-")   ' em dash ? guion normal
    s = Replace(s, "“", """")
    s = Replace(s, "”", """")

    ' 5. Eliminar espacios antes/después de guiones
    s = Replace(s, " -", "-")
    s = Replace(s, "- ", "-")

    ' 6. Trim final conn conversión a minúsculas
    s = LCase$(Trim$(s))

    ' 7. Guardar en el DTO
    objDTO.TextoNormalizado = s

End Sub

Private Sub SilabearAuto()

    Dim texto As String
    Dim col As New Collection
    Dim i As Long, ini As Long
    Dim c1 As String, c2 As String, c3 As String, c4 As String
    Dim par As String

    texto = objDTO.TextoNormalizado
    texto = Trim$(texto)

    If Len(texto) = 0 Then
        ReDim objDTO.SilabasAuto(0)
        Exit Sub
    End If

    ini = 1

    For i = 2 To Len(texto)

        c1 = Mid$(texto, i - 1, 1)
        c2 = Mid$(texto, i, 1)
        par = c1 & c2

        ' ============================================================
        ' 0. ESPACIOS
        ' ============================================================
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

        ' ============================================================
        ' 1. DÍGRAFOS INSEPARABLES
        ' ============================================================
        If par = "ch" Or par = "ll" Or par = "rr" Then GoTo siguiente

        ' ============================================================
        ' 2. GRUPOS CONSONÁNTICOS INSEPARABLES
        ' ============================================================
        If EsConsonante_ES(c1) And EsConsonante_ES(c2) Then
            If EsGrupoInseparable_ES(par) Then GoTo siguiente
        End If

        ' ============================================================
        ' 3. REGLAS VOCÁLICAS (VV / VVV)
        ' ============================================================

        ' --- TRIPTONGO: débil + fuerte + débil ---
        If i < Len(texto) - 1 Then
            c3 = Mid$(texto, i + 1, 1)
            If EsTriptongo(c1, c2, c3) Then
                GoTo siguiente
            End If
        End If

        ' --- HIATO POR TILDE: í / ú ---
        If (c1 = "í" Or c1 = "ú") Or (c2 = "í" Or c2 = "ú") Then
            col.Add Array(ini, i - 1)
            ini = i
            GoTo siguiente
        End If

        ' --- HIATO FUERTE + FUERTE ---
        If EsHiatoFuerteFuerte(c1, c2) Then
            col.Add Array(ini, i - 1)
            ini = i
            GoTo siguiente
        End If

        ' --- DIPTONGO ---
        If EsDiptongo(c1, c2) Then
            GoTo siguiente
        End If

        ' --- VV ? separar ---
        If EsVocal_ES(c1) And EsVocal_ES(c2) Then
            col.Add Array(ini, i - 1)
            ini = i
            GoTo siguiente
        End If

        ' ============================================================
        ' 4. CCV ? C | CV
        ' ============================================================
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

        ' ============================================================
        ' 5. VCV ? V | CV
        ' ============================================================
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

        ' ============================================================
        ' 6. EXCEPCIÓN A-H-U
        ' ============================================================
        If EsVocal_ES(c1) And EsConsonante_ES(c2) Then
            If i < Len(texto) Then
                c3 = Mid$(texto, i + 1, 1)
                If EsVocal_ES(c3) Then
                    If c1 = "a" And c2 = "h" And c3 = "u" Then
                        If i + 1 < Len(texto) Then
                            c4 = Mid$(texto, i + 2, 1)
                            If EsConsonante_ES(c4) Or c4 = "y" Then GoTo siguiente
                        End If
                    End If
                    col.Add Array(ini, i - 1)
                    ini = i
                    GoTo siguiente
                End If
            End If
        End If

siguiente:
    Next i

    ' Última sílaba
    If ini <= Len(texto) Then
        col.Add Array(ini, Len(texto))
    End If

    ' Convertir rangos ? array de sílabas
    ReDim objDTO.SilabasAuto(1 To col.Count)

    For i = 1 To col.Count
        objDTO.SilabasAuto(i) = Mid$(texto, col(i)(0), col(i)(1) - col(i)(0) + 1)
    Next i

End Sub

Private Sub DetectarTonicaGeneral()

    Dim i As Long, j As Long
    Dim texto As String
    Dim sils() As String
    Dim vocalesTilde As String
    Dim idx As Long

    texto = objDTO.TextoNormalizado          ' ya está en minúsculas
    sils = objDTO.SilabasAuto
    'vocalesTilde = "ÁÉÍÓÚáéíóú"             ' solo usará las minúsculas
    vocalesTilde = "áéíóú"                   ' solo minúsculas

    ' ============================================================
    ' 1. Buscar vocal con tilde en las sílabas automáticas
    ' ============================================================
    For i = 1 To UBound(sils)
        For j = 1 To Len(sils(i))
            If InStr(vocalesTilde, Mid$(sils(i), j, 1)) > 0 Then
                idx = i
                Exit For
            End If
        Next j
        If idx > 0 Then Exit For
    Next i

    ' ============================================================
    ' 2. Si no hay tilde ? aplicar reglas generales RAE
    ' ============================================================
    If idx = 0 Then
        Dim ultima As String
        ultima = Right$(Trim$(texto), 1)     ' ya está en minúsculas

        ' Palabras llanas si terminan en vocal, n o s
        If ultima = "n" Or ultima = "s" Or InStr("aeiou", ultima) > 0 Then
            If UBound(sils) >= 2 Then
                idx = UBound(sils) - 1       ' penúltima
            Else
                idx = UBound(sils)           ' única sílaba
            End If
        Else
            idx = UBound(sils)               ' aguda ? última
        End If
    End If

    ' ============================================================
    ' 3. Guardar resultado en el DTO
    ' ============================================================
    ReDim objDTO.SilabaTonica(1 To 1)
    objDTO.SilabaTonica(1) = idx

End Sub

Private Sub RevisionSilabeo()

    Dim s As String
    Dim partes() As String
    Dim resultado() As String
    Dim tFinal() As Byte
    Dim texto As String
    Dim i As Long
    Dim idxTonica As Collection
    Dim raw As String
    Dim limpio As String
    Dim p1 As Long, p2 As Long

    texto = objDTO.TextoNormalizado

    ' ============================================================
    ' 1. Construir cadena editable a partir de SilabasAuto()
    ' ============================================================
    s = ""
    For i = 1 To UBound(objDTO.SilabasAuto)
        s = s & objDTO.SilabasAuto(i)
        If i < UBound(objDTO.SilabasAuto) Then s = s & "-"
    Next i

    ' ============================================================
    ' 2. Mostrar formulario de revisión
    ' ============================================================
    s = RevisarSilabas_EnFormulario(texto, s)

    ' Si el usuario cancela ? mantener silabeo automático
    If s = "" Then
        objDTO.SilabasFinal = objDTO.SilabasAuto
        objDTO.SilabaTonica = objDTO.SilabaTonica
        Exit Sub
    End If

    ' ============================================================
    ' 3. Dividir por sílabas (sin eliminar espacios)
    ' ============================================================
    partes = Split(s, "-")

    Set idxTonica = New Collection

    ' ============================================================
    ' 4. Procesar cada sílaba
    ' ============================================================
    For i = LBound(partes) To UBound(partes)

        raw = partes(i)

        ' Detectar sílaba espacio " "
        If raw = " " Then
            partes(i) = " "
            GoTo siguiente
        End If

        ' Detectar delimitadores tónica "*…*"
        If InStr(raw, "*") > 0 Then

            ' Localizar primer y último asterisco
            p1 = InStr(1, raw, "*")
            p2 = InStrRev(raw, "*")

            ' Registrar índice de sílaba tónica (1-based)
            idxTonica.Add i + 1

            ' Extraer contenido interno
            limpio = Mid$(raw, p1 + 1, p2 - p1 - 1)

            ' Asignar sílaba limpia
            partes(i) = limpio

        Else
            ' Sílaba normal
            partes(i) = raw
        End If

siguiente:
    Next i

    ' ============================================================
    ' 5. Reconstruir SilabasFinal()
    ' ============================================================
    ReDim resultado(1 To UBound(partes) + 1)

    For i = 1 To UBound(resultado)
        resultado(i) = partes(i - 1)
    Next i

    objDTO.SilabasFinal = resultado

    ' ============================================================
    ' 6. Reconstruir SilabaTonica()
    ' ============================================================
    If idxTonica.Count > 0 Then
        ReDim tFinal(1 To idxTonica.Count)
        For i = 1 To idxTonica.Count
            tFinal(i) = idxTonica(i)
        Next i
        objDTO.SilabaTonica = tFinal
    End If

End Sub

Private Sub MarcarTonicaEnSilaba()

    Dim i As Long, j As Long
    Dim resultado As String
    Dim esTonica As Boolean

    If UBound(objDTO.SilabasFinal) < 1 Then Exit Sub
    If UBound(objDTO.SilabaTonica) < 1 Then Exit Sub

    resultado = ""

    For i = 1 To UBound(objDTO.SilabasFinal)

        esTonica = False
        For j = 1 To UBound(objDTO.SilabaTonica)
            If objDTO.SilabaTonica(j) = i Then
                esTonica = True
                Exit For
            End If
        Next j

        If esTonica Then
            resultado = resultado & "[" & objDTO.SilabasFinal(i) & "]"
        Else
            resultado = resultado & objDTO.SilabasFinal(i)
        End If

        If i < UBound(objDTO.SilabasFinal) Then
            resultado = resultado & "-"
        End If
    Next i

    objDTO.TextoFinal = resultado

End Sub

' ============================================================
'  MOTOR FONÉTICO ESPAÑOL (DTO-CÉNTRICO, SIN PARÁMETROS)
' ============================================================

Private Sub ConvertirSilabasAFonemas()

    Dim i As Long, j As Long
    Dim strFinal As String
    Dim esTonica As Boolean
    Dim sil As String
    Dim arrFon() As Long
    Dim f As Long

    If UBound(objDTO.SilabasFinal) < 1 Then Exit Sub

    strFinal = ""

    For i = 1 To UBound(objDTO.SilabasFinal)

        sil = objDTO.SilabasFinal(i)

        ' ¿Es sílaba tónica?
        esTonica = False
        If UBound(objDTO.SilabaTonica) >= 1 Then
            For j = 1 To UBound(objDTO.SilabaTonica)
                If objDTO.SilabaTonica(j) = i Then
                    esTonica = True
                    Exit For
                End If
            Next j
        End If

        objDTO.IndiceSilabaActual = i
        objDTO.EsTonicaActual = esTonica

        ' Marca de acento (61) si es tónica
        If esTonica Then
            strFinal = strFinal & "61, "
        End If

        ' Espacio como sílaba
        If sil = " " Then
            strFinal = strFinal & "0 - "
            GoTo siguiente
        End If

        ' Fonemas de la sílaba actual
        arrFon = ConvertirGrafemasDeSilabaAIdFonemas()

        For Each f In arrFon
            strFinal = strFinal & CStr(f) & ", "
        Next f

        ' Separador de sílaba
        strFinal = strFinal & "- "

siguiente:
    Next i

    objDTO.TextoFinal = Trim$(strFinal)

End Sub


Private Function ConvertirGrafemasDeSilabaAIdFonemas() As Long()

    Dim sil As String
    Dim s As String
    Dim i As Long
    Dim graf As String
    Dim fon As Byte
    Dim arr() As Long
    Dim idx As Long

    sil = objDTO.SilabasFinal(objDTO.IndiceSilabaActual)
    s = LCase$(sil)

    ReDim arr(1 To 1)
    idx = 1
    i = 1

    Do While i <= Len(s)

        objDTO.GrafAnterior = ""
        objDTO.GrafActual = ""
        objDTO.GrafSiguiente = ""

        ' ---------------- TRIGRAFEMAS ----------------
        If i <= Len(s) - 2 Then
            graf = Mid$(s, i, 3)
            If graf = "güe" Or graf = "güi" Or _
               graf = "gue" Or graf = "gui" Or _
               graf = "que" Or graf = "qui" Then

                If i > 1 Then objDTO.GrafAnterior = Mid$(s, i - 1, 1)
                If i < Len(s) - 2 Then objDTO.GrafSiguiente = Mid$(s, i + 3, 1)
                objDTO.GrafActual = graf

                fon = ReglasCastellano()

                If fon > 0 Then
                    arr(idx) = fon
                    idx = idx + 1
                    ReDim Preserve arr(1 To idx)
                    i = i + 3
                    GoTo siguiente
                End If
            End If
        End If

        ' ---------------- DÍGRAFOS ----------------
        If i <= Len(s) - 1 Then
            graf = Mid$(s, i, 2)
            If graf = "ch" Or graf = "ll" Or graf = "rr" Or _
               graf = "gu" Or graf = "qu" Or _
               graf = "ai" Or graf = "ei" Or graf = "oi" Or graf = "ou" Or graf = "au" Then

                If i > 1 Then objDTO.GrafAnterior = Mid$(s, i - 1, 1)
                If i < Len(s) - 1 Then objDTO.GrafSiguiente = Mid$(s, i + 2, 1)
                objDTO.GrafActual = graf

                fon = ReglasCastellano()

                If fon > 0 Then
                    arr(idx) = fon
                    idx = idx + 1
                    ReDim Preserve arr(1 To idx)
                    i = i + 2
                    GoTo siguiente
                End If
            End If
        End If

        ' ---------------- MONÓGRAFOS ----------------
        graf = Mid$(s, i, 1)
        objDTO.GrafActual = graf
        If i > 1 Then objDTO.GrafAnterior = Mid$(s, i - 1, 1)
        If i < Len(s) Then objDTO.GrafSiguiente = Mid$(s, i + 1, 1)

        fon = ReglasCastellano()

        If fon > 0 Then
            arr(idx) = fon
            idx = idx + 1
            ReDim Preserve arr(1 To idx)
        End If

        i = i + 1

siguiente:
    Loop

    If idx > 1 Then
        ReDim Preserve arr(1 To idx - 1)
    Else
        ReDim arr(1 To 0)
    End If

    ConvertirGrafemasDeSilabaAIdFonemas = arr

End Function


Private Function ReglasCastellano() As Byte

    Dim g As String
    Dim ant As String
    Dim sig As String
    Dim esTonica As Boolean

    g = LCase$(objDTO.GrafActual)
    ant = LCase$(objDTO.GrafAnterior)
    sig = LCase$(objDTO.GrafSiguiente)
    esTonica = objDTO.EsTonicaActual   ' ahora mismo no se usa, pero queda disponible

    ' ================= TRIGRAFEMAS =================

    If g = "güe" Or g = "güi" Then
        ReglasCastellano = 57: Exit Function
    End If

    If g = "gue" Or g = "gui" Then
        ReglasCastellano = 31: Exit Function
    End If

    If g = "que" Or g = "qui" Then
        ReglasCastellano = 30: Exit Function
    End If

    ' ============ DÍGRAFOS Y ESPECIALES ============

    If g = "ch" Then ReglasCastellano = 50: Exit Function
    If g = "ll" Then ReglasCastellano = 44: Exit Function
    If g = "rr" Then ReglasCastellano = 46: Exit Function
    If g = "ñ" Then ReglasCastellano = 41: Exit Function

    If g = "gu" And (sig = "a" Or sig = "o" Or sig = "u") Then
        ReglasCastellano = 31: Exit Function
    End If

    If g = "qu" And (sig = "a" Or sig = "o" Or sig = "u") Then
        ReglasCastellano = 30: Exit Function
    End If

    ' ============ DÍGRAFOS VOCÁLICOS ============

    If g = "ai" Then ReglasCastellano = 12: Exit Function
    If g = "ei" Then ReglasCastellano = 13: Exit Function
    If g = "oi" Then ReglasCastellano = 14: Exit Function
    If g = "ou" Then ReglasCastellano = 15: Exit Function
    If g = "au" Then ReglasCastellano = 16: Exit Function

    ' ============ VOCALES ============

    If g = "a" Then ReglasCastellano = 1: Exit Function
    If g = "e" Then ReglasCastellano = 5: Exit Function
    If g = "i" Then ReglasCastellano = 9: Exit Function
    If g = "o" Then ReglasCastellano = 7: Exit Function
    If g = "u" Then ReglasCastellano = 10: Exit Function

    ' ============ CONSONANTES ============

    If g = "p" Then ReglasCastellano = 26: Exit Function
    If g = "b" Or g = "v" Then ReglasCastellano = 27: Exit Function
    If g = "t" Then ReglasCastellano = 28: Exit Function
    If g = "d" Then ReglasCastellano = 29: Exit Function
    If g = "k" Then ReglasCastellano = 30: Exit Function
    If g = "g" Then ReglasCastellano = 31: Exit Function

    If g = "f" Then ReglasCastellano = 32: Exit Function

    If g = "c" And (sig = "e" Or sig = "i") Then
        ReglasCastellano = 54: Exit Function
    End If
    If g = "z" Then
        ReglasCastellano = 54: Exit Function
    End If

    If g = "s" Then ReglasCastellano = 34: Exit Function

    If g = "j" Then ReglasCastellano = 58: Exit Function
    If g = "g" And (sig = "e" Or sig = "i") Then
        ReglasCastellano = 58: Exit Function
    End If

    If g = "m" Then ReglasCastellano = 39: Exit Function
    If g = "n" Then ReglasCastellano = 40: Exit Function

    If g = "l" Then ReglasCastellano = 43: Exit Function
    If g = "r" Then ReglasCastellano = 45: Exit Function

    If g = "h" Then ReglasCastellano = 38: Exit Function

    ReglasCastellano = 0

End Function



'==================================================================

' ================================
' Auxiliares ES (versión minúsculas)
' ================================

Private Function EsVocal_ES(c As String) As Boolean
    ' Vocales fuertes y débiles, con y sin tilde
    EsVocal_ES = (InStr("aeiouáéíóú", c) > 0)
End Function

Private Function EsConsonante_ES(c As String) As Boolean
    ' Cualquier carácter que no sea vocal ni espacio
    EsConsonante_ES = (c <> " " And Not EsVocal_ES(c))
End Function


Private Function EsVocalFuerte(c As String) As Boolean
    EsVocalFuerte = (InStr("aeoáéó", c) > 0)
End Function

Private Function EsVocalDebil(c As String) As Boolean
    EsVocalDebil = (InStr("iuíú", c) > 0)
End Function

Private Function EsDiptongo(c1 As String, c2 As String) As Boolean
    ' Diptongo si:
    ' - débil + fuerte (sin tilde en la débil)
    ' - fuerte + débil (sin tilde en la débil)
    ' - débil + débil (sin tilde)
    
    If EsVocalDebil(c1) And EsVocalDebil(c2) Then
        ' ii, iu, ui, uu ? diptongo si no hay tilde
        If c1 <> "í" And c1 <> "ú" And c2 <> "í" And c2 <> "ú" Then
            EsDiptongo = True
            Exit Function
        End If
    End If

    ' débil + fuerte
    If EsVocalDebil(c1) And EsVocalFuerte(c2) Then
        If c1 <> "í" And c1 <> "ú" Then
            EsDiptongo = True
            Exit Function
        End If
    End If

    ' fuerte + débil
    If EsVocalFuerte(c1) And EsVocalDebil(c2) Then
        If c2 <> "í" And c2 <> "ú" Then
            EsDiptongo = True
            Exit Function
        End If
    End If
End Function

Private Function EsTriptongo(c1 As String, c2 As String, c3 As String) As Boolean
    ' débil + fuerte + débil (sin tildes en las débiles)
    If EsVocalDebil(c1) And EsVocalFuerte(c2) And EsVocalDebil(c3) Then
        If c1 <> "í" And c1 <> "ú" And c3 <> "í" And c3 <> "ú" Then
            EsTriptongo = True
        End If
    End If
End Function

Private Function EsHiatoFuerteFuerte(c1 As String, c2 As String) As Boolean
    ' AE, EA, AO, OA, EO, OE ? hiato siempre
    If EsVocalFuerte(c1) And EsVocalFuerte(c2) Then
        EsHiatoFuerteFuerte = True
    End If
End Function

Private Function EsGrupoInseparable_ES(par As String) As Boolean
    Dim g As Variant, lista As Variant

    ' Grupos consonánticos inseparables en minúsculas
    lista = Array("br", "bl", "cr", "cl", "dr", "fr", "gr", "gl", "pr", "pl", "tr")

    For Each g In lista
        If par = g Then
            EsGrupoInseparable_ES = True
            Exit Function
        End If
    Next g
End Function

Private Function MF_NormalizarVocales_ES(ByVal texto As String) As String

    ' a
    texto = Replace(texto, "á", "a")
    texto = Replace(texto, "à", "a")
    texto = Replace(texto, "ä", "a")
    texto = Replace(texto, "â", "a")

    ' e
    texto = Replace(texto, "é", "e")
    texto = Replace(texto, "è", "e")
    texto = Replace(texto, "ë", "e")
    texto = Replace(texto, "ê", "e")

    ' i
    texto = Replace(texto, "í", "i")
    texto = Replace(texto, "ì", "i")
    texto = Replace(texto, "ï", "i")
    texto = Replace(texto, "î", "i")

    ' o
    texto = Replace(texto, "ó", "o")
    texto = Replace(texto, "ò", "o")
    texto = Replace(texto, "ö", "o")
    texto = Replace(texto, "ô", "o")

    ' u (sin tocar ü)
    texto = Replace(texto, "ú", "u")
    texto = Replace(texto, "ù", "u")
    texto = Replace(texto, "û", "u")
    texto = Replace(texto, "ü", "ü") ' no tocar

    MF_NormalizarVocales_ES = texto

End Function

'==================================================================

'' ============================================================
''   ReglasCastellano (ESP)
''   Devuelve idFonema según la fonética del castellano.
''   Si no aplica, devuelve 0 para que el motor siga probando.
'' ============================================================
'Private Function ReglasCastellano( _
'        ByVal graf As String, _
'        ByVal ant As String, _
'        ByVal sig As String, _
'        ByVal esTonica As Boolean _
'    ) As Byte
'
'    Dim g As String
'    g = LCase$(graf)
'
'    ' ============================================================
'    '   TRIGRAFEMAS
'    ' ============================================================
'
'    ' güe / güi --> /gw/ ? id 57
'    If g = "güe" Or g = "güi" Then
'        ReglasCastellano = 57
'        Exit Function
'    End If
'
'    ' gue / gui --> /g/ (u muda) ? id 31
'    If g = "gue" Or g = "gui" Then
'        ReglasCastellano = 31
'        Exit Function
'    End If
'
'    ' que / qui --> /k/ ? id 30
'    If g = "que" Or g = "qui" Then
'        ReglasCastellano = 30
'        Exit Function
'    End If
'
'
'    ' ============================================================
'    '   DÍGRAFOS Y CASOS ESPECIALES
'    ' ============================================================
'
'    ' ch ? id 50
'    If g = "ch" Then
'        ReglasCastellano = 50
'        Exit Function
'    End If
'
'    ' ll ? id 44
'    If g = "ll" Then
'        ReglasCastellano = 44
'        Exit Function
'    End If
'
'    ' rr ? id 46
'    If g = "rr" Then
'        ReglasCastellano = 46
'        Exit Function
'    End If
'
'    ' ñ ? id 41
'    If g = "ñ" Then
'        ReglasCastellano = 41
'        Exit Function
'    End If
'
'    ' gu + vocal ? /g/ ? id 31
'    If g = "gu" And (sig = "a" Or sig = "o" Or sig = "u") Then
'        ReglasCastellano = 31
'        Exit Function
'    End If
'
'    ' qu + vocal ? /k/ ? id 30
'    If g = "qu" And (sig = "a" Or sig = "o" Or sig = "u") Then
'        ReglasCastellano = 30
'        Exit Function
'    End If
'
'
'    ' ============================================================
'    '   DÍGRAFOS VOCÁLICOS (diptongos)
'    ' ============================================================
'
'    If g = "ai" Then ReglasCastellano = 12: Exit Function
'    If g = "ei" Then ReglasCastellano = 13: Exit Function
'    If g = "oi" Then ReglasCastellano = 14: Exit Function
'    If g = "ou" Then ReglasCastellano = 15: Exit Function
'    If g = "au" Then ReglasCastellano = 16: Exit Function
'
'
'    ' ============================================================
'    '   MONÓGRAFOS — VOCALES
'    ' ============================================================
'
'    If g = "a" Then ReglasCastellano = 1: Exit Function
'    If g = "e" Then ReglasCastellano = 5: Exit Function
'    If g = "i" Then ReglasCastellano = 9: Exit Function
'    If g = "o" Then ReglasCastellano = 7: Exit Function
'    If g = "u" Then ReglasCastellano = 10: Exit Function
'
'
'    ' ============================================================
'    '   MONÓGRAFOS — CONSONANTES
'    ' ============================================================
'
'    If g = "p" Then ReglasCastellano = 26: Exit Function
'    If g = "b" Or g = "v" Then ReglasCastellano = 27: Exit Function
'    If g = "t" Then ReglasCastellano = 28: Exit Function
'    If g = "d" Then ReglasCastellano = 29: Exit Function
'    If g = "k" Then ReglasCastellano = 30: Exit Function
'    If g = "g" Then ReglasCastellano = 31: Exit Function
'
'    If g = "f" Then ReglasCastellano = 32: Exit Function
'
'    ' c/z ? /?/ (castellano estándar)
'    If g = "c" And (sig = "e" Or sig = "i") Then
'        ReglasCastellano = 54
'        Exit Function
'    End If
'    If g = "z" Then
'        ReglasCastellano = 54
'        Exit Function
'    End If
'
'    ' s ? /s/
'    If g = "s" Then ReglasCastellano = 34: Exit Function
'
'    ' j / g + e/i ? /x/ ? id 58
'    If g = "j" Then ReglasCastellano = 58: Exit Function
'    If g = "g" And (sig = "e" Or sig = "i") Then
'        ReglasCastellano = 58
'        Exit Function
'    End If
'
'    ' m / n
'    If g = "m" Then ReglasCastellano = 39: Exit Function
'    If g = "n" Then ReglasCastellano = 40: Exit Function
'
'    ' l / r simple
'    If g = "l" Then ReglasCastellano = 43: Exit Function
'    If g = "r" Then ReglasCastellano = 45: Exit Function
'
'    ' h muda ? id 38
'    If g = "h" Then ReglasCastellano = 38: Exit Function
'
'
'    ' ============================================================
'    '   SI NO APLICA
'    ' ============================================================
'    ReglasCastellano = 0
'
'End Function


'Private Function EsVocalFuerte_ES(c As String) As Boolean
'    ' Vocales fuertes: a, e, o (con y sin tilde)
'    EsVocalFuerte_ES = (InStr("aáeéoó", c) > 0)
'End Function
'
'Private Function EsVocalDebilTilde(c As String) As Boolean
'    ' Vocales débiles tildadas: í, ú
'    EsVocalDebilTilde = (c = "í" Or c = "ú")
'End Function


'Private Function EsDiptongo_ES(c1 As String, c2 As String) As Boolean
'    Dim d As Variant, lista As Variant
'
'    ' Diptongos en minúsculas
'    lista = Array( _
'        "ai", "ei", "oi", "ui", _
'        "au", "eu", "ou", _
'        "ia", "ie", "io", "iu", _
'        "ua", "ue", "uo" _
'    )
'
'    For Each d In lista
'        If c1 & c2 = d Then
'            EsDiptongo_ES = True
'            Exit Function
'        End If
'    Next d
'End Function

'Private Function EsTriptongo_ES(c1 As String, c2 As String, c3 As String) As Boolean
'    ' débil + fuerte + débil
'    If InStr("iu", c1) > 0 And InStr("aeo", c2) > 0 And InStr("iu", c3) > 0 Then
'        EsTriptongo_ES = True
'    End If
'End Function

'Private Sub RevisionSilabeo()
'
'    Dim s As String
'    Dim partes() As String
'    Dim sils() As String
'    Dim i As Long, j As Long
'    Dim idxTonicaManual As Long
'    Dim idxTonicaVisual As Long
'    Dim resultado() As String
'    Dim tFinal() As Byte
'    Dim texto As String
'
'    texto = objDTO.TextoNormalizado
'
'    ' ============================================================
'    ' 1. Construir cadena editable a partir de SilabasAuto()
'    ' ============================================================
'    s = ""
'    For i = 1 To UBound(objDTO.SilabasAuto)
'        s = s & objDTO.SilabasAuto(i)
'        If i < UBound(objDTO.SilabasAuto) Then s = s & "-"
'    Next i
'
'    ' ============================================================
'    ' 2. Mostrar formulario de revisión
'    ' ============================================================
'    s = RevisarSilabas_EnFormulario(texto, s)
'
'    ' Si el usuario cancela ? mantener silabeo automático
'    If s = "" Then
'        objDTO.SilabasFinal = objDTO.SilabasAuto
'        objDTO.SilabaTonica = objDTO.SilabaTonica   ' automática
'        Exit Sub
'    End If
'
'    ' ============================================================
'    ' 3. Detectar tónica manual con "*"
'    ' ============================================================
'    partes = Split(Replace(s, " ", ""), "-")
'    idxTonicaManual = 0
'
'    For i = LBound(partes) To UBound(partes)
'        If Left$(partes(i), 1) = "*" Then
'            idxTonicaManual = i + 1
'            partes(i) = Mid$(partes(i), 2)
'        End If
'    Next i
'
'    ' ============================================================
'    ' 4. Detectar tónica visual (MAYÚSCULAS)
'    ' ============================================================
'    idxTonicaVisual = 0
'
'    For i = LBound(partes) To UBound(partes)
'        If partes(i) = UCase$(partes(i)) And partes(i) <> LCase$(partes(i)) Then
'            idxTonicaVisual = i + 1
'        End If
'    Next i
'
'    ' ============================================================
'    ' 5. Reconstruir SilabasFinal()
'    ' ============================================================
'    ReDim resultado(1 To UBound(partes) + 1)
'
'    For i = 1 To UBound(resultado)
'        resultado(i) = partes(i - 1)
'    Next i
'
'    objDTO.SilabasFinal = resultado
'
'    ' ============================================================
'    ' 6. Reconstruir SilabaTonica()
'    ' ============================================================
'    If idxTonicaManual > 0 Then
'        ReDim tFinal(1 To 1)
'        tFinal(1) = idxTonicaManual
'
'    ElseIf idxTonicaVisual > 0 Then
'        ReDim tFinal(1 To 1)
'        tFinal(1) = idxTonicaVisual
'
'    Else
'        ' Mantener la tónica automática
'        objDTO.SilabasFinal = resultado
'        Exit Sub
'    End If
'
'    objDTO.SilabaTonica = tFinal
'
'End Sub


'Private Sub MarcarTonicaEnSilaba()
'
'    Dim i As Long, j As Long
'    Dim resultado As String
'    Dim idx As Long
'    Dim esTonica As Boolean
'
'    ' Validaciones
'    If UBound(objDTO.SilabasFinal) < 1 Then Exit Sub
'    If UBound(objDTO.SilabaTonica) < 1 Then Exit Sub
'
'    resultado = ""
'
'    ' Recorrer sílabas finales
'    For i = 1 To UBound(objDTO.SilabasFinal)
'
'        ' ¿Esta sílaba es tónica?
'        esTonica = False
'        For j = 1 To UBound(objDTO.SilabaTonica)
'            If objDTO.SilabaTonica(j) = i Then
'                esTonica = True
'                Exit For
'            End If
'        Next j
'
'        ' Añadir sílaba marcada o normal
'        If esTonica Then
'            resultado = resultado & UCase$(objDTO.SilabasFinal(i))
'        Else
'            resultado = resultado & objDTO.SilabasFinal(i)
'        End If
'
'        ' Separador opcional
'        If i < UBound(objDTO.SilabasFinal) Then
'            resultado = resultado & "-"
'        End If
'    Next i
'
'    ' Guardar resultado final
'    objDTO.TextoFinal = resultado
'
'End Sub

