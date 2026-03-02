Attribute VB_Name = "bas_Motor_IB_Fon_3"

'Option Compare Database
'Option Explicit
'
'' ============================================================
''   MOTOR FONÈTIC — BALEAR (IB) — Versió corregida V6.0a
'' ============================================================
'
'' ============================================================
''   BLOQUE 1 — ConstruirCadenaFonemas_IB
'' ============================================================
'Public Sub ConstruirCadenaFonemas_IB()
'
'    Dim arrSilabas As Variant
'    Dim i As Long
'    Dim silabaCruda As String
'    Dim ultimaSilaba As String
'    Dim siguienteSilaba As String
'    Dim frase As String
'    Dim arrFon As Variant
'    Dim Fon As Variant
'    Dim res As String
'    Dim cadenaReglas As String
'
'    ObjDTO.IdsFonemas = ""
'    ObjDTO.FonemasFinal = ""
'
'    frase = NormalizarVocales(ObjDTO.SilabasFinal)
'    arrSilabas = Split(frase, "|")
'
'    For i = 0 To UBound(arrSilabas)
'
'        silabaCruda = arrSilabas(i)
'
'        ' Separador de palabra
'        If Trim$(silabaCruda) = "" Then
'
'            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "#"
'
'            ultimaSilaba = BuscarSilabaRealAnterior(arrSilabas, i)
'            siguienteSilaba = BuscarSilabaRealPosterior(arrSilabas, i)
'
'            If HayLigaduraAutomatica(ultimaSilaba, siguienteSilaba) Then
'                ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "84,"
'            End If
'
'            GoTo SiguienteIteracion
'        End If
'
'        ' Acento
'        If InStr(silabaCruda, "(") > 0 Then
'            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "80,"
'        ElseIf InStr(silabaCruda, "[") > 0 Then
'            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "81,"
'        Else
'            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "82,"
'        End If
'
'        ' Procesar grafemas
'        ProcesarSilaba silabaCruda
'
'        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "|"
'
'SiguienteIteracion:
'    Next i
'
'    ' Limpieza final
'    ObjDTO.IdsFonemas = Replace(ObjDTO.IdsFonemas, "|#", "#")
'    ObjDTO.IdsFonemas = Replace(ObjDTO.IdsFonemas, "#|", "#")
'    ObjDTO.IdsFonemas = Replace(ObjDTO.IdsFonemas, ",|", "|")
'    ObjDTO.IdsFonemas = Replace(ObjDTO.IdsFonemas, "84,84,", "84,")
'
'    While Right$(ObjDTO.IdsFonemas, 1) = "|" Or _
'          Right$(ObjDTO.IdsFonemas, 1) = "#" Or _
'          Right$(ObjDTO.IdsFonemas, 1) = ","
'
'        ObjDTO.IdsFonemas = Left$(ObjDTO.IdsFonemas, Len(ObjDTO.IdsFonemas) - 1)
'    Wend
'
'    ' ============================================================
'    '   APLICAR REGLAS FONÉTICAS IB
'    ' ============================================================
'    cadenaReglas = Replace(ObjDTO.IdsFonemas, ",", " ")
'
'    cadenaReglas = AplicarNeutralizaciones(cadenaReglas)
'    cadenaReglas = AplicarAssimilaciones(cadenaReglas)
'    cadenaReglas = AplicarSchwa(cadenaReglas)
'    cadenaReglas = AplicarReducciones(cadenaReglas)
'
'    While InStr(cadenaReglas, "  ") > 0
'        cadenaReglas = Replace(cadenaReglas, "  ", " ")
'    Wend
'
'    cadenaReglas = Trim$(Replace(cadenaReglas, " ", ","))
'    ObjDTO.IdsFonemas = cadenaReglas
'
'    ' ============================================================
'    '   CONSTRUCCIÓN IPA
'    ' ============================================================
'    arrSilabas = Split(ObjDTO.IdsFonemas, "|")
'    res = ""
'
'    For i = 0 To UBound(arrSilabas)
'
'        If i = 0 Then res = res & "/"
'
'        arrFon = Split(arrSilabas(i), ",")
'
'        For Each Fon In arrFon
'
'            If Fon = "#84" Then
'                Fon = 84
'            ElseIf Fon = "#82" Then
'                res = res & "/ /"
'                Fon = 82
'            End If
'
'            If Left$(Fon, 1) = "#" Then
'                res = res & "/ /"
'                Fon = Replace(Fon, "#", "")
'            End If
'
'            If Trim$(Fon) <> "" And Trim$(Fon) <> "82" Then
'                res = res & Replace(ObtenerIPA(Fon), "/", "")
'            End If
'
'        Next Fon
'
'        If i = UBound(arrSilabas) Then
'            res = res & "/ "
'        End If
'
'    Next i
'
'    ObjDTO.FonemasFinal = Trim$(res)
'
'End Sub
'
'' ============================================================
''   BLOQUE 2 — ProcesarSilaba (versión mallorquina corregida)
'' ============================================================
'Private Sub ProcesarSilaba(ByVal silabaCruda As String)
'
'    Dim silaba As String
'    Dim ligaduraID As Byte
'    Dim i As Long
'    Dim grafema As String
'    Dim id As Byte
'    Dim antCh As String
'    Dim sigCh As String
'    Dim esAtona As Boolean
'
'    esAtona = True
'    If InStr(silabaCruda, "(") > 0 Then esAtona = False
'    If InStr(silabaCruda, "[") > 0 Then esAtona = False
'
'    silaba = silabaCruda
'    silaba = Replace(silaba, "(", "")
'    silaba = Replace(silaba, ")", "")
'    silaba = Replace(silaba, "[", "")
'    silaba = Replace(silaba, "]", "")
'    silaba = Replace(silaba, "'", "")
'    silaba = Replace(silaba, "’", "")
'    ' NO eliminar "_" aquí — se usa para ligadura manual
'
'    silaba = Trim$(silaba)
'    If silaba = "" Then Exit Sub
'
'    ' Ligadura manual
'    ligaduraID = DetectarLigaduraManual(silaba)
'    If ligaduraID <> 0 Then
'        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & CStr(ligaduraID) & ","
'    End If
'
'    ' Ahora sí podemos eliminar "_" porque ya se detectó
'    silaba = Replace(silaba, "_", "")
'
'    i = 1
'    Do While i <= Len(silaba)
'
'        grafema = DetectarDigrafo(silaba, i)
'
'        ' Seguridad: si DetectarDigrafo devolviera vacío (no debería), evitar 255
'        If grafema = "" Then
'            i = i + 1
'            GoTo SegGraf
'        End If
'
'        If i > 1 Then antCh = Mid$(silaba, i - 1, 1) Else antCh = ""
'        If i + Len(grafema) <= Len(silaba) Then
'            sigCh = Mid$(silaba, i + Len(grafema), 1)
'        Else
'            sigCh = ""
'        End If
'
'        ' SCHWA IB (solo proclíticos)
'        If esAtona Then
'            If EsProclitic(silabaCruda) Then
'                If grafema = "a" Or grafema = "e" Then
'                    ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "8,"
'                    i = i + Len(grafema)
'                    GoTo SegGraf
'                End If
'            End If
'        End If
'
'        ' --- GU + a/o/u ? /g/ + /w/ ---
'        If grafema = "gu" Then
'            If sigCh Like "[aouàòóú]" Then
'                ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "35,22,"
'                i = i + 2
'                GoTo SegGraf
'            End If
'        End If
'
'        ' --- GÜ + e/i ? /g/ + /w/ ---
'        If grafema = "gü" Then
'            If sigCh Like "[eéiíè]" Then
'                ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "35,22,"
'                i = i + 2
'                GoTo SegGraf
'            End If
'        End If
'
'        ' --- QU + e/i ? /k/ ---
'        If grafema = "qu" Then
'            If sigCh Like "[eéiíè]" Then
'                ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "34,"
'                i = i + 2
'                GoTo SegGraf
'            End If
'        End If
'
'        ' --- QÜ + e/i ? /k/ + /w/ ---
'        If grafema = "qü" Then
'            If sigCh Like "[eéiíè]" Then
'                ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "34,22,"
'                i = i + 3
'                GoTo SegGraf
'            End If
'        End If
'
'        ' --- X intervocàlica ? /ks/ ---
'        If grafema = "x" Then
'            If antCh Like "[aeiouàèéíïòóúü]" And sigCh Like "[aeiouàèéíïòóúü]" Then
'                ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "34,42,"
'                i = i + 1
'                GoTo SegGraf
'            End If
'        End If
'
'        ' --- EX + vocal ? /eks/ ---
'        If grafema = "x" Then
'            If antCh = "e" And sigCh Like "[aeiouàèéíïòóúü]" Then
'                ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "34,42,"
'                i = i + 1
'                GoTo SegGraf
'            End If
'        End If
'
'        ' --- IX ? /?/ ---
'        If grafema = "ix" Then
'            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "44,"
'            i = i + 2
'            GoTo SegGraf
'        End If
'
'        ' --- TX ? /t?/ ---
'        If grafema = "tx" Then
'            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "57,"
'            i = i + 2
'            GoTo SegGraf
'        End If
'
'        ' --- TG / TJ ? /d?/ ---
'        If grafema = "tg" Or grafema = "tj" Then
'            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "58,"
'            i = i + 2
'            GoTo SegGraf
'        End If
'
'        ' --- TS / TZ ? /ts/ ---
'        If grafema = "ts" Or grafema = "tz" Then
'            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "46,"
'            i = i + 2
'            GoTo SegGraf
'        End If
'
'        ' Fonema base
'        id = AsignarFonemaBase(grafema, sigCh, antCh)
'        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & CStr(id) & ","
'        i = i + Len(grafema)
'
'SegGraf:
'    Loop
'
'End Sub
'
'Private Function EsProclitic(sil As String) As Boolean
'    Dim s As String
'    s = LCase$(Trim$(sil))
'    EsProclitic = (s = "de" Or s = "me" Or s = "te" Or s = "se" Or s = "es" Or s = "sa" Or s = "ses")
'End Function
'
'' ============================================================
''   DETECTAR DÍGRAFS — BALEAR (IB)  (VERSIÓ CORREGIDA)
'' ============================================================
'Private Function DetectarDigrafo(t As String, pos As Long) As String
'
'    Dim L As Long: L = Len(t)
'    Dim par As String, tri As String, sig As String
'
'    ' --- 1. TRIGRAMA L·L ---
'    If pos + 2 <= L Then
'        tri = Mid$(t, pos, 3)
'        If tri = "l·l" Then
'            DetectarDigrafo = tri
'            Exit Function
'        End If
'    End If
'
'    ' --- 2. DÍGRAFS DE 2 LLETRES ---
'    If pos + 1 <= L Then
'        par = Mid$(t, pos, 2)
'
'        Select Case par
'            Case "ny", "ll", "rr", "ss"
'                DetectarDigrafo = par: Exit Function
'
'            Case "tx", "tg", "tj", "ts", "tz", "ix"
'                DetectarDigrafo = par: Exit Function
'        End Select
'
'        ' QU + e/i
'        If par = "qu" Then
'            If pos + 2 <= L Then
'                sig = Mid$(t, pos + 2, 1)
'                If sig Like "[eéiíè]" Then
'                    DetectarDigrafo = par
'                    Exit Function
'                End If
'            End If
'        End If
'
'        ' QÜ + e/i
'        If par = "qü" Then
'            If pos + 3 <= L Then
'                sig = Mid$(t, pos + 3, 1)
'                If sig Like "[eéiíè]" Then
'                    DetectarDigrafo = par
'                    Exit Function
'                End If
'            End If
'        End If
'
'        ' GÜ + e/i
'        If par = "gü" Then
'            If pos + 3 <= L Then
'                sig = Mid$(t, pos + 3, 1)
'                If sig Like "[eéiíè]" Then
'                    DetectarDigrafo = par
'                    Exit Function
'                End If
'            End If
'        End If
'
'        ' GU + a/o/u
'        If par = "gu" Then
'            If pos + 2 <= L Then
'                sig = Mid$(t, pos + 2, 1)
'                If sig Like "[aouàòóú]" Then
'                    DetectarDigrafo = par
'                    Exit Function
'                End If
'            End If
'        End If
'
'    End If
'
'    ' --- 3. IG final ? TX ---
'    If pos + 1 = L Then
'        If Mid$(t, pos, 2) = "ig" Then
'            DetectarDigrafo = "tx"
'            Exit Function
'        End If
'    End If
'
'    ' --- 4. Si no es dígrafo ? 1 letra ---
'    DetectarDigrafo = Mid$(t, pos, 1)
'
'End Function
'
'' ============================================================
''   FONEMA BASE — BALEAR (IB) — Versió corregida i completa
'' ============================================================
'Private Function AsignarFonemaBase(grafema As String, _
'                                   Optional sig As String = "", _
'                                   Optional ant As String = "") As Byte
'
'    grafema = LCase$(grafema)
'    sig = LCase$(sig)
'    ant = LCase$(ant)
'
'    ' ------------------------------------------------------------
'    ' VOCALS
'    ' ------------------------------------------------------------
'    Select Case grafema
'        Case "a", "à", "á": AsignarFonemaBase = 1: Exit Function
'        Case "e", "é": AsignarFonemaBase = 2: Exit Function
'        Case "è": AsignarFonemaBase = 3: Exit Function
'        Case "i", "í", "ï": AsignarFonemaBase = 4: Exit Function
'        Case "o", "ó": AsignarFonemaBase = 5: Exit Function
'        Case "ò": AsignarFonemaBase = 6: Exit Function
'        Case "u", "ú", "ü": AsignarFonemaBase = 7: Exit Function
'    End Select
'
'    ' ------------------------------------------------------------
'    ' SEMIVOCALS
'    ' ------------------------------------------------------------
'    If grafema = "j" Or grafema = "y" Then AsignarFonemaBase = 21: Exit Function
'    If grafema = "w" Then AsignarFonemaBase = 22: Exit Function
'
'    ' ------------------------------------------------------------
'    ' NASALS
'    ' ------------------------------------------------------------
'    Select Case grafema
'        Case "m": AsignarFonemaBase = 36: Exit Function
'        Case "n": AsignarFonemaBase = 37: Exit Function
'        Case "ny": AsignarFonemaBase = 38: Exit Function
'        Case "ng": AsignarFonemaBase = 39: Exit Function
'    End Select
'
'    ' ------------------------------------------------------------
'    ' LATERALS
'    ' ------------------------------------------------------------
'    If grafema = "l" Then AsignarFonemaBase = 62: Exit Function
'    If grafema = "ll" Then AsignarFonemaBase = 63: Exit Function
'    If grafema = "l·l" Then AsignarFonemaBase = 63: Exit Function
'
'    ' ------------------------------------------------------------
'    ' VIBRANTS
'    ' ------------------------------------------------------------
'    If grafema = "rr" Then AsignarFonemaBase = 60: Exit Function
'
'    If grafema = "r" Then
'        If ant = "" Or ant = " " Or Not ant Like "[aeiouàèéíïòóúü]" Then
'            AsignarFonemaBase = 60
'        Else
'            AsignarFonemaBase = 59
'        End If
'        Exit Function
'    End If
'
'    ' ------------------------------------------------------------
'    ' AFRICADES
'    ' ------------------------------------------------------------
'    If grafema = "tx" Then AsignarFonemaBase = 57: Exit Function
'    If grafema = "tg" Or grafema = "tj" Then AsignarFonemaBase = 58: Exit Function
'    If grafema = "ts" Or grafema = "tz" Then AsignarFonemaBase = 46: Exit Function
'
'    ' ------------------------------------------------------------
'    ' FRICATIVES
'    ' ------------------------------------------------------------
'    If grafema = "s" Then AsignarFonemaBase = 42: Exit Function
'    If grafema = "ss" Then AsignarFonemaBase = 42: Exit Function
'    If grafema = "z" Then AsignarFonemaBase = 43: Exit Function
'    If grafema = "x" Then AsignarFonemaBase = 44: Exit Function
'    If grafema = "ix" Then AsignarFonemaBase = 44: Exit Function
'    If grafema = "j" Then AsignarFonemaBase = 44: Exit Function
'    If grafema = "ge" Or grafema = "gi" Then AsignarFonemaBase = 45: Exit Function
'
'    ' ------------------------------------------------------------
'    ' OCLUSIVES
'    ' ------------------------------------------------------------
'    Select Case grafema
'        Case "p": AsignarFonemaBase = 30: Exit Function
'        Case "b": AsignarFonemaBase = 31: Exit Function
'        Case "t": AsignarFonemaBase = 32: Exit Function
'        Case "d": AsignarFonemaBase = 33: Exit Function
'
'        Case "c"
'            If sig Like "[eéiíè]" Then
'                AsignarFonemaBase = 42   ' /s/
'            Else
'                AsignarFonemaBase = 34   ' /k/
'            End If
'            Exit Function
'
'        Case "k", "qu": AsignarFonemaBase = 34: Exit Function
'
'        Case "g"
'            If sig Like "[eéiíè]" Then
'                AsignarFonemaBase = 45   ' /?/
'            Else
'                AsignarFonemaBase = 35   ' /g/
'            End If
'            Exit Function
'
'        Case "gu": AsignarFonemaBase = 35: Exit Function
'
'        Case "v": AsignarFonemaBase = 41: Exit Function
'        Case "f": AsignarFonemaBase = 40: Exit Function
'    End Select
'
'    AsignarFonemaBase = 255
'End Function
'
''Private Function AsignarFonemaBase(grafema As String, _
''                                   Optional sig As String = "", _
''                                   Optional ant As String = "") As Byte
''
''    grafema = LCase$(grafema)
''    sig = LCase$(sig)
''    ant = LCase$(ant)
''
''    ' ------------------------------------------------------------
''    ' VOCALS
''    ' ------------------------------------------------------------
''    Select Case grafema
''        Case "a", "à", "á": AsignarFonemaBase = 1: Exit Function
''        Case "e", "é": AsignarFonemaBase = 2: Exit Function
''        Case "è": AsignarFonemaBase = 3: Exit Function
''        Case "i", "í", "ï": AsignarFonemaBase = 4: Exit Function
''        Case "o", "ó": AsignarFonemaBase = 5: Exit Function
''        Case "ò": AsignarFonemaBase = 6: Exit Function
''        Case "u", "ú", "ü": AsignarFonemaBase = 7: Exit Function
''    End Select
''
''    ' ------------------------------------------------------------
''    ' SEMIVOCALS
''    ' ------------------------------------------------------------
''    If grafema = "j" Or grafema = "y" Then AsignarFonemaBase = 21: Exit Function
''    If grafema = "w" Then AsignarFonemaBase = 22: Exit Function
''
''    ' ------------------------------------------------------------
''    ' NASALS
''    ' ------------------------------------------------------------
''    Select Case grafema
''        Case "m": AsignarFonemaBase = 36: Exit Function
''        Case "n": AsignarFonemaBase = 37: Exit Function
''        Case "ny": AsignarFonemaBase = 38: Exit Function
''        Case "ng": AsignarFonemaBase = 39: Exit Function
''    End Select
''
''    ' ------------------------------------------------------------
''    ' LATERALS
''    ' ------------------------------------------------------------
''    If grafema = "l" Then AsignarFonemaBase = 62: Exit Function
''    If grafema = "ll" Then AsignarFonemaBase = 63: Exit Function
''    If grafema = "l·l" Then AsignarFonemaBase = 63: Exit Function
''
''    ' ------------------------------------------------------------
''    ' VIBRANTS
''    ' ------------------------------------------------------------
''    If grafema = "rr" Then AsignarFonemaBase = 60: Exit Function
''
''    If grafema = "r" Then
''        ' r inicial o després de consonant ? /r/
''        If ant = "" Or ant = " " Or Not ant Like "[aeiouàèéíïòóúü]" Then
''            AsignarFonemaBase = 60
''        Else
''            ' r intervocàlica ? /?/
''            AsignarFonemaBase = 59
''        End If
''        Exit Function
''    End If
''
''    ' ------------------------------------------------------------
''    ' AFRICADES
''    ' ------------------------------------------------------------
''    If grafema = "tx" Then AsignarFonemaBase = 57: Exit Function
''    If grafema = "tg" Or grafema = "tj" Then AsignarFonemaBase = 58: Exit Function
''    If grafema = "ts" Or grafema = "tz" Then AsignarFonemaBase = 46: Exit Function
''
''    ' ------------------------------------------------------------
''    ' FRICATIVES
''    ' ------------------------------------------------------------
''    If grafema = "s" Then AsignarFonemaBase = 42: Exit Function
''    If grafema = "z" Then AsignarFonemaBase = 43: Exit Function
''    If grafema = "x" Then AsignarFonemaBase = 44: Exit Function
''    If grafema = "ix" Then AsignarFonemaBase = 44: Exit Function
''    If grafema = "j" Then AsignarFonemaBase = 45: Exit Function
''    If grafema = "ge" Or grafema = "gi" Then AsignarFonemaBase = 45: Exit Function
''
''' ------------------------------------------------------------
''' OCLUSIVES
''' ------------------------------------------------------------
''Select Case grafema
''    Case "p": AsignarFonemaBase = 30: Exit Function
''    Case "b": AsignarFonemaBase = 31: Exit Function
''    Case "t": AsignarFonemaBase = 32: Exit Function
''    Case "d": AsignarFonemaBase = 33: Exit Function
''
''    ' C: /k/ o /s/ según vocal siguiente
''    Case "c"
''        If sig Like "[eéiíè]" Then
''            AsignarFonemaBase = 42   ' /s/
''        Else
''            AsignarFonemaBase = 34   ' /k/
''        End If
''        Exit Function
''
''    ' K y QU: siempre /k/
''    Case "k", "qu": AsignarFonemaBase = 34: Exit Function
''
''    ' G: /g/ o /?/ según vocal siguiente
''    Case "g"
''        If sig Like "[eéiíè]" Then
''            AsignarFonemaBase = 45   ' /?/
''        Else
''            AsignarFonemaBase = 35   ' /g/
''        End If
''        Exit Function
''
''    ' GU: /gw/ ante a/o/u, /g/ en otros casos
''    Case "gu"
''        If sig Like "[aouàòóú]" Then
''            AsignarFonemaBase = 35   ' /g/ (la /w/ ya la tratamos aparte)
''        Else
''            AsignarFonemaBase = 35
''        End If
''        Exit Function
''
''    Case "v": AsignarFonemaBase = 41: Exit Function
''
''    ' F: aquí corrige el error ? /f/
''    Case "f": AsignarFonemaBase = 40: Exit Function
''End Select
''
'''    ' ------------------------------------------------------------
'''    ' OCLUSIVES
'''    ' ------------------------------------------------------------
'''    Select Case grafema
'''        Case "v": AsignarFonemaBase = 41: Exit Function
'''        Case "p": AsignarFonemaBase = 30: Exit Function
'''        Case "b": AsignarFonemaBase = 31: Exit Function
'''        Case "t": AsignarFonemaBase = 32: Exit Function
'''        Case "d": AsignarFonemaBase = 33: Exit Function
'''        Case "c", "k", "qu": AsignarFonemaBase = 34: Exit Function
'''        Case "g", "gu": AsignarFonemaBase = 35: Exit Function
'''        Case "q": AsignarFonemaBase = 34: Exit Function   ' ? CORRECCIÓN IB
'''        'Case "f": AsignarFonemaBase = 30: Exit Function   ' ? CORRECCIÓN IB
'''
'''        ' Seleccionar una de las dos
'''        'Case "f": AsignarFonemaBase = 42 'f
'''        Case "f": AsignarFonemaBase = 45 'f bilabial mallorquina
'''
'''    End Select
''
''    ' ------------------------------------------------------------
''    ' SI NO ES RECONEIX ? 255
''    ' ------------------------------------------------------------
''    AsignarFonemaBase = 255
''
''End Function
'
'' ============================================================
''   NEUTRALITZACIONS IB
'' ============================================================
'' ============================================================
''   NEUTRALITZACIONS IB — Versió corregida (no neutralitza QU)
'' ============================================================
'' ============================================================
''   NEUTRALITZACIONS IB — VERSIÓ DEFINITIVA
'' ============================================================
'Private Function AplicarNeutralizaciones(cadena As String) As String
'    Dim arr As Variant
'    Dim i As Long
'    Dim out As String
'
'    arr = Split(Trim$(cadena), " ")
'
'    For i = 0 To UBound(arr)
'
'        ' ------------------------------------------------------------
'        ' 1) ASSIMILACIÓ DE NASALS
'        ' ------------------------------------------------------------
'        ' n ? m davant de p/b
'        If arr(i) = "37" Then
'            If i < UBound(arr) Then
'                If arr(i + 1) = "30" Or arr(i + 1) = "31" Then
'                    arr(i) = "36"   ' m
'                End If
'            End If
'        End If
'
'        ' n ? ? davant de k/g
'        If arr(i) = "37" Then
'            If i < UBound(arr) Then
'                If arr(i + 1) = "34" Or arr(i + 1) = "35" Then
'                    arr(i) = "39"   ' ng
'                End If
'            End If
'        End If
'
'        ' ------------------------------------------------------------
'        ' 2) REDUCCIÓ DE GEMINADES
'        ' ------------------------------------------------------------
'        If arr(i) = "86" Then
'            arr(i) = ""   ' eliminar marca de geminació
'        End If
'
'        ' ------------------------------------------------------------
'        ' 3) SCHWA: eliminar si està entre consonants
'        ' ------------------------------------------------------------
'        If arr(i) = "8" Then
'            If i > 0 And i < UBound(arr) Then
'                If arr(i - 1) <> "" And arr(i + 1) <> "" Then
'                    If CInt(arr(i - 1)) > 20 And CInt(arr(i + 1)) > 20 Then
'                        arr(i) = ""
'                    End If
'                End If
'            End If
'        End If
'
'        ' ------------------------------------------------------------
'        ' 4) ENLLAÇ / LIGADURA
'        ' ------------------------------------------------------------
''        If arr(i) = "84" Then
''            arr(i) = ""   ' la lligadura no és un fonema
''        End If
'
'    Next i
'
'    ' ------------------------------------------------------------
'    ' 5) R FINAL SIMPLE ? eliminar
'    ' ------------------------------------------------------------
'    If arr(UBound(arr)) = "59" Then
'        arr(UBound(arr)) = ""
'    End If
'
'    out = Join(arr, " ")
'    AplicarNeutralizaciones = Trim$(out)
'End Function
'
'
''Private Function AplicarNeutralizaciones(cadena As String) As String
''    Dim arr As Variant
''    Dim i As Long
''    Dim out As String
''
''    arr = Split(Trim$(cadena), " ")
''
''    For i = 0 To UBound(arr)
''
'''        ' ------------------------------------------------------------
'''        ' CE / CI ? SE / SI
'''        ' PERÒ NO NEUTRALITZAR QU + E/I  (quests, que, qui, quelcom...)
'''        ' ------------------------------------------------------------
'''        If arr(i) = "34" Then
'''            If i < UBound(arr) Then
'''
'''                ' Vocal següent = E/I
'''                If arr(i + 1) = "2" Or arr(i + 1) = "4" Then
'''
'''                    ' Miram si a la cadena original hi ha "qu" just abans
'''                    ' Com que no tenim grafemes aquí, usam una heurística:
'''                    ' Si hi ha un 22 (semivocal /w/) just abans, prové de QU.
'''                    If i > 0 Then
'''                        If arr(i - 1) = "22" Then
'''                            ' És QU ? NO neutralitzar
'''                        Else
'''                            ' No és QU ? neutralitzar
'''                            arr(i) = "42"
'''                        End If
'''                    Else
'''                        ' Inici de cadena ? neutralitzar normal
'''                        arr(i) = "42"
'''                    End If
'''
'''                End If
'''            End If
'''        End If
'''
'''        ' ------------------------------------------------------------
'''        ' GE / GI ? JE / JI
'''        ' ------------------------------------------------------------
'''        If arr(i) = "35" Then
'''            If i < UBound(arr) Then
'''                If arr(i + 1) = "2" Or arr(i + 1) = "4" Then
'''                    arr(i) = "45"
'''                End If
'''            End If
'''        End If
'''
'''    Next i
''
''    ' ------------------------------------------------------------
''    ' R final simple ? eliminar
''    ' ------------------------------------------------------------
''    If arr(UBound(arr)) = "59" Then
''        arr(UBound(arr)) = ""
''    End If
''
''    out = Join(arr, " ")
''    AplicarNeutralizaciones = Trim$(out)
''End Function
'
''Private Function AplicarNeutralizaciones(cadena As String) As String
''    Dim arr As Variant
''    Dim i As Long
''    Dim out As String
''
''    arr = Split(Trim$(cadena), " ")
''
''    For i = 0 To UBound(arr)
''
''' CE / CI ? SE / SI
''' PERÒ NO NEUTRALITZAR "KE" (quests, que, quelcom...)
''If arr(i) = "34" Then
''    If i < UBound(arr) Then
''        ' Solo neutralizar si la letra original era "c" o "ç"
''        ' No neutralizar si venimos de "qu" o "k"
''        If arr(i - 1) <> "qu" And arr(i - 1) <> "k" Then
''            If arr(i + 1) = "2" Or arr(i + 1) = "4" Then
''                arr(i) = "42"
''            End If
''        End If
''    End If
''End If
''
'''        If arr(i) = "34" Then
'''            If i < UBound(arr) Then
'''                If arr(i + 1) = "2" Or arr(i + 1) = "4" Then
'''                    arr(i) = "42"
'''                End If
'''            End If
'''        End If
''
''        ' GE / GI ? JE / JI
''        If arr(i) = "35" Then
''            If i < UBound(arr) Then
''                If arr(i + 1) = "2" Or arr(i + 1) = "4" Then
''                    arr(i) = "45"
''                End If
''            End If
''        End If
''
''    Next i
''
''    ' R final simple ? eliminar
''    If arr(UBound(arr)) = "59" Then
''        arr(UBound(arr)) = ""
''    End If
''
''    out = Join(arr, " ")
''    AplicarNeutralizaciones = Trim$(out)
''End Function
'
'
'' ============================================================
''   ASSIMILACIONS IB
'' ============================================================
'Private Function AplicarAssimilaciones(cadena As String) As String
'    Dim arr As Variant
'    Dim i As Long
'
'    arr = Split(Trim$(cadena), " ")
'
'    For i = 0 To UBound(arr)
'
'        ' S intervocàlica ? Z
'        If arr(i) = "42" Then
'            If i > 0 And i < UBound(arr) Then
'                If EsVocalID(arr(i - 1)) And EsVocalID(arr(i + 1)) Then
'                    arr(i) = "43"
'                End If
'            End If
'        End If
'
'        ' N + K/G ? ?
'        If arr(i) = "37" Then
'            If i < UBound(arr) Then
'                If arr(i + 1) = "34" Or arr(i + 1) = "35" Then
'                    arr(i) = "39"
'                End If
'            End If
'        End If
'
'        ' R inicial ? RR
'        If i = 0 And arr(i) = "59" Then
'            arr(i) = "60"
'        End If
'
'        ' U + N ? U + ?
'        If arr(i) = "7" Then
'            If i < UBound(arr) Then
'                If arr(i + 1) = "37" Then
'                    arr(i + 1) = "39"
'                End If
'            End If
'        End If
'
'    Next i
'
'    AplicarAssimilaciones = Join(arr, " ")
'End Function
'
'
'' ============================================================
''   SCHWA IB
'' ============================================================
'Private Function AplicarSchwa(cadena As String) As String
'    Dim arr As Variant
'    Dim i As Long
'
'    arr = Split(Trim$(cadena), " ")
'
'    For i = 0 To UBound(arr)
'
'        ' de ? 33 8
'        If arr(i) = "33" Then
'            If i < UBound(arr) Then
'                If EsVocalID(arr(i + 1)) Then
'                    arr(i + 1) = "8"
'                End If
'            End If
'        End If
'
'        ' se ? 42 8
'        If arr(i) = "42" Then
'            If i < UBound(arr) Then
'                If EsVocalID(arr(i + 1)) Then
'                    arr(i + 1) = "8"
'                End If
'            End If
'        End If
'
'    Next i
'
'    AplicarSchwa = Join(arr, " ")
'End Function
'
'
'' ============================================================
''   REDUCCIONS IB
'' ============================================================
'Private Function AplicarReducciones(cadena As String) As String
'
'    Do While InStr(cadena, "  ") > 0
'        cadena = Replace(cadena, "  ", " ")
'    Loop
'
'    AplicarReducciones = Trim$(cadena)
'End Function
'
'
'' ============================================================
''   VOCALS — IDENTIFICACIÓ
'' ============================================================
'Private Function EsVocalID(id As Variant) As Boolean
'    Select Case id
'        Case "1", "2", "3", "4", "5", "6", "7", "8"
'            EsVocalID = True
'        Case Else
'            EsVocalID = False
'    End Select
'End Function
'
'' ============================================================
''   DETECTAR LIGADURA MANUAL
'' ============================================================
'Private Function DetectarLigaduraManual(ByRef silaba As String) As Byte
'    DetectarLigaduraManual = 0
'    If InStr(silaba, "_") > 0 Then
'        silaba = Replace(silaba, "_", "")
'        DetectarLigaduraManual = 84
'    End If
'End Function
'
'
'' ============================================================
''   LIGADURA AUTOMÀTICA ENTRE PARAULES
'' ============================================================
'Private Function HayLigaduraAutomatica(ultimaSilaba As String, primeraSilaba As String) As Boolean
'
'    Dim ult As String
'    Dim pri As String
'
'    HayLigaduraAutomatica = False
'
'    If ultimaSilaba = "" Or primeraSilaba = "" Then Exit Function
'
'    ult = Right$(Trim$(ultimaSilaba), 1)
'    pri = Left$(Trim$(primeraSilaba), 1)
'
'    ' Vocal + vocal
'    If ult Like "[aeiouáéíóúü]" And pri Like "[aeiouáéíóúü]" Then
'        HayLigaduraAutomatica = True
'        Exit Function
'    End If
'
'    ' Vocal + h + vocal
'    If ult Like "[aeiouáéíóúü]" And primeraSilaba Like "h[aeiouáéíóúü]*" Then
'        HayLigaduraAutomatica = True
'        Exit Function
'    End If
'
'End Function
'
'
'' ============================================================
''   NORMALITZACIÓ DE VOCALS (placeholder)
'' ============================================================
'Private Function NormalizarVocales(ByVal texto As String) As String
'    NormalizarVocales = texto
'End Function
'
'
'' ============================================================
''   BUSCAR SÍL·LABA REAL ANTERIOR
'' ============================================================
'Private Function BuscarSilabaRealAnterior(arr As Variant, pos As Long) As String
'    Dim j As Long
'    For j = pos - 1 To 0 Step -1
'        If Trim$(arr(j)) <> "" Then
'            BuscarSilabaRealAnterior = Trim$(arr(j))
'            Exit Function
'        End If
'    Next j
'    BuscarSilabaRealAnterior = ""
'End Function
'
'
'' ============================================================
''   BUSCAR SÍL·LABA REAL POSTERIOR
'' ============================================================
'Private Function BuscarSilabaRealPosterior(arr As Variant, pos As Long) As String
'    Dim j As Long
'    For j = pos + 1 To UBound(arr)
'        If Trim$(arr(j)) <> "" Then
'            BuscarSilabaRealPosterior = Trim$(arr(j))
'            Exit Function
'        End If
'    Next j
'    BuscarSilabaRealPosterior = ""
'End Function
'
'
'' ============================================================
''   OBTENIR IPA (amb cache i control de 255)
'' ============================================================
'Private Function ObtenerIPA(ByVal idFonema As Long) As String
'    Dim rs As DAO.Recordset
'
'    If IPA_Cache Is Nothing Then
'        Set IPA_Cache = CreateObject("Scripting.Dictionary")
'    End If
'
'    ' Si ja està en cache
'    If IPA_Cache.Exists(idFonema) Then
'        ObtenerIPA = IPA_Cache(idFonema)
'        Exit Function
'    End If
'
'    ' 255 ? buit
'    If idFonema = 255 Then
'        IPA_Cache.Add idFonema, ""
'        ObtenerIPA = ""
'        Exit Function
'    End If
'
'    ' Consulta a la BD
'    Set rs = CurrentDb.OpenRecordset( _
'        "SELECT IPA FROM qryFonemasValor WHERE ID=" & idFonema & ";", _
'        dbOpenSnapshot)
'
'    If Not (rs.EOF And rs.BOF) Then
'        IPA_Cache.Add idFonema, Nz(rs!ipa, "")
'        ObtenerIPA = Nz(rs!ipa, "")
'    Else
'        IPA_Cache.Add idFonema, ""
'        ObtenerIPA = ""
'    End If
'
'    rs.Close
'    Set rs = Nothing
'End Function

