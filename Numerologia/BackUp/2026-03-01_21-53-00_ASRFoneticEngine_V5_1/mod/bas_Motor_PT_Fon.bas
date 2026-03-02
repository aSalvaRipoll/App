Attribute VB_Name = "bas_Motor_PT_Fon"

'Option Compare Database
'Option Explicit
'
'' ============================================================
''   MOTOR FONÉTICO — PORTUGUÉS EUROPEO (PT-EU)
''   PARTE 1 — Arquitectura base + fonemas simples
''
''   Autor: Alba Salvá Ripoll
''
''   Este módulo convierte sílabas PT-EU en fonemas AFI.
''   La PARTE 2 contiene las reglas contextuales complejas.
'' ============================================================
'' ============================================
'' PUNTO DE ENTRADA AL PROCESADOR FONÉTICO PT
'' ============================================
'
'Public Sub ConstruirCadenaFonemas_PT()
'
'    Dim arrSilabas As Variant
'    Dim i As Long
'    Dim silabaCruda As String
'    Dim ultimaSilaba As String
'    Dim siguienteSilaba As String
'
'    Dim arrFon As Variant
'    Dim Fon As Variant
'    Dim res As String
'    Dim ligaduraID As String
'
'    ObjDTO.IdsFonemas = ""
'    ObjDTO.FonemasFinal = ""
'
'    ' En PT usamos las sílabas acentuadas tal cual
'    arrSilabas = Split(ObjDTO.SilabasFinal, "|")
'
'    For i = 0 To UBound(arrSilabas)
'
'        silabaCruda = Trim$(arrSilabas(i))
'
'        ' 1. Separador de palabra
'        If silabaCruda = "" Then
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
'
'        End If
'
'        ' 2. Ligadura manual
'        ligaduraID = DetectarLigaduraManual(silabaCruda)
'        If ligaduraID <> 0 Then
'            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & CStr(ligaduraID) & ","
'        End If
'
'        ' 3. Acento (igual que GL)
'        If InStr(silabaCruda, "(") > 0 Then
'            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "80,"
'        ElseIf InStr(silabaCruda, "[") > 0 Then
'            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "81,"
'        Else
'            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "82,"
'        End If
'
'        silabaCruda = Replace(silabaCruda, "(", "")
'        silabaCruda = Replace(silabaCruda, ")", "")
'        silabaCruda = Replace(silabaCruda, "[", "")
'        silabaCruda = Replace(silabaCruda, "]", "")
'
'        ' 4. Procesar grafemas (versión PT)
'        ProcesarSilaba_PT silabaCruda
'
'        ' 5. Separador silábico
'        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "|"
'
'SiguienteIteracion:
'    Next i
'
'    ' 6. Limpieza final (igual que GL)
'    ObjDTO.IdsFonemas = Replace(ObjDTO.IdsFonemas, "|#", "#")
'    ObjDTO.IdsFonemas = Replace(ObjDTO.IdsFonemas, "#|", "#")
'    ObjDTO.IdsFonemas = Replace(ObjDTO.IdsFonemas, ",|", "|")
'    ObjDTO.IdsFonemas = Replace(ObjDTO.IdsFonemas, "84,84,", "84,")
'
'    While Right$(ObjDTO.IdsFonemas, 1) = "|" Or Right$(ObjDTO.IdsFonemas, 1) = "#" Or Right$(ObjDTO.IdsFonemas, 1) = ","
'        ObjDTO.IdsFonemas = Left$(ObjDTO.IdsFonemas, Len(ObjDTO.IdsFonemas) - 1)
'    Wend
'
'    ' 7. Generar IPA usando tu misma lógica GL
'    ObjDTO.FonemasFinal = GenerarIPA()
'
'End Sub
'
'Private Sub ProcesarSilaba_PT(ByVal silabaCruda As String)
'
'    Dim silaba As String
'    Dim ligaduraID As Byte
'    Dim i As Long
'    Dim grafema As String
'    Dim id As Byte
'    Dim antCh As String
'    Dim sigCh As String
'
'    silaba = Trim$(silabaCruda)
'
'    ' 1. Normalizar grafemas PT (h muda, variantes)
'    silaba = NormalizarGrafema_PT(silaba)
'
'    ' 2. Ligadura manual
'    ligaduraID = DetectarLigaduraManual(silaba)
'    If ligaduraID <> 0 Then
'        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & CStr(ligaduraID) & ","
'    End If
'
'    ' 0. Trígrafos nasales
'If Len(silaba) - i + 1 >= 3 Then
'    If ProcesarTrigramaNasal_PT(Mid$(silaba, i, 3)) Then
'        i = i + 3
'        GoTo SiguienteGrafema
'    End If
'End If
'
'    ' 3. Recorrido de grafemas (1–2 letras, nasalizaciones, dígrafos PT)
'    i = 1
'    Do While i <= Len(silaba)
'
'        grafema = DetectarDigrafo_PT(silaba, i)
'
'        If i > 1 Then
'            antCh = Mid$(silaba, i - 1, 1)
'        Else
'            antCh = ""
'        End If
'
'        If i + Len(grafema) <= Len(silaba) Then
'            sigCh = Mid$(silaba, i + Len(grafema), 1)
'        Else
'            sigCh = ""
'        End If
'
'' 1. Reducción vocálica
'If ProcesarReduccionVocal_PT(grafema, (ObjDTO.IdsFonemas Like "*82,*")) Then
'    i = i + Len(grafema)
'    GoTo SiguienteGrafema
'End If
'
'' 2. Procesar X
'If grafema = "x" Then
'    If ProcesarX_PT(antCh, sigCh) Then
'        i = i + 1
'        GoTo SiguienteGrafema
'    End If
'End If
'
'' 3. Procesar S
'If grafema = "s" Then
'    If ProcesarS_PT(antCh, sigCh) Then
'        i = i + 1
'        GoTo SiguienteGrafema
'    End If
'End If
'
'        id = AsignarFonemaBase_PT(grafema, sigCh, antCh)
'
'        If id <> 0 Then
'            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & CStr(id) & ","
'        End If
'
'        i = i + Len(grafema)
'
'SiguienteGrafema:
'
'    Loop
'
'End Sub
'
'Private Function NormalizarGrafema_PT(ByVal texto As String) As String
'
''    ' Vocales abiertas alternativas ? estándar
''    texto = Replace(texto, "è", "ê")
''    texto = Replace(texto, "ë", "ê")
''
''    texto = Replace(texto, "ò", "ô")
''    texto = Replace(texto, "ö", "ô")
'
'    ' H muda
'    texto = Replace(texto, "h", "")
'
'    NormalizarGrafema_PT = texto
'
'End Function
'
'Private Function DetectarDigrafo_PT(texto As String, pos As Long) As String
'
'    Dim L As Long
'    Dim g3 As String, g2 As String, g1 As String
'
'    L = Len(texto)
'
'    ' ============================================================
'    ' 1) TRÍGRAFOS PT (solo detección; PARTE 2 los procesa)
'    ' ============================================================
'    If pos + 2 <= L Then
'        g3 = Mid$(texto, pos, 3)
'
'        Select Case g3
'            Case "ão", "õe", "ãe"
'                DetectarDigrafo_PT = g3
'                Exit Function
'        End Select
'    End If
'
'    ' ============================================================
'    ' 2) DÍGRAFOS PT
'    ' ============================================================
'    If pos + 1 <= L Then
'        g2 = Mid$(texto, pos, 2)
'
'        ' Dígrafos inseparables PT-EU
'        Select Case g2
'            Case "nh", "lh", "ch", "rr", "ss"
'                DetectarDigrafo_PT = g2
'                Exit Function
'        End Select
'
'        ' gu (antes de e/i o no — PARTE 1 ya lo resuelve)
'        If g2 = "gu" Then
'            DetectarDigrafo_PT = "gu"
'            Exit Function
'        End If
'
'        ' qu (antes de e/i o no — PARTE 1 ya lo resuelve)
'        If g2 = "qu" Then
'            DetectarDigrafo_PT = "qu"
'            Exit Function
'        End If
'
'        ' Nasales finales (solo si NO va seguida de vocal)
'        Select Case g2
'            Case "am", "an", "em", "en", "im", "in", "om", "on", "um", "un"
'
'                ' Final de palabra ? nasal
'                If pos + 2 > L Then
'                    DetectarDigrafo_PT = g2
'                    Exit Function
'                End If
'
'                ' Si NO va seguida de vocal ? nasal
'                If Not (Mid$(texto, pos + 2, 1) Like "[aeiouáéíóúâêôãõ]") Then
'                    DetectarDigrafo_PT = g2
'                    Exit Function
'                End If
'
'        End Select
'    End If
'
'    ' ============================================================
'    ' 3) GRAFEMA SIMPLE
'    ' ============================================================
'    g1 = Mid$(texto, pos, 1)
'    DetectarDigrafo_PT = g1
'
'End Function
'
''Private Function DetectarDigrafo_PT(texto As String, pos As Long) As String
''
''    Dim L As Long
''    Dim g3 As String, g2 As String, g1 As String
''
''    L = Len(texto)
''
''    ' ============================
''    ' 1) TRÍGRAFOS PT
''    ' ============================
''    If pos + 2 <= L Then
''        g3 = Mid$(texto, pos, 3)
''
''        Select Case g3
''            Case "ão", "õe", "ãe"
''                DetectarDigrafo_PT = g3
''                Exit Function
''        End Select
''    End If
''
''    ' ============================
''    ' 2) DÍGRAFOS PT
''    ' ============================
''    If pos + 1 <= L Then
''        g2 = Mid$(texto, pos, 2)
''
''        Select Case g2
''            Case "nh", "lh", "ch", "rr", "ss", "ç"
''                DetectarDigrafo_PT = g2
''                Exit Function
''        End Select
''
''        ' gu + vocal
''        If g2 = "gu" Then
''            DetectarDigrafo_PT = "gu"
''            Exit Function
''        End If
''
''        ' qu + vocal
''        If g2 = "qu" Then
''            DetectarDigrafo_PT = "qu"
''            Exit Function
''        End If
''
''        ' Nasales finales
''        Select Case g2
''            Case "am", "an", "em", "en", "im", "in", "om", "on", "um", "un"
''                If pos + 2 > L Then
''                    DetectarDigrafo_PT = g2
''                    Exit Function
''                End If
''                If Not (Mid$(texto, pos + 2, 1) Like "[aeiouáéíóúâêôãõ]") Then
''                    DetectarDigrafo_PT = g2
''                    Exit Function
''                End If
''        End Select
''    End If
''
''    ' ============================
''    ' 3) GRAFEMA SIMPLE
''    ' ============================
''    g1 = Mid$(texto, pos, 1)
''    DetectarDigrafo_PT = g1
''
''End Function
'
''Private Function DetectarDigrafo_PT(texto As String, pos As Long) As String
''
''    Dim par As String
''    Dim sig As String
''
''    ' Última letra ? 1 grafema
''    If pos + 1 > Len(texto) Then
''        DetectarDigrafo_PT = Mid$(texto, pos, 1)
''        Exit Function
''    End If
''
''    par = Mid$(texto, pos, 2)
''
''    If pos + 2 <= Len(texto) Then
''        sig = Mid$(texto, pos + 2, 1)
''    Else
''        sig = ""
''    End If
''
''    ' Dígrafos PT inseparables
''    Select Case par
''        Case "nh", "lh", "ch", "rr", "qu"
''            DetectarDigrafo_PT = par
''            Exit Function
''    End Select
''
''    ' GU + e/i ? /g/
''    If par = "gu" And (sig = "e" Or sig = "i") Then
''        DetectarDigrafo_PT = "gu"
''        Exit Function
''    End If
''
''    ' Nasales vocálicas simples (final de sílaba/palabra)
''    ' am, an, em, en, im, in, om, on, um, un
''    If par = "am" Or par = "an" Or _
''       par = "em" Or par = "en" Or _
''       par = "im" Or par = "in" Or _
''       par = "om" Or par = "on" Or _
''       par = "um" Or par = "un" Then
''
''        ' Si después no viene vocal ? nasal vocálica
''        If sig = "" Or Not (sig Like "[aeiouáéíóúâêôãõ]") Then
''            DetectarDigrafo_PT = par
''            Exit Function
''        End If
''    End If
''
''    ' Si no es dígrafo ? 1 letra
''    DetectarDigrafo_PT = Mid$(texto, pos, 1)
''
''End Function
'
'Private Function AsignarFonemaBase_PT(grafema As String, _
'                             Optional ByVal sig As String = "", _
'                             Optional ByVal ant As String = "") As Byte
'
'    '===========================================================
'    ' SEMIVOCALES /w/ y /j/
'    '===========================================================
'
'    ' /w/ ? u + vocal fuerte
'    If grafema = "u" And (sig Like "[aeoáéóâêôãõ]") Then
'        AsignarFonemaBase_PT = 22   ' /w/
'        Exit Function
'    End If
'
'    ' /j/ ? i + vocal fuerte (no tónica marcada aquí)
'    If grafema = "i" And (sig Like "[aeouáéóâêôãõ]") Then
'        AsignarFonemaBase_PT = 21   ' /j/
'        Exit Function
'    End If
'
'    '===========================================================
'    ' VOCALES ORALES
'    ' (usamos los mismos IDs básicos que GL)
'    '===========================================================
'
'    Select Case grafema
'        Case "a", "á", "â": AsignarFonemaBase_PT = 1: Exit Function
'        Case "e", "é", "ê": AsignarFonemaBase_PT = 2: Exit Function
'        Case "i", "í":      AsignarFonemaBase_PT = 4: Exit Function
'        Case "o", "ó", "ô": AsignarFonemaBase_PT = 5: Exit Function
'        Case "u", "ú":      AsignarFonemaBase_PT = 7: Exit Function
'    End Select
'
'    '===========================================================
'    ' VOCALES NASALES (IDs 14–18 de tu tabla)
'    '===========================================================
'
'    ' ã ? nasal /a/
'    If grafema = "ã" Then
'        AsignarFonemaBase_PT = 14
'        Exit Function
'    End If
'
'    ' õ ? nasal /o/
'    If grafema = "õ" Then
'        AsignarFonemaBase_PT = 17
'        Exit Function
'    End If
'
'    ' am / an ? nasal /a/
'    If grafema = "am" Or grafema = "an" Then
'        AsignarFonemaBase_PT = 14
'        Exit Function
'    End If
'
'    ' em / en ? nasal /e/
'    If grafema = "em" Or grafema = "en" Then
'        AsignarFonemaBase_PT = 15
'        Exit Function
'    End If
'
'    ' im / in ? nasal /i/
'    If grafema = "im" Or grafema = "in" Then
'        AsignarFonemaBase_PT = 16
'        Exit Function
'    End If
'
'    ' om / on ? nasal /o/
'    If grafema = "om" Or grafema = "on" Then
'        AsignarFonemaBase_PT = 17
'        Exit Function
'    End If
'
'    ' um / un ? nasal /u/
'    If grafema = "um" Or grafema = "un" Then
'        AsignarFonemaBase_PT = 18
'        Exit Function
'    End If
'
'    '===========================================================
'    ' OCLUSIVAS
'    '===========================================================
'
'    Select Case grafema
'        Case "p": AsignarFonemaBase_PT = 30: Exit Function
'        Case "b": AsignarFonemaBase_PT = 31: Exit Function
'        Case "t": AsignarFonemaBase_PT = 32: Exit Function
'        Case "d": AsignarFonemaBase_PT = 33: Exit Function
'    End Select
'
'    ' K / C / Q
'    If grafema = "k" Then AsignarFonemaBase_PT = 34: Exit Function
'
'    If grafema = "c" Then
'        If sig = "e" Or sig = "i" Or sig = "é" Or sig = "ê" Then
'            AsignarFonemaBase_PT = 42   ' /s/
'        Else
'            AsignarFonemaBase_PT = 34   ' /k/
'        End If
'        Exit Function
'    End If
'
'    If grafema = "qu" Then
'        AsignarFonemaBase_PT = 34      ' /k/
'        Exit Function
'    End If
'
'    If grafema = "g" Then
'        If sig = "e" Or sig = "i" Or sig = "é" Or sig = "ê" Then
'            AsignarFonemaBase_PT = 45   ' /?/ aprox. (usas ID 45 fricativa palatal sonora)
'        Else
'            AsignarFonemaBase_PT = 35   ' /g/
'        End If
'        Exit Function
'    End If
'
'    If grafema = "gu" Then
'        AsignarFonemaBase_PT = 35      ' /g/
'        Exit Function
'    End If
'
'    '===========================================================
'    ' FRICATIVAS: F, V, S, Z, X
'    '===========================================================
'
'    If grafema = "f" Then
'        AsignarFonemaBase_PT = 40
'        Exit Function
'    End If
'
'    If grafema = "v" Then
'        AsignarFonemaBase_PT = 41
'        Exit Function
'    End If
'
'    ' S: aquí solo base /s/; refinable por contexto si quieres
'    If grafema = "s" Then
'        AsignarFonemaBase_PT = 42
'        Exit Function
'    End If
'
'    ' Z
'    If grafema = "z" Then
'        AsignarFonemaBase_PT = 43
'        Exit Function
'    End If
'
'    ' X portuguesa (simplificación inicial):
'    '   - ex- + vocal ? /z/ (ID 43 ya lo tienes marcado en la tabla)
'    '   - resto ? /?/ (ID 44)
'    If grafema = "x" Then
'        If ant = "e" And (sig Like "[aeiouáéíóúâêôãõ]") Then
'            AsignarFonemaBase_PT = 43   ' /z/
'        Else
'            AsignarFonemaBase_PT = 44   ' /?/
'        End If
'        Exit Function
'    End If
'
'    '===========================================================
'    ' NASALES
'    '===========================================================
'
'    If grafema = "nh" Then
'        AsignarFonemaBase_PT = 38      ' nasal palatal
'        Exit Function
'    End If
'
'    If grafema = "m" Then
'        AsignarFonemaBase_PT = 36
'        Exit Function
'    End If
'
'    If grafema = "n" Then
'        AsignarFonemaBase_PT = 37
'        Exit Function
'    End If
'
'    '===========================================================
'    ' LATERALES
'    '===========================================================
'
'    If grafema = "lh" Then
'        AsignarFonemaBase_PT = 63      ' lateral palatal
'        Exit Function
'    End If
'
'    If grafema = "l" Then
'        AsignarFonemaBase_PT = 62      ' lateral alveolar
'        Exit Function
'    End If
'
'    '===========================================================
'    ' VIBRANTES (simplificación PT-EU)
'    '===========================================================
'
'    If grafema = "rr" Then
'        AsignarFonemaBase_PT = 60      ' usamos vibrante múltiple como /?/
'        Exit Function
'    End If
'
'    If grafema = "r" Then
'        ' Inicio de palabra o después de consonante ? /?/ (ID 60)
'        If ant = "" Or Not (ant Like "[aeiouáéíóúâêôãõ]") Then
'            AsignarFonemaBase_PT = 60
'        Else
'            AsignarFonemaBase_PT = 59   ' intervocálica ? /?/
'        End If
'        Exit Function
'    End If
'
'    '===========================================================
'    ' AFRICADA
'    '===========================================================
'
'    If grafema = "ch" Then
'        AsignarFonemaBase_PT = 57
'        Exit Function
'    End If
'
'    '===========================================================
'    ' H MUDA
'    '===========================================================
'
'    If grafema = "h" Then
'        AsignarFonemaBase_PT = 0
'        Exit Function
'    End If
'
'    '===========================================================
'    ' DESCONOCIDO
'    '===========================================================
'
'    AsignarFonemaBase_PT = 255
'
'End Function
'
'Private Function ProcesarTrigramaNasal_PT(grafema As String) As Boolean
'
'    Select Case grafema
'
'        ' ão ? /?~/ + /w~/ ? 14 + 24
'        Case "ão"
'            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "14,24,"
'            ProcesarTrigramaNasal_PT = True
'            Exit Function
'
'        ' õe ? /õ/ + /j~/ ? 17 + 23
'        Case "õe"
'            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "17,23,"
'            ProcesarTrigramaNasal_PT = True
'            Exit Function
'
'        ' ãe ? /?~/ + /j~/ ? 14 + 23
'        Case "ãe"
'            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "14,23,"
'            ProcesarTrigramaNasal_PT = True
'            Exit Function
'
'    End Select
'
'    ProcesarTrigramaNasal_PT = False
'
'End Function
'Private Function ProcesarX_PT(ant As String, sig As String) As Boolean
'
'    ' ex- + vocal ? /z/
'    If ant = "e" And (sig Like "[aeiouáéíóúâêôãõ]") Then
'        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "43,"
'        ProcesarX_PT = True
'        Exit Function
'    End If
'
'    ' intervocálica ? /z/
'    If (ant Like "[aeiouáéíóúâêôãõ]") And (sig Like "[aeiouáéíóúâêôãõ]") Then
'        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "43,"
'        ProcesarX_PT = True
'        Exit Function
'    End If
'
'    ' antes de consonante sonora ? /z/
'    If sig Like "[bdgvzmnrl]" Then
'        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "43,"
'        ProcesarX_PT = True
'        Exit Function
'    End If
'
'    ' antes de consonante sorda ? /?/
'    If sig Like "[ptkfscx]" Then
'        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "44,"
'        ProcesarX_PT = True
'        Exit Function
'    End If
'
'    ' final de sílaba ? /?/
'    If sig = "" Then
'        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "44,"
'        ProcesarX_PT = True
'        Exit Function
'    End If
'
'    ' /ks/
'    If sig = "c" Or sig = "s" Then
'        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "34,42,"
'        ProcesarX_PT = True
'        Exit Function
'    End If
'
'    ' /gz/
'    If sig = "g" Or sig = "z" Then
'        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "35,43,"
'        ProcesarX_PT = True
'        Exit Function
'    End If
'
'    ' por defecto ? /?/
'    ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "44,"
'    ProcesarX_PT = True
'
'End Function
'
'Private Function ProcesarS_PT(ant As String, sig As String) As Boolean
'
'    ' ss ya viene como dígrafo ? /s/
'    ' ç ya viene como grafema simple ? /s/
'
'    ' intervocálica ? /z/
'    If (ant Like "[aeiouáéíóúâêôãõ]") And (sig Like "[aeiouáéíóúâêôãõ]") Then
'        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "43,"
'        ProcesarS_PT = True
'        Exit Function
'    End If
'
'    ' final de sílaba ? /?/
'    If sig = "" Then
'        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "44,"
'        ProcesarS_PT = True
'        Exit Function
'    End If
'
'    ' antes de consonante sorda ? /?/
'    If sig Like "[ptkfscx]" Then
'        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "44,"
'        ProcesarS_PT = True
'        Exit Function
'    End If
'
'    ' antes de consonante sonora ? /z/
'    If sig Like "[bdgvzmnrl]" Then
'        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "43,"
'        ProcesarS_PT = True
'        Exit Function
'    End If
'
'    ' por defecto ? /s/
'    ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "42,"
'    ProcesarS_PT = True
'
'End Function
'
'Private Function ProcesarReduccionVocal_PT(grafema As String, esAtona As Boolean) As Boolean
'
'    If Not esAtona Then
'        ProcesarReduccionVocal_PT = False
'        Exit Function
'    End If
'
'    Select Case grafema
'
'        ' a átona ? /?/ ? ID 9
'        Case "a"
'            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "9,"
'            ProcesarReduccionVocal_PT = True
'            Exit Function
'
'        ' e átona ? /?/ ? ID 10
'        Case "e"
'            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "10,"
'            ProcesarReduccionVocal_PT = True
'            Exit Function
'
'        ' o átono ? /u/ ? ID 7
'        Case "o"
'            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "7,"
'            ProcesarReduccionVocal_PT = True
'            Exit Function
'
'    End Select
'
'    ProcesarReduccionVocal_PT = False
'
'End Function


