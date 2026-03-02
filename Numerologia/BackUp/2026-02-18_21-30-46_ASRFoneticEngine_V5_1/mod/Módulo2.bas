Attribute VB_Name = "Módulo2"
'Option Compare Database
'Option Explicit
'
'
'Private Const strSQL As String = _
'        "SELECT Prefijo FROM tbmPrefijos " & _
'        "WHERE Activo = 1 " & _
'        "AND Tipo Like 'auténtico' " & _
'        "AND [ca-ib] = True " & _
'        "ORDER BY Len(Prefijo) DESC, Prefijo ASC"
'
'' ============================================================
''   ENTRADA PRINCIPAL DEL MOTOR (ILLAS BALEARS)
'' ============================================================
'' ----------------------------------------------------------------
'' Procedimiento: Entrada_Motor_IB
'' Propósito:     Punto de entrada al motor fonético del Mallorquín
'' Tipo proc.:    Function
'' Acceso proc.:  Public
'
'' Parameter Texto (String): Texto que se recibe (nombre o apellido)
'
'' Tipo retorno: String -> Texto que contiene la lista de fonemas
''   resultado de la conversión
'
'' Autor:        Alba Salvá
'' Fecha:        16/02/2026
'' ----------------------------------------------------------------
'Public Function Entrada_Motor_IB(texto As String) As String
'
'    Set ObjDTO = New clsDTO_Motor
'
''    DebugMotor = True
''    DebugDTO = False
'
'    ' 1) Asignamos el texto recibido y
'    '    Normalización (dentro del DTO)
'    ObjDTO.TextoOriginal = texto
'    ObjDTO.NormalizaEntrada
'
'    ' 2) Silabeo automático
'    Call SilabearFrase_IB
'
'    ' 3) Detectar tónicas
'    Call CalcularTonicas_IB
'
'    ' 4) Detectar secundarias
'    Call CalcularSecundarias_IB
'
'    ' 5) Marcar Tónicas y Secundarias
'    Call MarcarTonicaYSecundariaEnCadena_IB
'
'    ObjDTO.SilabasFinal = ObjDTO.SilabasAcentuadas
'
'    ' 6) Generar fonética
'    Call ConstruirCadenaFonemas_IB
'
'    '    << Eliminar en producción >>
'    Call MF_DebugDTO("Silabear")
'
'    ' 7) Devolver resultado (texto plano)
'    '    << Eliminar en producción >>
'    Entrada_Motor_IB = ObjDTO.SilabasAuto
'
'End Function
'
'Private Sub SilabearFrase_IB()
'
'    Dim frase As String
'    Dim palabras() As String
'    Dim resultado As String
'    Dim i As Long
'    Dim limpia As String
'    Dim sil As String
'
'    usarSilabeoMorfologico = True
'    modoPrefijosEstrictos = True
'    respetarPrefijos = True
'
'    frase = ObjDTO.TextoNormalizado
'
'    palabras = Split(frase, " ")
'
'    For i = LBound(palabras) To UBound(palabras)
'        limpia = Trim$(palabras(i))
'        If limpia <> "" Then
'            sil = SilabearPalabra_IB(limpia)
'            If resultado = "" Then
'                resultado = sil
'            Else
'                resultado = resultado & " |   | " & sil
'            End If
'        End If
'    Next i
'
'    ObjDTO.SilabasAuto = resultado
'
'End Sub
'
'Public Function SilabearPalabra_IB(ByVal texto As String) As String
'    Dim t As String
'
'    t = LCase$(Trim$(texto))
'
'    If usarSilabeoMorfologico Then
'        SilabearPalabra_IB = SilabearMorfologico_IB(t)
'    Else
'        SilabearPalabra_IB = SilabearOrtog_IB(t)
'    End If
'
'End Function
'
'Public Function SilabearOrtog_IB(ByVal t As String) As String
'    Dim nucIni() As Integer, nucFin() As Integer
'    Dim silIni() As Integer, silFin() As Integer
'    Dim nNuc As Integer, i As Integer
'    Dim silabas() As String
'
'    If Len(Trim$(t)) < 2 Then
'        SilabearOrtog_IB = t
'        Exit Function
'    End If
'
'    ' 1. Localizar núcleos mallorquines
'    ReDim nucIni(1 To 200)
'    ReDim nucFin(1 To 200)
'    LocalizarNucleosOrtog_IB t, nucIni, nucFin, nNuc
'
'    ' 2. Calcular sílabas mallorquinas
'    ReDim silIni(1 To nNuc)
'    ReDim silFin(1 To nNuc)
'    CalcularSilabas_IB t, nucIni, nucFin, nNuc, silIni, silFin
'
'    ' 3. Construir sílabas
'    ReDim silabas(1 To nNuc)
'    For i = 1 To nNuc
'        silabas(i) = Mid$(t, silIni(i), silFin(i) - silIni(i) + 1)
'        If DebugMotor Then
'            addLog "Sílaba " & i & ": " & silabas(i)
'        End If
'    Next i
'
'    SilabearOrtog_IB = Join(silabas, " | ")
'End Function
'
'Function SilabearOrtog_IB(ByVal t As String) As String
'    Dim s As String
'    Dim nucIni() As Integer, nucFin() As Integer
'    Dim silIni() As Integer, silFin() As Integer
'    Dim silabas() As String
'    Dim nNuc As Integer
'    Dim i As Integer
'
'    ' 1. Normalitzar ortogràficament
'    s = Normalizar_IB(t)
'
'    ' 2. Redimensionar arrays (màxim 200 nuclis per paraula)
'    ReDim nucIni(1 To 200)
'    ReDim nucFin(1 To 200)
'
'    ' 3. Localitzar nuclis vocàlics mallorquins
'    LocalizarNucleosOrtog_IB s, nucIni, nucFin, nNuc
'
'    ' 4. Redimensionar arrays de síl·labes
'    ReDim silIni(1 To nNuc)
'    ReDim silFin(1 To nNuc)
'
'    ' 5. Calcular síl·labes mallorquines
'    CalcularSilabas_IB s, nucIni, nucFin, nNuc, silIni, silFin
'
'    ' 6. Construir les síl·labes finals
'    ReDim silabas(1 To nNuc)
'    For i = 1 To nNuc
'        silabas(i) = Mid$(s, silIni(i), silFin(i) - silIni(i) + 1)
'    Next i
'
'    ' 7. Retornar síl·labes separades per " | "
'    SilabearOrtog_IB = Join(silabas, " | ")
'End Function
'
'
'Sub LocalizarNucleosOrtog_IB(ByVal t As String, _
'                             ByRef nucIni() As Integer, _
'                             ByRef nucFin() As Integer, _
'                             ByRef nNuc As Integer)
'
'    Dim i As Integer, L As Integer
'    Dim c1 As String, c2 As String
'
'    L = Len(t)
'    nNuc = 0
'    i = 1
'
'    Do While i <= L
'
'        c1 = Mid$(t, i, 1)
'
'        If EsVocal_IB(c1) Then
'
'            ' Possible diptong si hi ha una vocal següent
'            If i < L Then
'                c2 = Mid$(t, i + 1, 1)
'
'                If EsVocal_IB(c2) Then
'                    ' Diptong mallorquí?
'                    If EsDiptongo_IB(c1, c2) Then
'                        nNuc = nNuc + 1
'                        nucIni(nNuc) = i
'                        nucFin(nNuc) = i + 1
'                        i = i + 2
'                        GoTo Següent
'                    End If
'                End If
'            End If
'
'            ' Si no hi ha diptong ? vocal sola (hiat o vocal simple)
'            nNuc = nNuc + 1
'            nucIni(nNuc) = i
'            nucFin(nNuc) = i
'            i = i + 1
'            GoTo Següent
'
'        End If
'
'        ' No és vocal ? continuar
'        i = i + 1
'
'Següent:
'    Loop
'
'End Sub
'
'Sub CalcularSilabas_IB(ByVal t As String, _
'                       ByRef nucIni() As Integer, _
'                       ByRef nucFin() As Integer, _
'                       ByVal nNuc As Integer, _
'                       ByRef silIni() As Integer, _
'                       ByRef silFin() As Integer)
'
'    Dim i As Integer, L As Integer
'    Dim a As Integer, b As Integer
'    Dim k As Integer
'    Dim c1 As String, c2 As String, C3 As String, C4 As String
'    Dim grup As String
'
'    L = Len(t)
'    silIni(1) = 1
'
'    For i = 1 To nNuc - 1
'
'        a = nucFin(i)
'        b = nucIni(i + 1)
'
'        k = IIf(b > a + 1, b - a - 1, 0)
'
'        Select Case k
'
'            Case 0
'                silFin(i) = a
'                silIni(i + 1) = a + 1
'
'            Case 1
'                silFin(i) = a
'                silIni(i + 1) = a + 1
'
'            Case 2
'                c1 = Mid$(t, a + 1, 1)
'                c2 = Mid$(t, a + 2, 1)
'                grup = c1 & c2
'
'                ' Dígrafs indivisibles
'                If grup = "rr" Or grup = "ll" Or grup = "ch" Then
'                    silFin(i) = a
'                    silIni(i + 1) = a + 1
'                    GoTo Següent
'                End If
'
'                ' Grups d'atac vàlids
'                If EsGrupoAtaque_IB(grup) Then
'                    silFin(i) = a
'                    silIni(i + 1) = a + 1
'                Else
'                    silFin(i) = a + 1
'                    silIni(i + 1) = a + 2
'                End If
'
'            Case 3
'                c1 = Mid$(t, a + 1, 1)
'                c2 = Mid$(t, a + 2, 1)
'                C3 = Mid$(t, a + 3, 1)
'
'                ' --- REGLA MALLORQUINA: L·L + VOCAL ÉS INDIVISIBLE ---
'                If c1 = "l" And c2 = "·" And C3 = "l" Then
'                    If b <= L Then
'                        C4 = Mid$(t, a + 4, 1)
'                        If EsVocal_IB(C4) Then
'                            ' il·lu ? una sola síl·laba
'                            silFin(i) = a + 4
'                            silIni(i + 1) = a + 5
'                            GoTo Següent
'                        End If
'                    End If
'                End If
'                ' --- FI REGLA MALLORQUINA ---
'
'                ' Regla general
'                If PuedeCerrarSilaba_IB(c2) Then
'                    silFin(i) = a + 2
'                    silIni(i + 1) = a + 3
'                Else
'                    silFin(i) = a + 1
'                    silIni(i + 1) = a + 2
'                End If
'
'            Case Else
'                silFin(i) = a + 2
'                silIni(i + 1) = a + 3
'
'        End Select
'
'Següent:
'    Next i
'
'    silFin(nNuc) = L
'
'End Sub
'
'
'Function Normalizar_IB(t As String) As String
'    Dim s As String
'
'    ' 1. Pasar a minúsculas (no es necesario, ya lo hace el DTO
'    's = LCase$(t)
'
'    ' 2. Normalizar apóstrofos “raros” a apóstrofo simple
'    s = Replace(s, "’", "'")
'    s = Replace(s, "´", "'")
'    s = Replace(s, "`", "'")
'
'    ' 3. NO tocar vocales propias balear/mallorquinas:
'    '    á, à, a, è, é, e, í, ï, i, ò, ó, o, ú, ü, u, ç
'    '    ? todas son legítimas y se conservan tal cual.
'    '
'    '    Por tanto, aquí NO hacemos Replace de vocales.
'    '    (Nada de "á"?"a", ni "á"?"à", etc.)
'
'    ' 4. NO tocar artículo salat ni apócopes:
'    '    es, sa, ses, son, s', d', n', can', ca', ca' s'avi, can' toni, etc.
'    '    ? se conservan exactamente como están.
'
'    ' 5. Normalizar guiones tipográficos a guion simple
'    s = Replace(s, "–", "-")
'    s = Replace(s, "—", "-")
'
'    ' 6. Normalizar espacios múltiples a un solo espacio
'    Do While InStr(s, "  ") > 0
'        s = Replace(s, "  ", " ")
'    Loop
'
'    ' 7. Quitar espacios al inicio y al final
'    s = Trim$(s)
'
'    Normalizar_IB = s
'End Function
'
'Function EsVocal_IB(c As String) As Boolean
'    Select Case c
'        Case "a", "à", "á", _
'             "e", "è", "é", _
'             "i", "í", "ï", _
'             "o", "ò", "ó", _
'             "u", "ú", "ü"
'            EsVocal_IB = True
'        Case Else
'            EsVocal_IB = False
'    End Select
'End Function
'
'Function EsVocalForta_IB(c As String) As Boolean
'    Select Case c
'        Case "a", "à", "á", _
'             "e", "è", "é", _
'             "o", "ò", "ó"
'            EsVocalForta_IB = True
'        Case Else
'            EsVocalForta_IB = False
'    End Select
'End Function
'
'Function EsVocalFeble_IB(c As String) As Boolean
'    Select Case c
'        Case "i", "í", "ï", _
'             "u", "ú", "ü"
'            EsVocalFeble_IB = True
'        Case Else
'            EsVocalFeble_IB = False
'    End Select
'End Function
'
'Function EsSemivocal_IB(c As String) As Boolean
'    Select Case c
'        Case "i", "í", "ï", _
'             "u", "ú", "ü"
'            EsSemivocal_IB = True
'        Case Else
'            EsSemivocal_IB = False
'    End Select
'End Function
'
'Function EsDiptong_IB(c1 As String, c2 As String) As Boolean
'    Dim par As String
'    par = c1 & c2
'
'    ' --- Secuencias explícitamente NO diptongo ---
'    Select Case par
'        Case "aï", "eï", "oï", "uï", _
'             "aü", "eü", "oü", _
'             "qü", "qüe", "qüi", "qüo"
'            EsDiptong_IB = False
'            Exit Function
'    End Select
'
'    ' --- Diptongos decrecientes mallorquines ---
'    Select Case par
'        Case "ai", "ei", "oi", "ui", _
'             "au", "eu", "ou"
'            EsDiptong_IB = True
'            Exit Function
'    End Select
'
'    ' --- Diptongos crecientes mallorquines ---
'    Select Case par
'        Case "ia", "ie", "io", "iu", _
'             "ua", "ue", "uo", "ui"
'            EsDiptong_IB = True
'            Exit Function
'    End Select
'
'    ' --- Diptongos con dièresi ---
'    Select Case par
'        Case "üa", "üe", "üi", "üo"
'            EsDiptong_IB = True
'            Exit Function
'    End Select
'
'    ' Por defecto: no es diptongo
'    EsDiptong_IB = False
'End Function
'
'Function EsHiat_IB(c1 As String, c2 As String) As Boolean
'    If EsVocal_IB(c1) And EsVocal_IB(c2) Then
'        EsHiat_IB = Not EsDiptong_IB(c1, c2)
'    Else
'        EsHiat_IB = False
'    End If
'End Function
'
'Private prefijosEstrictos_IB As Variant
'Private prefijosCargados_IB As Boolean
'
'Private Sub CargarPrefijos_IB()
'    Dim rs As DAO.Recordset
'    Dim sql As String
'    Dim i As Long
'
'    If prefijosCargados_IB Then Exit Sub
'
''    sql = "SELECT Prefijo FROM tbmPrefijos " & _
'          "WHERE Activo = 1 " & _
'          "AND Tipo Like 'auténtico' " & _
'          "AND [ca-ib] = True " & _
'          "ORDER BY Len(Prefijo) DESC, Prefijo ASC"
'
'    sql = strSQL
'
'    Set rs = CurrentDb.OpenRecordset(sql)
'
'    If Not rs.EOF Then
'        rs.MoveLast
'        ReDim prefijosEstrictos_IB(1 To rs.RecordCount)
'        rs.MoveFirst
'
'        i = 1
'        Do Until rs.EOF
'            prefijosEstrictos_IB(i) = LCase$(rs!Prefijo)
'            i = i + 1
'            rs.MoveNext
'        Loop
'    End If
'
'    rs.Close
'    prefijosCargados_IB = True
'End Sub
'
'Private Function DetectarPrefijo_IB(ByVal t As String) As String
'    Dim p As Variant
'
'    If Not prefijosCargados_IB Then CargarPrefijos_IB
'
'    For Each p In prefijosEstrictos_IB
'        If Len(t) > Len(p) Then
'            If Left$(t, Len(p)) = p Then
'                DetectarPrefijo_IB = p
'                Exit Function
'            End If
'        End If
'    Next p
'
'    DetectarPrefijo_IB = ""
'End Function
'
'Public Function SilabearMorfologico_IB(ByVal t As String) As String
'    Dim pref As String
'    Dim resto As String
'
'    If Not respetarPrefijos Then
'        SilabearMorfologico_IB = SilabearOrtog_IB(t)
'        Exit Function
'    End If
'
'    pref = DetectarPrefijo_IB(t)
'
'    If pref = "" Then
'        SilabearMorfologico_IB = SilabearOrtog_IB(t)
'        Exit Function
'    End If
'
'    resto = Mid$(t, Len(pref) + 1)
'
'    SilabearMorfologico_IB = pref & " | " & SilabearOrtog_IB(resto)
'End Function
'
'Public Sub ConstruirCadenaFonemas_IB()
'
'    Dim sils As Variant
'    Dim i As Long
'    Dim fonemaSils() As String
'
'    sils = Split(ObjDTO.SilabasFinal, " | ")
'
'    ReDim fonemaSils(LBound(sils) To UBound(sils))
'
'    For i = LBound(sils) To UBound(sils)
'        fonemaSils(i) = ConvertirSilaba_IB(Trim$(sils(i)))
'    Next i
'
'    ' Aplicar processos fonètics globals
'    Dim cadena As String
'    cadena = Join(fonemaSils, " . ")
'
'    cadena = AplicarAssimilacions_IB(cadena)
'    cadena = AplicarReduccions_IB(cadena)
'    cadena = AplicarSchwa_IB(cadena)
'
'    ObjDTO.Fonemas = cadena
'End Sub
'
'Private Function ConvertirSilaba_IB(sil As String) As String
'    Dim out As String
'    Dim i As Long, c As String
'
'    i = 1
'    Do While i <= Len(sil)
'        c = Mid$(sil, i, 1)
'
'        ' Diftongs
'        If i < Len(sil) Then
'            Dim c2 As String
'            c2 = Mid$(sil, i + 1, 1)
'
'            If EsDiptong_IB(c, c2) Then
'                out = out & ConvertirDiptong_IB(c, c2)
'                i = i + 2
'                GoTo Següent
'            End If
'        End If
'
'        ' Consonants dobles
'        If i < Len(sil) Then
'            Dim grup As String
'            grup = Mid$(sil, i, 2)
'
'            If EsGrupConsonantic_IB(grup) Then
'                out = out & ConvertirGrup_IB(grup)
'                i = i + 2
'                GoTo Següent
'            End If
'        End If
'
'        ' Vocal sola
'        If EsVocal_IB(c) Then
'            out = out & ConvertirVocal_IB(c)
'        Else
'            out = out & ConvertirConsonant_IB(c)
'        End If
'
'Següent:
'        i = i + 1
'    Loop
'
'    ConvertirSilaba_IB = out
'End Function
'
'Private Function ConvertirVocal_IB(v As String) As String
'    Select Case v
'        Case "a", "à", "á"
'            ConvertirVocal_IB = "1"   ' /a/
'
'        Case "e", "é"
'            ConvertirVocal_IB = "2"   ' /e/
'
'        Case "è"
'            ConvertirVocal_IB = "3"   ' /?/
'
'        Case "i", "í", "ï"
'            ConvertirVocal_IB = "4"   ' /i/
'
'        Case "o", "ó"
'            ConvertirVocal_IB = "5"   ' /o/
'
'        Case "ò"
'            ConvertirVocal_IB = "6"   ' /?/
'
'        Case "u", "ú", "ü"
'            ConvertirVocal_IB = "7"   ' /u/
'
'        Case Else
'            ConvertirVocal_IB = "0"
'    End Select
'End Function
'
'Private Function ConvertirConsonant_IB(c As String) As String
'    Select Case c
'        Case "p": ConvertirConsonant_IB = "30"
'        Case "b": ConvertirConsonant_IB = "31"
'        Case "t": ConvertirConsonant_IB = "32"
'        Case "d": ConvertirConsonant_IB = "33"
'
'        Case "k", "c", "q"
'            ConvertirConsonant_IB = "34"   ' /k/
'
'        Case "g"
'            ConvertirConsonant_IB = "35"   ' /g/
'
'        Case "f": ConvertirConsonant_IB = "40"
'        Case "v": ConvertirConsonant_IB = "41"
'
'        Case "s": ConvertirConsonant_IB = "42"
'        Case "z": ConvertirConsonant_IB = "43"
'
'        Case "m": ConvertirConsonant_IB = "36"
'        Case "n": ConvertirConsonant_IB = "37"
'        Case "l": ConvertirConsonant_IB = "62"
'
'        Case "r"
'            ConvertirConsonant_IB = "59"   ' /?/
'
'        Case Else
'            ConvertirConsonant_IB = "0"
'    End Select
'End Function
'
'Private Function ConvertirGrup_IB(g As String) As String
'    Select Case g
'        Case "ny": ConvertirGrup_IB = "38"   ' /?/
'        Case "ll": ConvertirGrup_IB = "63"   ' /?/
'        Case "rr": ConvertirGrup_IB = "60"   ' /r/
'        Case "ss": ConvertirGrup_IB = "42"   ' /s/
'
'        Case "tx": ConvertirGrup_IB = "57"   ' /t??/
'        Case "tg", "tj": ConvertirGrup_IB = "58"   ' /d??/
'
'        Case "ts", "tz": ConvertirGrup_IB = "46"   ' /t?s/
'
'        Case "ix": ConvertirGrup_IB = "44"   ' /?/
'
'        Case Else
'            ConvertirGrup_IB = "0"
'    End Select
'End Function
'
'Private Function AplicarSchwa_IB(cadena As String) As String
'
'    ' article salat
'    cadena = Replace(cadena, "es ", "8 42 ")     ' ? + s
'    cadena = Replace(cadena, "sa ", "42 8 ")     ' s + ?
'    cadena = Replace(cadena, "ses ", "42 8 42 ") ' s ? s
'
'    ' apòcopes
'    cadena = Replace(cadena, "can' ", "34 8 37 ") ' k ? n
'    cadena = Replace(cadena, "ca' ", "34 8 ")     ' k ?
'
'    AplicarSchwa_IB = cadena
'End Function
'
'
'
'Private Function AplicarAssimilacions_IB(cadena As String) As String
'
'    ' s + vocal ? z
'    cadena = Replace(cadena, "s a", "z a")
'    cadena = Replace(cadena, "s e", "z e")
'    cadena = Replace(cadena, "s i", "z i")
'    cadena = Replace(cadena, "s o", "z o")
'    cadena = Replace(cadena, "s u", "z u")
'
'    AplicarAssimilacions_IB = cadena
'End Function
'
'Private Function AplicarReduccions_IB(cadena As String) As String
'
'    ' eliminar dobles espais
'    Do While InStr(cadena, "  ") > 0
'        cadena = Replace(cadena, "  ", " ")
'    Loop
'
'    AplicarReduccions_IB = Trim$(cadena)
'End Function
'
'Private Function EsGrupConsonantic_IB(g As String) As Boolean
'    Select Case g
'        Case "ny", "ll", "rr", "ss", "tx", "tg", "tj", "ts", "tz", "ix"
'            EsGrupConsonantic_IB = True
'        Case Else
'            EsGrupConsonantic_IB = False
'    End Select
'End Function
'
''=================================================================
''=================================================================
''                 SECCIÓN MÓDULO ACENTOS
''=================================================================
''=================================================================
'
'
'
'
'
'
'Private Function TieneTilde_IB(ByVal silaba As String) As Boolean
'    TieneTilde_IB = (InStr(silaba, "à") > 0 Or _
'                     InStr(silaba, "á") > 0 Or _
'                     InStr(silaba, "è") > 0 Or _
'                     InStr(silaba, "é") > 0 Or _
'                     InStr(silaba, "í") > 0 Or _
'                     InStr(silaba, "ï") > 0 Or _
'                     InStr(silaba, "ò") > 0 Or _
'                     InStr(silaba, "ó") > 0 Or _
'                     InStr(silaba, "ú") > 0)
'End Function
'
'Private Function DetectarTonica_IB(w As Collection) As Byte
'    Dim i As Byte
'    Dim paraula As String
'    Dim ultima As String
'
'    ' 1) Si alguna síl·laba té accent gràfic ? tònica directa
'    For i = 1 To w.count
'        If TieneTilde_IB(w(i)) Then
'            DetectarTonica_IB = i
'            Exit Function
'        End If
'    Next i
'
'    ' 2) Sense accent: regla general balear ? AGUDA
'    DetectarTonica_IB = w.count
'End Function
'
'Private Sub CalcularTonicas_IB()
'
'    Dim tGlobal As New Collection
'    Dim elements As Collection
'    Set elements = ObtenerPalabrasDesdeSilabasAuto_IB() ' reutilitzem la mateixa funció
'
'    Dim globalIndex As Long
'    Dim i As Long
'
'    globalIndex = 0
'
'    For i = 1 To elements.count
'
'        If TypeName(elements(i)) = "Collection" Then
'            Dim w As Collection
'            Set w = elements(i)
'
'            Dim tLocal As Long
'            tLocal = DetectarTonica_IB(w)
'
'            If tLocal > 0 Then
'                tGlobal.Add globalIndex + tLocal
'            End If
'
'            globalIndex = globalIndex + w.count
'
'        Else
'            globalIndex = globalIndex + 1
'        End If
'
'    Next i
'
'    ObjDTO.SilabasTonicas = JoinCollection_IB(tGlobal)
'End Sub
'
'Private Function DetectarSecundarias_IB(w As Collection, tPos As Byte) As Collection
'    Dim secs As New Collection
'    Dim n As Byte
'    Dim pos2 As Byte
'
'    n = w.count
'
'    If n < 4 Then
'        Set DetectarSecundarias_IB = secs
'        Exit Function
'    End If
'
'    ' Primera secundària sempre a la 1
'    secs.Add 1
'
'    ' Paraules llargues ? segona secundària
'    If n >= 6 Then
'        pos2 = tPos - 2
'        If pos2 > 1 Then secs.Add pos2
'    End If
'
'    Set DetectarSecundarias_IB = secs
'End Function
'
'Private Sub CalcularSecundarias_IB()
'
'    Dim sGlobal As New Collection
'    Dim elements As Collection
'    Set elements = ObtenerPalabrasDesdeSilabasAuto_IB()
'
'    Dim globalIndex As Long
'    Dim i As Long
'    Dim tLocal As Byte
'
'    globalIndex = 0
'
'    For i = 1 To elements.count
'
'        If TypeName(elements(i)) = "Collection" Then
'
'            Dim w As Collection
'            Set w = elements(i)
'
'            tLocal = DetectarTonica_IB(w)
'
'            Dim secs As Collection
'            Set secs = DetectarSecundarias_IB(w, tLocal)
'
'            Dim x As Variant
'            For Each x In secs
'                sGlobal.Add globalIndex + CByte(x)
'            Next x
'
'            globalIndex = globalIndex + w.count
'
'        Else
'            globalIndex = globalIndex + 1
'        End If
'
'    Next i
'
'    ObjDTO.SilabasSecundarias = JoinCollection_IB(sGlobal)
'
'End Sub
'
'Private Sub MarcarTonicaYSecundariaEnCadena_IB()
'
'    Dim sils As Variant
'    Dim i As Long
'    Dim out() As String
'
'    sils = Split(ObjDTO.SilabasAuto, " | ")
'    ReDim out(LBound(sils) To UBound(sils))
'
'    For i = LBound(sils) To UBound(sils)
'        out(i) = sils(i)
'    Next i
'
'    ' Tòniques
'    If ObjDTO.SilabasTonicas <> "" Then
'        Dim t As Variant, x As Variant
'        t = Split(ObjDTO.SilabasTonicas, ",")
'        For Each x In t
'            Dim idx As Long
'            idx = CLng(x) - 1
'            If idx >= LBound(out) And idx <= UBound(out) Then
'                out(idx) = "( " & out(idx) & " )"
'            End If
'        Next x
'    End If
'
'    ' Secundàries
'    If ObjDTO.SilabasSecundarias <> "" Then
'        Dim s As Variant, y As Variant
'        s = Split(ObjDTO.SilabasSecundarias, ",")
'        For Each y In s
'            Dim idx2 As Long
'            idx2 = CLng(y) - 1
'            If idx2 >= LBound(out) And idx2 <= UBound(out) Then
'                out(idx2) = "[ " & out(idx2) & " ]"
'            End If
'        Next y
'    End If
'
'    ObjDTO.SilabasAcentuadas = Join(out, " | ")
'
'End Sub
'
'
'' ============================================================
''   JOIN COLLECTION
'' ============================================================
'Private Function JoinCollection_IB(col As Collection) As String
'
'    Dim arr() As String
'    Dim i As Byte
'
'    If col Is Nothing Then
'        JoinCollection_IB = ""
'        Exit Function
'    End If
'
'    If col.count = 0 Then
'        JoinCollection_IB = ""
'        Exit Function
'    End If
'
