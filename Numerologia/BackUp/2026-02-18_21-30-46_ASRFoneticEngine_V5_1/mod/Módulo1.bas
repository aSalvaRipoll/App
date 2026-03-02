Attribute VB_Name = "Módulo1"

Option Compare Database
Option Explicit

'' ============================================================
''   ENTRADA PRINCIPAL DEL MOTOR (ILLAS BALEARS)
'' ============================================================
'' ----------------------------------------------------------------
'' Procedimiento: Entrada_Motor_IB
'' Propósito:     Punto de entrada al motor fonético del español general
'' Tipo proc.:    Function
'' Acceso proc.:  Public
'
'' Parameter Texto (String): Texto que se recibe (nombre o apellido español general)
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
'    Dim Limpia As String
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
'        Limpia = Trim$(palabras(i))
'        If Limpia <> "" Then
'            sil = SilabearPalabra_IB(Limpia)
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
''Function Normalizar_IB(t As String) As String
''    Dim s As String
''
''    ' 1. Pasar a minúsculas
''    s = LCase$(t)
''
''    ' 2. Normalizar apóstrofos
''    s = Replace(s, "’", "'")
''    s = Replace(s, "´", "'")
''    s = Replace(s, "`", "'")
''
''    ' 3. Eliminar espacios duplicados
''    Do While InStr(s, "  ") > 0
''        s = Replace(s, "  ", " ")
''    Loop
''
''    ' 4. Normalizar guiones
''    s = Replace(s, "–", "-")
''    s = Replace(s, "—", "-")
''
''    ' 5. Mantener artículo salat sin elisión
''    ' (No hacemos nada: simplemente NO aplicamos reglas IEC)
''
''    ' 6. Normalizar caracteres especiales
''    s = Replace(s, "á", "á")
''    s = Replace(s, "à", "à")
''    s = Replace(s, "è", "è")
''    s = Replace(s, "é", "é")
''    s = Replace(s, "í", "í")
''    s = Replace(s, "ï", "ï")
''    s = Replace(s, "ò", "ò")
''    s = Replace(s, "ó", "ó")
''    s = Replace(s, "ú", "ú")
''    s = Replace(s, "ü", "ü")
''    s = Replace(s, "ç", "ç")
''
''    ' 7. Quitar espacios al inicio y final
''    s = Trim$(s)
''
''    Normalizar_IB = s
''End Function
'
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
''Function EsVocal_IB(ByVal c As String) As Boolean
''    Select Case c
''        Case "á", "à", "a", _
''             "è", "é", "e", _
''             "í", "ï", "i", _
''             "ò", "ó", "o", _
''             "ú", "ü", "u"
''            EsVocal_IB = True
''        Case Else
''            EsVocal_IB = False
''    End Select
''End Function
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
''Function EsDiptongo_IB(ByVal c1 As String, ByVal c2 As String) As Boolean
''    Dim par As String
''    par = c1 & c2
''
''    Select Case par
''
''        ' --- Diptongos decrecientes ---
''        Case "ai", "ei", "oi", "ui", _
''             "au", "eu", "ou"
''            EsDiptongo_IB = True
''            Exit Function
''
''        ' --- Diptongos crecientes ---
''        Case "ia", "ie", "io", "iu", _
''             "ua", "ue", "uo", "ui"
''            EsDiptongo_IB = True
''            Exit Function
''
''        ' --- Diptongos con dièresi ---
''        Case "üa", "üe", "üi", "üo"
''            EsDiptongo_IB = True
''            Exit Function
''
''        ' --- Secuencias que NO son diptongo ---
''        Case "aï", "eï", "oï", "uï", _
''             "aü", "eü", "oü", _
''             "qü", "qüe", "qüi", "qüo"
''            EsDiptongo_IB = False
''            Exit Function
''
''        ' Por defecto: no es diptongo
''        Case Else
''            EsDiptongo_IB = False
''
''    End Select
''End Function
'
'Function EsHiat_IB(c1 As String, c2 As String) As Boolean
'    If EsVocal_IB(c1) And EsVocal_IB(c2) Then
'        EsHiat_IB = Not EsDiptong_IB(c1, c2)
'    Else
'        EsHiat_IB = False
'    End If
'End Function
'
''Function EsHiato_IB(ByVal c1 As String, ByVal c2 As String) As Boolean
''    ' En mallorquín: si no es diptongo, es hiato
''    If EsVocal_IB(c1) And EsVocal_IB(c2) Then
''        EsHiato_IB = Not EsDiptongo_IB(c1, c2)
''    Else
''        EsHiato_IB = False
''    End If
''End Function



