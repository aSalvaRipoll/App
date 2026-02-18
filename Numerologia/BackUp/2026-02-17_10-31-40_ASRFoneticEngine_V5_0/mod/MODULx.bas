Attribute VB_Name = "MODULx"

'Option Compare Database
'Option Explicit
'
'' ============================================================
''   MOTOR DE SILABEO CATALÁN 1.0 — ORTOGRÁFICO + ACENTUACIÓN
''   Autor: Alba Salvá Ripoll + Copilot
''
''   Implementa:
''   · Diptongs catalanes
''   · Triftongs catalanes
''   · Hiats catalanes (aï, eï, oï, aü, eü, oü…)
''   · Vocal neutra (?)
''   · Dígrafos: ny, ll, tg, tj, tx, ix, ss
''   · Apóstrofos (l’, d’, m’, s’, n’)
''   · Acentuación catalana (aguda / plana / esdrúixola)
'' ============================================================
'
'Private usarSilabeoMorfologico As Boolean
'Private modoPrefijosEstrictos As Boolean
'Private respetarPrefijos As Boolean
'Private prefijosCargados As Boolean
'
'Private prefijosEstrictos As Variant
'
'' ============================================================
''   ENTRADA PRINCIPAL DEL MOTOR (CATALÁN)
'' ============================================================
'Public Function Entrada_Motor_CA(Texto As String) As String
'
'    Set ObjDTO = New clsDTO_Motor
'
'    ObjDTO.TextoOriginal = Texto
'    ObjDTO.NormalizaEntrada
'
'    Call SilabearFrase_CA
'    Call CalcularTonicas_CA
'    Call CalcularSecundarias_CA
'    Call MarcarTonicaYSecundariaEnCadena_CA
'
'    ObjDTO.SilabasFinal = ObjDTO.SilabasAcentuadas
'
'    ' La fonología catalana irá en modFonologia_CA
'    Call ConstruirCadenaFonemas_CA
'
'    Entrada_Motor_CA = ObjDTO.SilabasAuto
'
'End Function
'
'' ============================================================
''   SILABEO DE FRASE
'' ============================================================
'Private Sub SilabearFrase_CA()
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
'    palabras = Split(frase, " ")
'
'    For i = LBound(palabras) To UBound(palabras)
'
'        limpia = Trim$(palabras(i))
'        If limpia <> "" Then
'
'            sil = SilabearPalabra_CA(limpia)
'
'            If resultado = "" Then
'                resultado = sil
'            Else
'                resultado = resultado & " |   | " & sil
'            End If
'
'        End If
'    Next i
'
'    ObjDTO.SilabasAuto = resultado
'
'End Sub
'
'' ============================================================
''   SILABEO DE PALABRA
'' ============================================================
'Public Function SilabearPalabra_CA(ByVal Texto As String) As String
'
'    Dim t As String
'    t = LCase$(Trim$(Texto))
'
'    If usarSilabeoMorfologico Then
'        SilabearPalabra_CA = SilabearMorfologico_CA(t)
'    Else
'        SilabearPalabra_CA = SilabearOrtog_CA(t)
'    End If
'
'End Function
'
'' ============================================================
''   SILABEO MORFOLÓGICO (opcional)
'' ============================================================
'Public Function SilabearMorfologico_CA(ByVal t As String) As String
'
'    Dim pref As String
'    Dim resto As String
'
'    If Not respetarPrefijos Then
'        SilabearMorfologico_CA = SilabearOrtog_CA(t)
'        Exit Function
'    End If
'
'    pref = DetectarPrefijo_CA(t)
'
'    If pref = "" Then
'        SilabearMorfologico_CA = SilabearOrtog_CA(t)
'        Exit Function
'    End If
'
'    resto = Mid$(t, Len(pref) + 1)
'
'    SilabearMorfologico_CA = pref & " | " & SilabearOrtog_CA(resto)
'
'End Function
'
'Private Function DetectarPrefijo_CA(ByVal t As String) As String
'    ' Puedes usar la misma tabla de prefijos, cambiando el filtro a [ca-es] = true
'    DetectarPrefijo_CA = ""
'End Function
'
'' ============================================================
''   SILABEO ORTOGRÁFICO CATALÁN
'' ============================================================
'Public Function SilabearOrtog_CA(ByVal t As String) As String
'
'    Dim nucIni() As Byte, nucFin() As Byte
'    Dim silIni() As Byte, silFin() As Byte
'    Dim nNuc As Byte, i As Byte
'    Dim silabas() As String
'
'    LocalizarNucleosOrtog_CA t, nucIni, nucFin, nNuc
'
'    ReDim silIni(1 To nNuc)
'    ReDim silFin(1 To nNuc)
'
'    CalcularSilabas_CA t, nucIni, nucFin, nNuc, silIni, silFin
'
'    ReDim silabas(1 To nNuc)
'    For i = 1 To nNuc
'        silabas(i) = Mid$(t, silIni(i), silFin(i) - silIni(i) + 1)
'    Next i
'
'    SilabearOrtog_CA = Join(silabas, " | ")
'
'End Function
'
'' ============================================================
''   LOCALIZAR NÚCLEOS (DIPTONGOS, TRIPTONGOS, HIATS)
'' ============================================================
'Private Sub LocalizarNucleosOrtog_CA(ByVal t As String, _
'                                     ByRef nucIni() As Byte, _
'                                     ByRef nucFin() As Byte, _
'                                     ByRef nNuc As Byte)
'
'    Dim i As Byte, L As Byte
'    Dim C1 As String, C2 As String, C3 As String
'
'    L = Len(t)
'    ReDim nucIni(1 To L)
'    ReDim nucFin(1 To L)
'    nNuc = 0
'
'    i = 1
'    Do While i <= L
'
'        C1 = Mid$(t, i, 1)
'
'        If EsVocal_CA(C1) Then
'
'            ' TRIFTONGOS
'            If i + 2 <= L Then
'                C2 = Mid$(t, i + 1, 1)
'                C3 = Mid$(t, i + 2, 1)
'                If EsTriftong_CA(C1, C2, C3) Then
'                    nNuc = nNuc + 1
'                    nucIni(nNuc) = i
'                    nucFin(nNuc) = i + 2
'                    i = i + 3
'                    GoTo Siguiente
'                End If
'            End If
'
'            ' DIPTONGOS
'            If i + 1 <= L Then
'                C2 = Mid$(t, i + 1, 1)
'                If EsDiptong_CA(C1, C2) Then
'                    nNuc = nNuc + 1
'                    nucIni(nNuc) = i
'                    nucFin(nNuc) = i + 1
'                    i = i + 2
'                    GoTo Siguiente
'                End If
'            End If
'
'            ' VOCAL SOLA
'            nNuc = nNuc + 1
'            nucIni(nNuc) = i
'            nucFin(nNuc) = i
'            i = i + 1
'
'        Else
'            i = i + 1
'        End If
'
'Siguiente:
'    Loop
'
'End Sub
'
'' ============================================================
''   CÁLCULO DE SÍLABAS (GRUPOS CONSONÁNTICOS CATALANES)
'' ============================================================
'Private Sub CalcularSilabas_CA(ByVal t As String, _
'                               ByRef nucIni() As Byte, _
'                               ByRef nucFin() As Byte, _
'                               ByVal nNuc As Byte, _
'                               ByRef silIni() As Byte, _
'                               ByRef silFin() As Byte)
'
'    Dim i As Byte, L As Byte
'    Dim a As Byte, b As Byte
'    Dim k As Byte
'    Dim C1 As String, C2 As String, grupo As String
'
'    L = Len(t)
'    silIni(1) = 1
'
'    For i = 1 To nNuc - 1
'
'        a = nucFin(i)
'        b = nucIni(i + 1)
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
'                C1 = Mid$(t, a + 1, 1)
'                C2 = Mid$(t, a + 2, 1)
'                grupo = C1 & C2
'
'                ' Dígrafos catalanes indivisibles
'                If grupo = "ny" Or grupo = "ll" Or grupo = "tx" Or grupo = "tg" Or grupo = "tj" Then
'                    silFin(i) = a
'                    silIni(i + 1) = a + 1
'                    GoTo Siguiente
'                End If
'
'                ' Grupos de ataque catalanes
'                If EsGrupoAtaque_CA(grupo) Then
'                    silFin(i) = a
'                    silIni(i + 1) = a + 1
'                Else
'                    silFin(i) = a + 1
'                    silIni(i + 1) = a + 2
'                End If
'
'            Case 3
'                silFin(i) = a + 1
'                silIni(i + 1) = a + 2
'
'            Case Else
'                silFin(i) = a + 2
'                silIni(i + 1) = a + 3
'
'        End Select
'
'Siguiente:
'    Next i
'
'    silFin(nNuc) = L
'
'End Sub
'
'' ============================================================
''   FUNCIONES AUXILIARES — VOCALES CATALANAS
'' ============================================================
'Private Function EsVocal_CA(ByVal c As String) As Boolean
'    EsVocal_CA = InStr("aeiouàèéíïòóúü", c) > 0
'End Function
'
'Private Function EsVocalDebil_CA(ByVal c As String) As Boolean
'    EsVocalDebil_CA = InStr("iuüï", c) > 0
'End Function
'
'Private Function EsVocalForta_CA(ByVal c As String) As Boolean
'    EsVocalForta_CA = InStr("aàeèéoò", c) > 0
'End Function
'
'Private Function EsDiptong_CA(ByVal v1 As String, ByVal v2 As String) As Boolean
'
'    If Not EsVocal_CA(v1) Or Not EsVocal_CA(v2) Then Exit Function
'
'    ' Dos fortes ? NO diptong
'    If EsVocalForta_CA(v1) And EsVocalForta_CA(v2) Then Exit Function
'
'    ' Dos febles ? SI diptong
'    If EsVocalDebil_CA(v1) And EsVocalDebil_CA(v2) Then
'        EsDiptong_CA = True
'        Exit Function
'    End If
'
'    ' Hiats obligatoris
'    If v2 = "ï" Or v2 = "ü" Then Exit Function
'
'    EsDiptong_CA = True
'
'End Function
'
'Private Function EsTriftong_CA(ByVal v1 As String, ByVal v2 As String, ByVal v3 As String) As Boolean
'    EsTriftong_CA = (EsVocalDebil_CA(v1) And EsVocalForta_CA(v2) And EsVocalDebil_CA(v3))
'End Function
'
'' ============================================================
''   GRUPOS DE ATAQUE CATALANES
'' ============================================================
'Private Function EsGrupoAtaque_CA(ByVal g As String) As Boolean
'
'    Dim AC As Variant
'    AC = Array("pr", "br", "tr", "dr", "cr", "gr", "fr", _
'               "pl", "bl", "cl", "gl", "fl")
'
'    EsGrupoAtaque_CA = (UBound(Filter(AC, g)) >= 0)
'
'End Function
'
'' ============================================================
''   ACENTUACIÓN CATALANA
'' ============================================================
'Private Sub CalcularTonicas_CA()
'
'    Dim tGlobal As New Collection
'    Dim elementos As Collection
'    Set elementos = ObtenerPalabrasDesdeSilabasAuto()
'
'    Dim globalIndex As Byte
'    Dim i As Byte
'
'    globalIndex = 0
'
'    For i = 1 To elementos.count
'
'        If TypeName(elementos(i)) = "Collection" Then
'
'            Dim w As Collection
'            Set w = elementos(i)
'
'            Dim tLocal As Long
'            tLocal = DetectarTonica_CA(w)
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
'    ObjDTO.SilabasTonicas = JoinCollection(tGlobal)
'
'End Sub
'
'Private Function DetectarTonica_CA(w As Collection) As Byte
'
'    Dim i As Byte
'
'    ' 1) Si hay tilde ? esa es la tónica
'    For i = 1 To w.count
'        If TieneTilde_CA(w(i)) Then
'            DetectarTonica_CA = i
'            Exit Function
'        End If
'    Next i
'
'    ' 2) Si no hay tilde ? reglas catalanas
'    '    · Aguda si termina en: vocal, as, es, is, os, us, en, in
'    '    · Sino ? plana
'
'    Dim ultima As String
'    ultima = w(w.count)
'
'    If TerminaAguda_CA(ultima) Then
'        DetectarTonica_CA = w.count
'    Else
'        DetectarTonica_CA = w.count - 1
'    End If
'
'End Function
'
'Private Function TieneTilde_CA(ByVal sil As String) As Boolean
'    TieneTilde_CA = (sil Like "*[àèéíïòóúü]*")
'End Function
'
'Private Function TerminaAguda_CA(ByVal s As String) As Boolean
'
'    Dim t As String
'    t = LCase$(s)
'
'    If t Like "*[aeiouàèéíïòóúü]" Then TerminaAguda_CA = True: Exit Function
'    If t Like "*as" Or t Like "*es" Or t Like "*is" Or t Like "*os" Or t Like "*us" Then TerminaAguda_CA = True: Exit Function
'    If t Like "*en" Or t Like "*in" Then TerminaAguda_CA = True: Exit Function
'
'    TerminaAguda_CA = False
'
'End Function
'
'' ============================================================
''   SECUNDARIAS (idéntico al español)
'' ============================================================
'Private Sub CalcularSecundarias_CA()
'    Call CalcularSecundarias   ' reutilizamos el módulo ES
'End Sub
'
'Private Sub MarcarTonicaYSecundariaEnCadena_CA()
'    Call MarcarTonicaYSecundariaEnCadena   ' reutilizamos el módulo ES
'End Sub
'
'
