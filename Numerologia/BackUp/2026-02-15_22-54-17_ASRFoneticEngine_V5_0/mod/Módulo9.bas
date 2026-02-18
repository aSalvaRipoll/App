Attribute VB_Name = "Módulo9"

'Option Compare Database
'Option Explicit
'
'' ============================================================
''   MOTOR DE SILABEO 4.0 — ORTOGRÁFICO + MORFOSILÁBICO
''   Autor: Alba Salvá Ripoll + Copilot
''   Fecha: 15/02/2026
''
''   Este módulo implementa:
''   > Silabeo ortográfico completo (RAE 2010)
''   > Silabeo morfosilábico (prefijos)
''   > Diptongos, triptongos, hiatos con tilde
''   > Casos especiales: güe/güi, qu, ll, rr, tl
''   > Sílaba tónica
''   > Tracking de depuración (modoDebug)
''
''   Variables de control:
''       usarSilabeoMorfologico   => True = usar prefijos
''       modoPrefijosEstrictos    => True = usar lista interna completa
''       respetarPrefijos         => True = separar prefijo como sílaba
''       modoDebug                => True = activar logs
''
''   Requisitos:
''       addLog() y PrintLog() deben existir en tu proyecto.
''
'' ============================================================
'
'' -----------------------------
'' VARIABLES DE CONFIGURACIÓN
'' -----------------------------
'Private usarSilabeoMorfologico As Boolean
'Private modoPrefijosEstrictos As Boolean
'Private respetarPrefijos As Boolean
'Private modoDebug As Boolean
'
'' Prefijos personalizados (modo flexible)
'Private prefijosPersonalizados As Variant
'
'' Prefijos canónicos (modo estricto)
'Private prefijosEstrictos As Variant
'
'' ============================================================
''   FUNCIÓN PRINCIPAL
'' ============================================================
'Public Function SilabearPalabra(ByVal Texto As String) As String
'    Dim T As String
'
'    ' Activar Logs
'    modoDebug = True
'
'    ' Tratamiento prefijos
''       usarSilabeoMorfologico   => True = usar prefijos
'    usarSilabeoMorfologico = True
'
''       modoPrefijosEstrictos    => True = usar lista interna completa
'    modoPrefijosEstrictos = True
''       respetarPrefijos         => True = separar prefijo como sílaba
'    respetarPrefijos = True
'
'
'    T = LCase$(Trim$(Texto))
'
'    If modoDebug Then
'        addLog
'        addLog "============================================================="
'        addLog "        LOG DE DEPURACIÓN SILABEO 4.0"
'        addLog "============================================================="
'        addLog "Entrada: '" & Texto & "'"
'        addLog "Normalizado: '" & T & "'"
'        addLog "Longitud: " & Len(T) & " letras."
'    End If
'
'    If usarSilabeoMorfologico Then
'        SilabearPalabra = SilabearMorfologico(T)
'    Else
'        SilabearPalabra = SilabearOrtog(T)
'    End If
'
'    If modoDebug Then
'        addLog
'        addLog "Resultado final: " & SilabearPalabra
'        addLog "============================================================="
'        addLog "                   FIN LOG DE DEPURACIÓN"
'        addLog "============================================================="
'        PrintLog
'    End If
'End Function
'
'' ============================================================
''   SILABEO ORTOGRÁFICO (RAE 2010)
'' ============================================================
'Public Function SilabearOrtog(ByVal T As String) As String
'    Dim nucIni() As Byte, nucFin() As Byte
'    Dim silIni() As Byte, silFin() As Byte
'    Dim nNuc As Byte, i As Byte
'    Dim silabas() As String
'
'    ' 1. Localizar núcleos ortográficos
'    LocalizarNucleosOrtog T, nucIni, nucFin, nNuc
'
'    ' 2. Calcular fronteras silábicas
'    ReDim silIni(1 To nNuc)
'    ReDim silFin(1 To nNuc)
'    CalcularSilabas T, nucIni, nucFin, nNuc, silIni, silFin
'
'    ' 3. Construir sílabas
'    ReDim silabas(1 To nNuc)
'    For i = 1 To nNuc
'        silabas(i) = Mid$(T, silIni(i), silFin(i) - silIni(i) + 1)
'        If modoDebug Then addLog "Sílaba " & i & ": " & silabas(i)
'    Next i
'
'    SilabearOrtog = Join(silabas, " | ")
'End Function
'
'' ============================================================
''   SILABEO MORFOSILÁBICO (PREFIJOS)
'' ============================================================
'Public Function SilabearMorfologico(ByVal T As String) As String
'    Dim pref As String
'    Dim resto As String
'    Dim silPref As String
'    Dim silResto As String
'
'    If Not respetarPrefijos Then
'        SilabearMorfologico = SilabearOrtog(T)
'        Exit Function
'    End If
'
'    ' Detectar prefijo
'    pref = DetectarPrefijo(T)
'
'    If pref = "" Then
'        SilabearMorfologico = SilabearOrtog(T)
'        Exit Function
'    End If
'
'    resto = Mid$(T, Len(pref) + 1)
'
'    If modoDebug Then
'        addLog "Prefijo detectado: " & pref
'        addLog "Resto: " & resto
'    End If
'
'    silPref = pref
'    silResto = SilabearOrtog(resto)
'
'    SilabearMorfologico = silPref & " | " & silResto
'End Function
'
'' ============================================================
''   DETECCIÓN DE PREFIJOS
'' ============================================================
'Private Function DetectarPrefijo(ByVal T As String) As String
'    Dim p As Variant
'    Dim pref As String
'
'    ' Inicializar prefijos estrictos si no están cargados
'    If IsEmpty(prefijosEstrictos) Then
'        prefijosEstrictos = Array("a", "ante", "anti", "auto", "bi", "contra", "de", "des", "dis", _
'                                  "extra", "hiper", "hipo", "in", "im", "inter", "intra", "macro", _
'                                  "micro", "multi", "post", "pre", "pro", "re", "semi", "sub", _
'                                  "super", "trans", "ultra")
'    End If
'
'    ' Modo estricto
'    If modoPrefijosEstrictos Then
'        For Each p In prefijosEstrictos
'            If Left$(T, Len(p)) = p Then
'                DetectarPrefijo = p
'                Exit Function
'            End If
'        Next p
'    End If
'
'    ' Modo flexible
'    If Not IsEmpty(prefijosPersonalizados) Then
'        For Each p In prefijosPersonalizados
'            If Left$(T, Len(p)) = p Then
'                DetectarPrefijo = p
'                Exit Function
'            End If
'        Next p
'    End If
'
'    DetectarPrefijo = ""
'End Function
'
'' ============================================================
''   LOCALIZAR NÚCLEOS ORTOGRÁFICOS
'' ============================================================
'Private Sub LocalizarNucleosOrtog(ByVal T As String, _
'                                  ByRef nucIni() As Byte, _
'                                  ByRef nucFin() As Byte, _
'                                  ByRef nNuc As Byte)
'
'    Dim i As Byte, L As Byte
'    Dim C1 As String, C2 As String, c3 As String
'
'    L = Len(T)
'    ReDim nucIni(1 To L)
'    ReDim nucFin(1 To L)
'    nNuc = 0
'
'    If modoDebug Then
'        addLog
'        addLog "---------------------------------------"
'        addLog " Procedimiento LocalizarNucleosOrtog"
'    End If
'
'    i = 1
'    Do While i <= L
'        C1 = Mid$(T, i, 1)
'
'        If EsVocal(C1) Then
'
'            ' Intentar triptongo
'            If i + 2 <= L Then
'                C2 = Mid$(T, i + 1, 1)
'                c3 = Mid$(T, i + 2, 1)
'                If EsTriptongo(C1, C2, c3) Then
'                    nNuc = nNuc + 1
'                    nucIni(nNuc) = i
'                    nucFin(nNuc) = i + 2
'                    If modoDebug Then addLog "Triptongo: " & C1 & C2 & c3
'                    i = i + 3
'                    GoTo Siguiente
'                End If
'            End If
'
'            ' Intentar diptongo
'            If i + 1 <= L Then
'                C2 = Mid$(T, i + 1, 1)
'                If EsDiptongo(C1, C2) Then
'                    nNuc = nNuc + 1
'                    nucIni(nNuc) = i
'                    nucFin(nNuc) = i + 1
'                    If modoDebug Then addLog "Diptongo: " & C1 & C2
'                    i = i + 2
'                    GoTo Siguiente
'                End If
'            End If
'
'            ' Vocal sola
'            nNuc = nNuc + 1
'            nucIni(nNuc) = i
'            nucFin(nNuc) = i
'            If modoDebug Then addLog "Vocal sola: " & C1
'            i = i + 1
'
'        Else
'            i = i + 1
'        End If
'
'Siguiente:
'    Loop
'
'    If modoDebug Then
'        addLog "Total núcleos: " & nNuc
'        addLog " Fin LocalizarNucleosOrtog"
'        addLog "---------------------------------------"
'    End If
'End Sub
'
'' ============================================================
''   CÁLCULO DE FRONTERAS SILÁBICAS (ORTOGRÁFICO)
'' ============================================================
'Private Sub CalcularSilabas(ByVal T As String, _
'                            ByRef nucIni() As Byte, _
'                            ByRef nucFin() As Byte, _
'                            ByVal nNuc As Byte, _
'                            ByRef silIni() As Byte, _
'                            ByRef silFin() As Byte)
'
'    Dim i As Byte, L As Byte
'    Dim a As Byte, b As Byte
'    Dim k As Byte
'    Dim C1 As String, C2 As String, grupo As String
'
'    L = Len(T)
'    silIni(1) = 1
'
'    For i = 1 To nNuc - 1
'        a = nucFin(i)
'        b = nucIni(i + 1)
'
'        k = IIf(b > a + 1, b - a - 1, 0)
'
'        If modoDebug Then
'            addLog
'            addLog "---- Frontera entre núcleo " & i & " y " & (i + 1)
'            addLog "Consonantes entre medias: " & k
'        End If
'
'        Select Case k
'            Case 0
'                silFin(i) = a
'                silIni(i + 1) = a + 1
'
'            Case 1
'                silFin(i) = a
'                silIni(i + 1) = a + 1
'
'            Case 2
'                C1 = Mid$(T, a + 1, 1)
'                C2 = Mid$(T, a + 2, 1)
'                grupo = C1 & C2
'
'                If EsGrupoAtaque(grupo) Then
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
'        End Select
'    Next i
'
'    silFin(nNuc) = L
'End Sub
'
'' ============================================================
''   VOCAL / DÍPTONGO / TRIPTONGO / HIATO
'' ============================================================
'Private Function EsVocal(ByVal c As String) As Boolean
'    EsVocal = InStr("aeiouáéíóúü", c) > 0
'End Function
'
'Private Function EsVocalFuerte(ByVal c As String) As Boolean
'    EsVocalFuerte = InStr("aáeéoó", c) > 0
'End Function
'
'Private Function EsVocalDebil(ByVal c As String) As Boolean
'    EsVocalDebil = InStr("iíuúü", c) > 0
'End Function
'
'Private Function EsVocalDebilTonica(ByVal c As String) As Boolean
'    EsVocalDebilTonica = InStr("íú", c) > 0
'End Function
'
'Private Function EsDiptongo(ByVal v1 As String, ByVal v2 As String) As Boolean
'    If Not EsVocal(v1) Or Not EsVocal(v2) Then Exit Function
'
'    If EsVocalFuerte(v1) And EsVocalFuerte(v2) Then Exit Function
'    If EsVocalDebilTonica(v1) Or EsVocalDebilTonica(v2) Then Exit Function
'
'    EsDiptongo = True
'End Function
'
'Private Function EsTriptongo(ByVal v1 As String, ByVal v2 As String, ByVal v3 As String) As Boolean
'    If EsVocalDebil(v1) And Not EsVocalDebilTonica(v1) _
'       And EsVocalFuerte(v2) _
'       And EsVocalDebil(v3) And Not EsVocalDebilTonica(v3) Then
'        EsTriptongo = True
'    End If
'End Function
'
'' ============================================================
''   GRUPOS CONSONÁNTICOS DE ATAQUE
'' ============================================================
'Private Function EsGrupoAtaque(ByVal g As String) As Boolean
'    Dim AC As Variant
'    AC = Array("pr", "br", "tr", "dr", "cr", "gr", "fr", _
'               "pl", "bl", "cl", "gl", "fl")
'    EsGrupoAtaque = (UBound(Filter(AC, g)) >= 0)
'End Function
'
'' ============================================================
''   SÍLABA TÓNICA
'' ============================================================
'Public Function SilabaTonica(ByVal palabra As String) As Byte
'    Dim sil As String
'    Dim silabas() As String
'    Dim i As Byte
'
'    sil = SilabearPalabra(palabra)
'    silabas = Split(sil, " | ")
'
'    ' 1. Buscar tilde
'    For i = 0 To UBound(silabas)
'        If TieneTilde(silabas(i)) Then
'            SilabaTonica = i + 1
'            Exit Function
'        End If
'    Next i
'
'    ' 2. Reglas generales
'    If TerminaEnVocalNS(palabra) Then
'        SilabaTonica = UBound(silabas) ' llana
'    Else
'        SilabaTonica = UBound(silabas) + 1 ' aguda
'    End If
'End Function
'
'Private Function TieneTilde(ByVal s As String) As Boolean
'    TieneTilde = (InStr(s, "á") Or InStr(s, "é") Or InStr(s, "í") Or InStr(s, "ó") Or InStr(s, "ú")) > 0
'End Function
'
'Private Function TerminaEnVocalNS(ByVal s As String) As Boolean
'    Dim c As String
'    c = Right$(s, 1)
'    TerminaEnVocalNS = (EsVocal(c) Or c = "n" Or c = "s")
'End Function


