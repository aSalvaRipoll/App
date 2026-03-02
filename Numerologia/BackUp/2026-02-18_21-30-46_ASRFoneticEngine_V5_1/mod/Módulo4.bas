Attribute VB_Name = "Módulo4"
Option Compare Database
Option Explicit

'' ============================================================
''   SILABEO DE FRASE — IB
'' ============================================================
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
'    If DebugMotor Then
'        addLog
'        addLog "SilabearFrase_IB ? " & ObjDTO.SilabasAuto
'    End If
'
'End Sub

'' ============================================================
''   SILABEO DE PALABRA — IB
'' ============================================================
'Private Function SilabearPalabra_IB(ByVal texto As String) As String
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

'' ============================================================
''   SILABEO ORTOGRÁFICO — IB
''   (estructura idéntica a CA, reglas IB dentro)
'' ============================================================
'Private Function SilabearOrtog_IB(ByVal t As String) As String
'    Dim nucIni() As Byte, nucFin() As Byte
'    Dim silIni() As Byte, silFin() As Byte
'    Dim nNuc As Byte, i As Byte
'    Dim silabas() As String
'
'    If Len(Trim$(t)) < 2 Then
'        SilabearOrtog_IB = t
'        Exit Function
'    End If
'
'    LocalizarNucleosOrtog_IB t, nucIni, nucFin, nNuc
'
'    ReDim silIni(1 To nNuc)
'    ReDim silFin(1 To nNuc)
'
'    CalcularSilabas_IB t, nucIni, nucFin, nNuc, silIni, silFin
'
'    ReDim silabas(1 To nNuc)
'    For i = 1 To nNuc
'        silabas(i) = Mid$(t, silIni(i), silFin(i) - silIni(i) + 1)
'        If DebugMotor Then
'            addLog "IB — Sílaba " & i & ": " & silabas(i)
'        End If
'    Next i
'
'    SilabearOrtog_IB = Join(silabas, " | ")
'
'End Function

'Private Function SilabearMorfologico_IB(ByVal t As String) As String
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

'' ============================================================
''   CÁLCULO DE SÍLABAS — IB
'' ============================================================
'Private Sub CalcularSilabas_IB(ByVal t As String, _
'                            ByRef nucIni() As Byte, _
'                            ByRef nucFin() As Byte, _
'                            ByVal nNuc As Byte, _
'                            ByRef silIni() As Byte, _
'                            ByRef silFin() As Byte)
'
'    Dim i As Byte, L As Byte
'    Dim a As Byte, b As Byte
'    Dim k As Byte
'    Dim c1 As String, c2 As String, c3 As String, grupo As String
'
'    L = Len(t)
'    silIni(1) = 1
'
'    For i = 1 To nNuc - 1
'        a = nucFin(i)
'        b = nucIni(i + 1)
'
'        k = IIf(b > a + 1, b - a - 1, 0)
'
'        If DebugMotor Then
'            addLog
'            addLog "---- Frontera IB entre núcleo " & i & " y " & (i + 1)
'            addLog "Consonantes entre medias: " & k
'        End If
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
'                grupo = c1 & c2
'
'                ' Dígrafos indivisibles IB
'                If grupo = "rr" Or grupo = "ll" Or grupo = "ch" Then
'                    silFin(i) = a
'                    silIni(i + 1) = a + 1
'                    GoTo Siguiente
'                End If
'
'                ' Grupos de ataque IB
'                If EsGrupoAtaque_IB(grupo) Then
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
'                c3 = Mid$(t, a + 3, 1)
'
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
'Siguiente:
'    Next i
'
'    silFin(nNuc) = L
'
'End Sub

'' ============================================================
''   PREFIJOS IB (estructura CA, filtro IB)
'' ============================================================
'Private Function DetectarPrefijo_IB(ByVal t As String) As String
'    Dim p As Variant
'
'    If Not prefijosCargados_IB Then CargarPrefijos_IB
'
'    For Each p In prefijosEstrictos_IB
'        If Len(t) = Len(p) Then Exit For
'        If Left$(t, Len(p)) = p Then
'            DetectarPrefijo_IB = p
'            Exit Function
'        End If
'    Next p
'
'    DetectarPrefijo_IB = ""
'End Function

'Private Sub CargarPrefijos_IB()
'    Dim rs As DAO.Recordset
'    Dim sql As String
'    Dim i As Long
'
'    If prefijosCargados_IB Then Exit Sub
'
'    sql = "SELECT Prefijo FROM qryPrefijos " & _
'          "WHERE Activo = 1 " & _
'            "AND Tipo Like 'auténtico' " & _
'            "AND [ca-ib] = true " & _
'          "ORDER BY Len(Prefijo) DESC, Prefijo ASC"
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


'' ============================================================
''   LOCALIZAR NÚCLEOS ORTOGRÁFICOS — IB
'' ============================================================
'Private Sub LocalizarNucleosOrtog_IB(ByVal t As String, _
'                                  ByRef nucIni() As Byte, _
'                                  ByRef nucFin() As Byte, _
'                                  ByRef nNuc As Byte)
'
'    Dim i As Byte, L As Byte
'    Dim c1 As String, c2 As String, c3 As String
'
'    L = Len(t)
'    ReDim nucIni(1 To L)
'    ReDim nucFin(1 To L)
'    nNuc = 0
'
'    If DebugMotor Then
'        addLog
'        addLog "---------------------------------------"
'        addLog " Procedimiento LocalizarNucleosOrtog_IB"
'    End If
'
'    i = 1
'    Do While i <= L
'        c1 = Mid$(t, i, 1)
'
'        If EsVocal_IB(c1) Then
'
'            ' Triptongo IB
'            If i + 2 <= L Then
'                c2 = Mid$(t, i + 1, 1)
'                c3 = Mid$(t, i + 2, 1)
'                If EsTriptong_IB(c1, c2, c3) Then
'                    nNuc = nNuc + 1
'                    nucIni(nNuc) = i
'                    nucFin(nNuc) = i + 2
'                    If DebugMotor Then
'                        addLog "Triptongo IB: " & c1 & c2 & c3
'                    End If
'                    i = i + 3
'                    GoTo Siguiente
'                End If
'            End If
'
'            ' Diptongo IB
'            If i + 1 <= L Then
'                c2 = Mid$(t, i + 1, 1)
'                If EsDiptong_IB(c1, c2) Then
'                    nNuc = nNuc + 1
'                    nucIni(nNuc) = i
'                    nucFin(nNuc) = i + 1
'                    If DebugMotor Then
'                        addLog "Diptongo IB: " & c1 & c2
'                    End If
'                    i = i + 2
'                    GoTo Siguiente
'                End If
'            End If
'
'            ' Vocal sola
'            nNuc = nNuc + 1
'            nucIni(nNuc) = i
'            nucFin(nNuc) = i
'            If DebugMotor Then
'                addLog "Vocal sola IB: " & c1
'            End If
'            i = i + 1
'
'        Else
'            i = i + 1
'        End If
'
'Siguiente:
'    Loop
'
'    If DebugMotor Then
'        addLog "Total núcleos IB: " & nNuc
'        addLog " Fin LocalizarNucleosOrtog_IB"
'        addLog "---------------------------------------"
'    End If
'End Sub

' ============================================================
'   FUNCIONES DE VOCAL — IB
' ============================================================
'Private Function EsVocal_IB(ByVal c As String) As Boolean
'    EsVocal_IB = InStr("aeiouàèéíïòóúü", LCase$(c)) > 0
'End Function


'Private Function EsDiptong_IB(ByVal v1 As String, ByVal v2 As String) As Boolean
'
'    If Not EsVocal_IB(v1) Or Not EsVocal_IB(v2) Then Exit Function
'
'    ' Dos fortes ? hiat
'    If EsVocalForta_IB(v1) And EsVocalForta_IB(v2) Then Exit Function
'
'    ' Dièresi a la segona ? hiat
'    If v2 = "ï" Or v2 = "ü" Then
'        EsDiptong_IB = False
'        Exit Function
'    End If
'
'    ' Dos febles ? diptong
'    If EsVocalFeble_IB(v1) And EsVocalFeble_IB(v2) Then
'        EsDiptong_IB = True
'        Exit Function
'    End If
'
'    ' Feble tònica + forta ? hiat
'    If (EsVocalFebleTonica_IB(v1) And EsVocalForta_IB(v2)) _
'    Or (EsVocalForta_IB(v1) And EsVocalFebleTonica_IB(v2)) Then
'        Exit Function
'    End If
'
'    ' Resta ? diptong
'    EsDiptong_IB = True
'
'End Function

'Private Function EsGrupoAtaque_IB(ByVal g As String) As Boolean
'    Dim AC As Variant
'    AC = Array("pr", "br", "tr", "dr", "cr", "gr", "fr", _
'               "pl", "bl", "cl", "gl", "fl")
'    EsGrupoAtaque_IB = (UBound(Filter(AC, g)) >= 0)
'End Function

'Private Function EsTriptong_IB(ByVal v1 As String, ByVal v2 As String, ByVal v3 As String) As Boolean
'    If EsVocalFeble_IB(v1) And Not EsVocalFebleTonica_IB(v1) _
'       And EsVocalForta_IB(v2) _
'       And EsVocalFeble_IB(v3) And Not EsVocalFebleTonica_IB(v3) Then
'        EsTriptong_IB = True
'    End If
'End Function

'Private Function EsVocalFebleTonica_IB(ByVal c As String) As Boolean
'    EsVocalFebleTonica_IB = InStr("íú", LCase$(c)) > 0
'End Function

'Private Function EsVocalFeble_IB(ByVal c As String) As Boolean
'    EsVocalFeble_IB = InStr("iíïuúü", LCase$(c)) > 0
'End Function

'Private Function EsVocalForta_IB(ByVal c As String) As Boolean
'    EsVocalForta_IB = InStr("aàeèéoò", LCase$(c)) > 0
'End Function

'Private Function PuedeCerrarSilaba_IB(ByVal c As String) As Boolean
'    PuedeCerrarSilaba_IB = Not (c = "r" Or c = "l" Or c = "h")
'End Function

'Private Function TieneTilde_IB(ByVal silaba As String) As Boolean
'    TieneTilde_IB = (InStr(silaba, "à") > 0 Or _
'                     InStr(silaba, "á") > 0 Or _
'                     InStr(silaba, "è") > 0 Or _
'                     InStr(silaba, "é") > 0 Or _
'                     InStr(silaba, "í") > 0 Or _
'                     InStr(silaba, "ï") > 0 Or _
'                     InStr(silaba, "ò") > 0 Or _
'                     InStr(silaba, "ó") > 0 Or _
'                     InStr(silaba, "ú") > 0 Or _
'                     InStr(silaba, "ü") > 0)
'End Function

'===========================================================================================

'' ============================================================
''   DETECTAR SÍLABES TÒNIQUES — IB
''   (estructura idèntica al CA)
'' ============================================================
'Private Sub CalcularTonicas_IB()
'
'    Dim tGlobal As New Collection
'    Dim elementos As Collection
'    Set elementos = ObtenerPalabrasDesdeSilabasAuto_IB()
'
'    Dim globalIndex As Byte
'    Dim i As Byte
'
'    globalIndex = 0
'
'    For i = 1 To elementos.count
'
'        If TypeName(elementos(i)) = "Collection" Then
'            Dim w As Collection
'            Set w = elementos(i)
'
'            Dim tLocal As Byte
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
'
'    If DebugMotor Then
'        addLog "CalcularTonicas_IB ? " & ObjDTO.SilabasTonicas
'    End If
'
'End Sub

'' ============================================================
''   DETECTAR SÍLABES SECUNDÀRIES — IB
'' ============================================================
'Private Sub CalcularSecundarias_IB()
'
'    Dim sGlobal As New Collection
'    Dim elementos As Collection
'    Set elementos = ObtenerPalabrasDesdeSilabasAuto_IB()
'
'    Dim globalIndex As Byte
'    Dim i As Byte
'    Dim tLocal As Byte
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
'    If DebugMotor Then
'        addLog "CalcularSecundarias_IB ? " & ObjDTO.SilabasSecundarias
'    End If
'
'End Sub

'' ============================================================
''   DETECTAR TÒNICA LOCAL — IB
'' ============================================================
'Private Function DetectarTonica_IB(w As Collection) As Byte
'
'    Dim i As Byte
'    Dim palabra As String
'    Dim ultima As String
'    Dim terminaLlana As Boolean
'
'    ' 1) Si alguna síl·laba té accent gràfic ? tònica directa
'    For i = 1 To w.count
'        If TieneTilde_IB(w(i)) Then
'            DetectarTonica_IB = i
'            Exit Function
'        End If
'    Next i
'
'    ' 2) Reconstruir la paraula
'    palabra = ""
'    For i = 1 To w.count
'        palabra = palabra & w(i)
'    Next i
'
'    ultima = Right$(palabra, 1)
'    terminaLlana = False
'
'    ' Regles mallorquines (igual que CA, però amb vocals IB)
'    If InStr("aeiouàèéíïòóúü", ultima) > 0 Then terminaLlana = True
'    If LCase$(Right$(palabra, 2)) Like "*as" Then terminaLlana = True
'    If LCase$(Right$(palabra, 2)) Like "*es" Then terminaLlana = True
'    If LCase$(Right$(palabra, 2)) Like "*is" Then terminaLlana = True
'    If LCase$(Right$(palabra, 2)) Like "*os" Then terminaLlana = True
'    If LCase$(Right$(palabra, 2)) Like "*us" Then terminaLlana = True
'    If LCase$(Right$(palabra, 2)) Like "*en" Then terminaLlana = True
'    If LCase$(Right$(palabra, 2)) Like "*in" Then terminaLlana = True
'
'    If terminaLlana And w.count >= 2 Then
'        DetectarTonica_IB = w.count - 1
'    Else
'        DetectarTonica_IB = w.count
'    End If
'
'End Function

'Private Function DetectarSecundarias_IB(w As Collection, tPos As Byte) As Collection
'
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
'    secs.Add 1
'
'    If n >= 6 Then
'        pos2 = tPos - 2
'        If pos2 > 1 Then secs.Add pos2
'    End If
'
'    Set DetectarSecundarias_IB = secs
'
'End Function


'' ============================================================
''   MARCAR TÒNICA I SECUNDÀRIES — IB
'' ============================================================
'Private Sub MarcarTonicaYSecundariaEnCadena_IB()
'
'    Dim sils As Variant
'    Dim i As Byte
'    Dim out() As String
'
'    sils = Split(ObjDTO.SilabasAuto, " | ")
'    ReDim out(LBound(sils) To UBound(sils))
'
'    For i = LBound(sils) To UBound(sils)
'        out(i) = sils(i)
'    Next i
'
'    ' TÒNICA
'    If ObjDTO.SilabasTonicas <> "" Then
'        Dim t As Variant, x As Variant
'        t = Split(ObjDTO.SilabasTonicas, ",")
'        For Each x In t
'            Dim idx As Byte
'            idx = CByte(x) - 1
'            If idx >= LBound(out) And idx <= UBound(out) Then
'                out(idx) = "( " & out(idx) & " )"
'            End If
'        Next x
'    End If
'
'    ' SECUNDÀRIES
'    If ObjDTO.SilabasSecundarias <> "" Then
'        Dim s As Variant, y As Variant
'        s = Split(ObjDTO.SilabasSecundarias, ",")
'        For Each y In s
'            Dim idx2 As Byte
'            idx2 = CByte(y) - 1
'            If idx2 >= LBound(out) And idx2 <= UBound(out) Then
'                out(idx2) = "[ " & out(idx2) & " ]"
'            End If
'        Next y
'    End If
'
'    ObjDTO.SilabasAcentuadas = Join(out, " | ")
'
'    If DebugMotor Then
'        addLog "MarcarTonicaYSecundariaEnCadena_IB ? " & ObjDTO.SilabasAcentuadas
'    End If
'
'End Sub

'' ============================================================
''   OBTENIR PARAULES DES DE SILABAS AUTO — IB
'' ============================================================
'Private Function ObtenerPalabrasDesdeSilabasAuto_IB() As Collection
'
'    Dim resultado As New Collection
'    Dim palabraActual As New Collection
'
'    Dim sils As Variant
'    sils = Split(ObjDTO.SilabasAuto, " | ")
'
'    Dim i As Byte
'    For i = LBound(sils) To UBound(sils)
'
'        If Trim$(sils(i)) = "" Then
'            If palabraActual.count > 0 Then
'                resultado.Add palabraActual
'                Set palabraActual = New Collection
'            End If
'            resultado.Add "HUECO"
'        Else
'            palabraActual.Add sils(i)
'        End If
'
'    Next i
'
'    If palabraActual.count > 0 Then resultado.Add palabraActual
'
'    Set ObtenerPalabrasDesdeSilabasAuto_IB = resultado
'
'End Function

