Attribute VB_Name = "Módulo1"
'Option Compare Database
'Option Explicit
'
'' ============================================================
''   ENTRADA PRINCIPAL DEL MOTOR (VALENCIANO)
'' ============================================================
'Public Function Entrada_Motor_VA(texto As String) As String
'
'    Set ObjDTO = New clsDTO_Motor
'
'    ObjDTO.TextoOriginal = texto
'    ObjDTO.NormalizaEntrada
'
'    ' 1) Silabeo automático
'    Call SilabearFrase_VA
'
'    ' 2) Detectar tónicas
'    Call CalcularTonicas_VA
'
'    ' 3) Detectar secundarias
'    Call CalcularSecundarias_VA
'
'    ' 4) Marcar tónicas y secundarias
'    Call MarcarTonicaYSecundariaEnCadena_VA
'
'    ObjDTO.SilabasFinal = ObjDTO.SilabasAcentuadas
'
'    ' 5) Generar fonética (se implementará después)
'    ' Call ConstruirCadenaFonemas_VA
'
'    Entrada_Motor_VA = ObjDTO.SilabasAuto
'
'End Function
'
'Private Sub SilabearFrase_VA()
'
'    Dim frase As String
'    Dim palabras() As String
'    Dim resultado As String
'    Dim i As Long
'    Dim limpia As String
'    Dim sil As String
'
'    frase = ObjDTO.TextoNormalizado
'    palabras = Split(frase, " ")
'
'    For i = LBound(palabras) To UBound(palabras)
'        limpia = Trim$(palabras(i))
'        If limpia <> "" Then
'            sil = SilabearPalabra_VA(limpia)
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
'Private Function SilabearPalabra_VA(ByVal texto As String) As String
'    Dim t As String
'
'    t = LCase$(Trim$(texto))
'
'    SilabearPalabra_VA = SilabearOrtog_VA(t)
'
'End Function
'
'Private Function SilabearOrtog_VA(ByVal t As String) As String
'    Dim nucIni() As Byte, nucFin() As Byte
'    Dim silIni() As Byte, silFin() As Byte
'    Dim nNuc As Byte, i As Byte
'    Dim silabas() As String
'
'    If Len(t) < 2 Then
'        SilabearOrtog_VA = t
'        Exit Function
'    End If
'
'    LocalizarNucleosOrtog_VA t, nucIni, nucFin, nNuc
'
'    ReDim silIni(1 To nNuc)
'    ReDim silFin(1 To nNuc)
'
'    CalcularSilabas_VA t, nucIni, nucFin, nNuc, silIni, silFin
'
'    ReDim silabas(1 To nNuc)
'    For i = 1 To nNuc
'        silabas(i) = Mid$(t, silIni(i), silFin(i) - silIni(i) + 1)
'    Next i
'
'    SilabearOrtog_VA = Join(silabas, " | ")
'
'End Function
'
'Private Sub LocalizarNucleosOrtog_VA(ByVal t As String, _
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
'    i = 1
'    Do While i <= L
'
'        c1 = Mid$(t, i, 1)
'
'        If EsVocal_VA(c1) Then
'
'            ' Triptongo
'            If i + 2 <= L Then
'                c2 = Mid$(t, i + 1, 1)
'                c3 = Mid$(t, i + 2, 1)
'                If EsTriptongo_VA(c1, c2, c3) Then
'                    nNuc = nNuc + 1
'                    nucIni(nNuc) = i
'                    nucFin(nNuc) = i + 2
'                    i = i + 3
'                    GoTo Siguiente
'                End If
'            End If
'
'            ' Diptongo
'            If i + 1 <= L Then
'                c2 = Mid$(t, i + 1, 1)
'                If EsDiptongo_VA(c1, c2) Then
'                    nNuc = nNuc + 1
'                    nucIni(nNuc) = i
'                    nucFin(nNuc) = i + 1
'                    i = i + 2
'                    GoTo Siguiente
'                End If
'            End If
'
'            ' Vocal sola
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
'Private Sub CalcularSilabas_VA(ByVal t As String, _
'                            ByRef nucIni() As Byte, _
'                            ByRef nucFin() As Byte, _
'                            ByVal nNuc As Byte, _
'                            ByRef silIni() As Byte, _
'                            ByRef silFin() As Byte)
'
'    Dim i As Byte, L As Byte
'    Dim a As Byte, b As Byte
'    Dim k As Byte
'    Dim c1 As String, c2 As String, grupo As String
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
'            Case 0, 1
'                silFin(i) = a
'                silIni(i + 1) = a + 1
'
'            Case 2
'                c1 = Mid$(t, a + 1, 1)
'                c2 = Mid$(t, a + 2, 1)
'                grupo = c1 & c2
'
'                If grupo = "rr" Or grupo = "l·l" Or grupo = "ny" Or grupo = "tx" Then
'                    silFin(i) = a
'                    silIni(i + 1) = a + 1
'                ElseIf EsGrupoAtaque_VA(grupo) Then
'                    silFin(i) = a
'                    silIni(i + 1) = a + 1
'                Else
'                    silFin(i) = a + 1
'                    silIni(i + 1) = a + 2
'                End If
'
'            Case Else
'                silFin(i) = a + 1
'                silIni(i + 1) = a + 2
'
'        End Select
'
'    Next i
'
'    silFin(nNuc) = L
'
'End Sub
'
'Private Function EsVocal_VA(ByVal c As String) As Boolean
'    EsVocal_VA = InStr("aeiouàèéíòóú", c) > 0
'End Function
'
'Private Function EsVocalForta_VA(ByVal c As String) As Boolean
'    EsVocalForta_VA = InStr("aàeèéoò", c) > 0
'End Function
'
'Private Function EsVocalDebil_VA(ByVal c As String) As Boolean
'    EsVocalDebil_VA = InStr("iíuú", c) > 0
'End Function
'
'Private Function EsDiptongo_VA(ByVal v1 As String, ByVal v2 As String) As Boolean
'
'    If Not EsVocal_VA(v1) Or Not EsVocal_VA(v2) Then Exit Function
'
'    ' Dos fortes ? hiato
'    If EsVocalForta_VA(v1) And EsVocalForta_VA(v2) Then Exit Function
'
'    ' Dos débiles ? diptongo
'    If EsVocalDebil_VA(v1) And EsVocalDebil_VA(v2) Then
'        EsDiptongo_VA = True
'        Exit Function
'    End If
'
'    ' Resto ? diptongo
'    EsDiptongo_VA = True
'
'End Function
'
'Private Function EsTriptongo_VA(ByVal v1 As String, ByVal v2 As String, ByVal v3 As String) As Boolean
'    If EsVocalDebil_VA(v1) And EsVocalForta_VA(v2) And EsVocalDebil_VA(v3) Then
'        EsTriptongo_VA = True
'    End If
'End Function
'
'Private Function EsGrupoAtaque_VA(ByVal g As String) As Boolean
'    Dim AC As Variant
'    AC = Array("pr", "br", "tr", "dr", "cr", "gr", "fr", _
'               "pl", "bl", "cl", "gl", "fl")
'    EsGrupoAtaque_VA = (UBound(Filter(AC, g)) >= 0)
'End Function
'
'
''=====================================================================
'
''Private Sub CalcularTonicas_VA()
''
''    Dim silabas() As String
''    Dim i As Long
''
''    silabas = Split(ObjDTO.SilabasAuto, " | ")
''
''    ' 1) Buscar tilde
''    For i = LBound(silabas) To UBound(silabas)
''        If TieneTilde_VA(silabas(i)) Then
''            ObjDTO.SilabaTonica = CStr(i + 1)
''            Exit Sub
''        End If
''    Next i
''
''    ' 2) Sin tilde ? aplicar reglas generales valencianas
''    Call DetectarTonicaPorReglas_VA(silabas)
''
''End Sub
''
''Private Function TieneTilde_VA(ByVal s As String) As Boolean
''    TieneTilde_VA = (InStr(s, "à") Or InStr(s, "è") Or InStr(s, "é") Or _
''                     InStr(s, "í") Or InStr(s, "ò") Or InStr(s, "ó") Or _
''                     InStr(s, "ú")) > 0
''End Function
''
''Private Sub DetectarTonicaPorReglas_VA(ByRef silabas() As String)
''
''    Dim n As Long
''    n = UBound(silabas) + 1
''
''    Dim palabra As String
''    palabra = Replace(ObjDTO.TextoNormalizado, " ", "")
''
''    Dim ultima As String
''    ultima = Right$(palabra, 1)
''
''    ' Aguda si termina en vocal, vocal+s, en, in
''    If ultima Like "[aeiou]" Or _
''       Right$(palabra, 2) = "en" Or _
''       Right$(palabra, 2) = "in" Then
''
''        ObjDTO.SilabaTonica = CStr(n)
''        Exit Sub
''    End If
''
''    ' Si no ? llana
''    ObjDTO.SilabaTonica = CStr(n - 1)
''
''End Sub
''
''Private Sub CalcularSecundarias_VA()
''
''    Dim n As Long, t As Long
''    Dim lista As String
''
''    n = UBound(Split(ObjDTO.SilabasAuto, " | ")) + 1
''    t = CLng(ObjDTO.SilabaTonica)
''
''    lista = ""
''
''    Dim i As Long
''    For i = 1 To n
''        If i <> t And (i Mod 2 = 1) Then
''            If lista = "" Then
''                lista = CStr(i)
''            Else
''                lista = lista & "," & i
''            End If
''        End If
''    Next i
''
''    ObjDTO.SilabaSecundaria = lista
''
''End Sub
''
''Private Sub MarcarTonicaYSecundariaEnCadena_VA()
''
''    Dim silabas() As String
''    Dim i As Long
''    Dim t As Long
''    Dim sec() As String
''    Dim esSec As Boolean
''
''    silabas = Split(ObjDTO.SilabasAuto, " | ")
''    t = CLng(ObjDTO.SilabaTonica)
''
''    If ObjDTO.SilabaSecundaria <> "" Then
''        sec = Split(ObjDTO.SilabaSecundaria, ",")
''    End If
''
''    For i = LBound(silabas) To UBound(silabas)
''
''        esSec = False
''        If Not IsEmpty(sec) Then
''            If UBound(Filter(sec, CStr(i + 1))) >= 0 Then
''                esSec = True
''            End If
''        End If
''
''        If i + 1 = t Then
''            silabas(i) = "( " & silabas(i) & " )"
''        ElseIf esSec Then
''            silabas(i) = "[ " & silabas(i) & " ]"
''        End If
''
''    Next i
''
''    ObjDTO.SilabasAcentuadas = Join(silabas, " | ")
''
''End Sub
'
'
''================================================================
'
'' ============================================================
''   DETECTAR SÍLABAS TÓNICAS — VALENCIANO
'' ============================================================
'Private Sub CalcularTonicas_VA()
'
'    Dim tGlobal As New Collection
'    Dim elementos As Collection
'    Set elementos = ObtenerPalabrasDesdeSilabasAuto_VA()
'
'    Dim globalIndex As Byte
'    Dim i As Byte
'
'    globalIndex = 0
'
'    For i = 1 To elementos.count
'
'        If TypeName(elementos(i)) = "Collection" Then
'            ' palabra real
'            Dim w As Collection
'            Set w = elementos(i)
'
'            Dim tLocal As Long
'            tLocal = DetectarTonica_VA(w)
'
'            If tLocal > 0 Then
'                tGlobal.Add globalIndex + tLocal
'            End If
'
'            globalIndex = globalIndex + w.count
'
'        Else
'            ' HUECO
'            globalIndex = globalIndex + 1
'        End If
'
'    Next i
'
'    ObjDTO.SilabasTonicas = JoinCollection_VA(tGlobal)
'
'End Sub
'
'' ============================================================
''   DETECTAR SÍLABAS SECUNDARIAS — VALENCIANO
'' ============================================================
'Private Sub CalcularSecundarias_VA()
'
'    Dim sGlobal As New Collection
'    Dim elementos As Collection
'    Set elementos = ObtenerPalabrasDesdeSilabasAuto_VA()
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
'            ' Detectar tónica local
'            tLocal = DetectarTonica_VA(w)
'
'            ' Detectar secundarias locales
'            Dim secs As Collection
'            Set secs = DetectarSecundarias_VA(w, tLocal)
'
'            ' Convertir secundarias locales a globales
'            Dim x As Variant
'            For Each x In secs
'                sGlobal.Add globalIndex + CByte(x)
'            Next x
'
'            globalIndex = globalIndex + w.count
'
'        Else
'            ' HUECO
'            globalIndex = globalIndex + 1
'        End If
'
'    Next i
'
'    ObjDTO.SilabasSecundarias = JoinCollection_VA(sGlobal)
'
'End Sub
'
'' ============================================================
''   MARCAR TÓNICAS Y SECUNDARIAS EN LA CADENA FINAL — VA
'' ============================================================
'Private Sub MarcarTonicaYSecundariaEnCadena_VA()
'
'    Dim sils As Variant
'    Dim i As Byte
'    Dim out() As String
'
'    Dim t As Variant
'    Dim x As Variant
'
'    sils = Split(ObjDTO.SilabasAuto, " | ")
'
'    ReDim out(LBound(sils) To UBound(sils))
'
'    For i = LBound(sils) To UBound(sils)
'        out(i) = sils(i)   ' copia sin marcar
'    Next i
'
'    ' 1) TÓNICAS
'    If ObjDTO.SilabasTonicas <> "" Then
'
'        t = Split(ObjDTO.SilabasTonicas, ",")
'
'        For Each x In t
'            Dim idx As Long
'            idx = CByte(x) - 1   ' arrays base 0
'
'            If idx >= LBound(out) And idx <= UBound(out) Then
'                If Trim$(out(idx)) <> "" Then
'                    out(idx) = "( " & out(idx) & " )"
'                End If
'            End If
'        Next x
'    End If
'
'    ' 2) SECUNDARIAS
'    If ObjDTO.SilabasSecundarias <> "" Then
'
'        Dim s As Variant
'        Dim y As Variant
'        Dim idx2 As Byte
'
'        s = Split(ObjDTO.SilabasSecundarias, ",")
'
'        For Each y In s
'
'            idx2 = CByte(y) - 1
'
'            If idx2 >= LBound(out) And idx2 <= UBound(out) Then
'                If Trim$(out(idx2)) <> "" Then
'                    out(idx2) = "[ " & out(idx2) & " ]"
'                End If
'            End If
'        Next y
'    End If
'
'    ' 3) UNIR RESULTADO
'    ObjDTO.SilabasAcentuadas = Join(out, " | ")
'
'End Sub
'
''-------------------------------------------------------------
''             AUXILIARES TÓNICAS Y SECUNDARIAS — VA
''-------------------------------------------------------------
'Private Function DetectarTonica_VA(w As Collection) As Byte
'
'    Dim ultima As String
'    Dim palabra As String
'    Dim esAguda As Boolean
'    Dim i As Byte
'
'    palabra = ""
'
'    ' 1) Si alguna sílaba tiene tilde ? tónica directa
'    For i = 1 To w.count
'        If TieneTilde_VA(w(i)) Then
'            DetectarTonica_VA = i
'            Exit Function
'        End If
'    Next i
'
'    ' 2) Sin tilde: aplicar regla general valenciana
'    For i = 1 To w.count
'        palabra = palabra & w(i)
'    Next i
'
'    ultima = Right$(palabra, 1)
'    esAguda = False
'
'    ' Aguda si termina en vocal, vocal+s, -en, -in
'    If InStr("aeiouàèéíòóú", ultima) > 0 Then esAguda = True
'    If LCase$(Right$(palabra, 2)) Like "*as" Then esAguda = True
'    If LCase$(Right$(palabra, 2)) Like "*es" Then esAguda = True
'    If LCase$(Right$(palabra, 2)) Like "*is" Then esAguda = True
'    If LCase$(Right$(palabra, 2)) Like "*os" Then esAguda = True
'    If LCase$(Right$(palabra, 2)) Like "*us" Then esAguda = True
'    If LCase$(Right$(palabra, 2)) = "en" Then esAguda = True
'    If LCase$(Right$(palabra, 2)) = "in" Then esAguda = True
'
'    If esAguda Then
'        DetectarTonica_VA = w.count          ' última sílaba
'    ElseIf w.count >= 2 Then
'        DetectarTonica_VA = w.count - 1      ' penúltima
'    Else
'        DetectarTonica_VA = w.count
'    End If
'
'End Function
'
'Private Function DetectarSecundarias_VA(w As Collection, tPos As Byte) As Collection
'
'    Dim secs As New Collection
'    Dim n As Byte
'    Dim pos2 As Byte
'
'    n = w.count
'
'    ' Palabras de 1–3 sílabas ? sin secundaria
'    If n < 4 Then
'        Set DetectarSecundarias_VA = secs
'        Exit Function
'    End If
'
'    ' Primera secundaria SIEMPRE en la sílaba 1
'    secs.Add 1
'
'    ' Palabras de 6+ sílabas ? segunda secundaria
'    If n >= 6 Then
'        pos2 = tPos - 2   ' dos antes de la tónica
'
'        If pos2 > 1 Then
'            secs.Add pos2
'        End If
'    End If
'
'    Set DetectarSecundarias_VA = secs
'
'End Function
'
'' ============================================================
''   OBTENER PALABRAS DESDE SILABAS AUTO — VA
'' ============================================================
'Private Function ObtenerPalabrasDesdeSilabasAuto_VA() As Collection
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
'            ' Es un hueco
'            If palabraActual.count > 0 Then
'                resultado.Add palabraActual
'                Set palabraActual = New Collection
'            End If
'            resultado.Add "HUECO"
'        Else
'            ' Es una sílaba real
'            palabraActual.Add sils(i)
'        End If
'
'    Next i
'
'    ' Última palabra
'    If palabraActual.count > 0 Then resultado.Add palabraActual
'
'    Set ObtenerPalabrasDesdeSilabasAuto_VA = resultado
'
'End Function
'
'' ============================================================
''   JOIN COLLECTION — VA
'' ============================================================
'Private Function JoinCollection_VA(col As Collection) As String
'
'    Dim arr() As String
'    Dim i As Byte
'
'    If col Is Nothing Then
'        JoinCollection_VA = ""
'        Exit Function
'    End If
'
'    If col.count = 0 Then
'        JoinCollection_VA = ""
'        Exit Function
'    End If
'
'    ReDim arr(1 To col.count)
'
'    For i = 1 To col.count
'        arr(i) = CStr(col(i))
'    Next i
'
'    JoinCollection_VA = Join(arr, ",")
'
'End Function
'
'' ============================================================
''   VOCAL CON TILDE — VALENCIANO
'' ============================================================
'Private Function TieneTilde_VA(ByVal silaba As String) As Boolean
'
'    TieneTilde_VA = (InStr(silaba, "à") > 0 Or _
'                     InStr(silaba, "è") > 0 Or _
'                     InStr(silaba, "é") > 0 Or _
'                     InStr(silaba, "í") > 0 Or _
'                     InStr(silaba, "ò") > 0 Or _
'                     InStr(silaba, "ó") > 0 Or _
'                     InStr(silaba, "ú") > 0)
'
'End Function


