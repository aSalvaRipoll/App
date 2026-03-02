Attribute VB_Name = "bas_Motor_CA_Main"

Option Compare Database
Option Explicit

' ============================================================
'   MOTOR DE SILABEO 4.1 — ORTOGRÁFICO + MORFOSILÁBICO + ACENTUACIÓN
'   VERSIÓN CATALÁN
'
'   Autor: Alba Salvá Ripoll + Copilot
'
'   Este módulo implementa:
'   ? Silabeo ortográfico completo
'   ? Silabeo morfosilábico (prefijos)
'   ? Diptongos, triptongos, hiatos con tilde
'   ? Dígrafos indivisibles
'   ? Casos especiales: qu, gu, güe/güi, tl
'   ? Sílaba tónica
'   ? Acentuación secundaria
'   ? Tracking de depuración (DebugMotor)
'
'   Variables de control:
'       usarSilabeoMorfologico   ? True = usar prefijos
'       modoPrefijosEstrictos    ? True = usar lista interna completa
'       respetarPrefijos         ? True = separar prefijo como sílaba
'       DebugMotor                ? True = activar logs
'
' ============================================================

Private usarSilabeoMorfologico As Boolean
Private modoPrefijosEstrictos As Boolean
Private respetarPrefijos As Boolean

' Prefijos personalizados (modo flexible)
Private prefijosPersonalizados As Variant

' Prefijos canónicos (modo estricto)
Private prefijosEstrictos As Variant

Private prefijosCargados As Boolean

Private Const strSQL As String = "SELECT Prefijo FROM qryPrefijos " & _
          "WHERE Activo = 1 " & _
            "AND Tipo Like 'auténtico' " & _
            "AND [ca] = true " & _
          "ORDER BY Len(Prefijo) DESC, Prefijo ASC"


' ============================================================
'   ENTRADA PRINCIPAL DEL MOTOR (CATALÁN)
' ============================================================
' ----------------------------------------------------------------
' Procedimiento: Entrada_Motor_CA
' Propósito:     Punto de entrada al motor fonético del español general
' Tipo proc.:    Function
' Acceso proc.:  Public

' Parameter Texto (String): Texto que se recibe (nombre o apellido español general)

' Tipo retorno: String -> Texto que contiene la lista de fonemas
'   resultado de la conversión

' Autor:        Alba Salvá
' Fecha:        16/02/2026
' ----------------------------------------------------------------
Public Function Entrada_Motor_CA(texto As String) As String

    Set ObjDTO = New clsDTO_Motor

'    DebugMotor = True
'    DebugDTO = False
    
    ' 1) Asignamos el texto recibido y
    '    Normalización (dentro del DTO)
    ObjDTO.TextoOriginal = texto
    ObjDTO.NormalizaEntrada

    ' 2) Silabeo automático
    Call SilabearFrase_CA
    
    ' 3) Detectar tónicas
    Call CalcularTonicas_CA
    
    ' 4) Detectar secundarias
    Call CalcularSecundarias_CA
    
    ' 5) Marcar Tónicas y Secundarias
    Call MarcarTonicaYSecundariaEnCadena_CA
    
    ObjDTO.SilabasFinal = ObjDTO.SilabasAcentuadas
    
    ' 6) Generar fonética
    Call ConstruirCadenaFonemas_CA
    
    '    << Eliminar en producción >>
    Call MF_DebugDTO("Silabear")
    
    ' 7) Devolver resultado (texto plano)
    '    << Eliminar en producción >>
    Entrada_Motor_CA = ObjDTO.SilabasAuto

End Function

Private Sub SilabearFrase_CA()

    Dim frase As String
    Dim palabras() As String
    Dim resultado As String
    Dim i As Long
    Dim limpia As String
    Dim sil As String

    usarSilabeoMorfologico = True
    modoPrefijosEstrictos = True
    respetarPrefijos = True

    frase = ObjDTO.TextoNormalizado

    palabras = Split(frase, " ")

    For i = LBound(palabras) To UBound(palabras)
        limpia = Trim$(palabras(i))
        If limpia <> "" Then
            sil = SilabearPalabra_CA(limpia)
            If resultado = "" Then
                resultado = sil
            Else
                resultado = resultado & " |   | " & sil
            End If
        End If
    Next i

    ObjDTO.SilabasAuto = resultado

End Sub

Private Function SilabearPalabra_CA(ByVal texto As String) As String
    Dim t As String

    t = LCase$(Trim$(texto))

    If usarSilabeoMorfologico Then
        SilabearPalabra_CA = SilabearMorfologico_CA(t)
    Else
        SilabearPalabra_CA = SilabearOrtog_CA(t)
    End If

End Function

Private Function SilabearOrtog_CA(ByVal t As String) As String
    Dim nucIni() As Byte, nucFin() As Byte
    Dim silIni() As Byte, silFin() As Byte
    Dim nNuc As Byte, i As Byte
    Dim silabas() As String


    If Len(Trim$(t)) < 2 Then
        If t = "y" Then
            SilabearOrtog_CA = "y"
            Exit Function
        ElseIf t = "i" Then
            SilabearOrtog_CA = "i"
            Exit Function
        End If
    End If

    LocalizarNucleosOrtog_CA t, nucIni, nucFin, nNuc

    ReDim silIni(1 To nNuc)
    ReDim silFin(1 To nNuc)

    CalcularSilabas_CA t, nucIni, nucFin, nNuc, silIni, silFin

    ReDim silabas(1 To nNuc)
    For i = 1 To nNuc
        silabas(i) = Mid$(t, silIni(i), silFin(i) - silIni(i) + 1)
        If DebugMotor Then
            addLog "Sílaba " & i & ": " & silabas(i)
        End If
    Next i

    SilabearOrtog_CA = Join(silabas, " | ")
End Function

Private Function SilabearMorfologico_CA(ByVal t As String) As String
    Dim pref As String
    Dim resto As String

    If Not respetarPrefijos Then
        SilabearMorfologico_CA = SilabearOrtog_CA(t)
        Exit Function
    End If

    pref = DetectarPrefijo_CA(t)

    If pref = "" Then
        SilabearMorfologico_CA = SilabearOrtog_CA(t)
        Exit Function
    End If

    resto = Mid$(t, Len(pref) + 1)

    SilabearMorfologico_CA = pref & " | " & SilabearOrtog_CA(resto)
End Function

Private Function DetectarPrefijo_CA(ByVal t As String) As String
    Dim p As Variant
    Dim resto As String
    Dim primera As String, segunda As String

    If Not prefijosCargados_CA Then DetectarPrefijo_CA

    For Each p In prefijosEstrictos_CA

        ' 1) El prefijo debe coincidir con el inicio
        If Left$(t, Len(p)) <> p Then GoTo SiguienteP

        ' 2) Obtener el resto de la palabra
        resto = Mid$(t, Len(p) + 1)
        If Len(resto) = 0 Then GoTo SiguienteP

        primera = Left$(resto, 1)
        segunda = Mid$(resto, 2, 1)

        ' ============================================================
        '   CASO ESPECIAL: PREFIJO "a-" (privativo)
        ' ============================================================
        If p = "a" Then

            ' 2.1) La base debe empezar por vocal
            If Not EsVocal_CA(primera) Then
                GoTo SiguienteP
            End If

            ' 2.2) No puede romper un diptongo
            If EsDiptongo_CA(Left$(t, 1), Mid$(t, 2, 1)) Then
                GoTo SiguienteP
            End If

            ' 2.3) El ataque resultante debe ser válido
            If Not AtaqueSilabicoValido(primera, segunda) Then
                GoTo SiguienteP
            End If
            
            DetectarPrefijo_CA = p
            Exit Function
        End If

        ' ============================================================
        '   PREFIJOS NORMALES (anti-, inter-, sub-, re-, etc.)
        ' ============================================================

        ' 3) El ataque resultante debe ser válido
        If Not AtaqueSilabicoValido(primera, segunda) Then
            GoTo SiguienteP
        End If

        DetectarPrefijo_CA = p
        Exit Function

SiguienteP:
    Next p

    DetectarPrefijo_CA = ""
End Function

Private Function AtaqueSilabicoValido_CA(ByVal a As String, ByVal b As String) As Boolean
    Dim grupo As String
    grupo = LCase$(a & b)

    Select Case grupo
        Case "pr", "pl", "br", "bl", "tr", "dr", "cr", "cl", "gr", "gl", "fr", "fl"
            AtaqueSilabicoValido_CA = True

        ' En catalán central NO se permiten ataques con sC en prefijos
        ' por tanto NO añadimos: sp, st, sc, sf, sm, sn, sl, sr

        Case Else
            AtaqueSilabicoValido_CA = False
    End Select
End Function

'Private Function DetectarPrefijo_CA(ByVal t As String) As String
'    Dim p As Variant
'
'    If Not prefijosCargados Then CargarPrefijos_CA
'
'    For Each p In prefijosEstrictos
'        If Len(t) = Len(p) Then Exit For
'        If Left$(t, Len(p)) = p Then
'            DetectarPrefijo_CA = p
'            Exit Function
'        End If
'    Next p
'
'    DetectarPrefijo_CA = ""
'End Function

Private Sub CargarPrefijos_CA()
    Dim rs As dao.Recordset
    Dim sql As String
    Dim i As Long

    If prefijosCargados Then Exit Sub

'    sql = "SELECT Prefijo FROM tbmPrefijos " & _
          "WHERE Activo = 1 " & _
            "AND Tipo Like 'auténtico' " & _
            "AND [es-ca] = true " & _
          "ORDER BY Len(Prefijo) DESC, Prefijo ASC"

    sql = strSQL
    
    Set rs = CurrentDb.OpenRecordset(sql)

    If Not rs.EOF Then
        rs.MoveLast
        ReDim prefijosEstrictos(1 To rs.RecordCount)
        rs.MoveFirst

        i = 1
        Do Until rs.EOF
            prefijosEstrictos(i) = LCase$(rs!Prefijo)
            i = i + 1
            rs.MoveNext
        Loop
    End If

    rs.Close
    prefijosCargados = True
End Sub

Private Sub LocalizarNucleosOrtog_CA(ByVal t As String, _
                                  ByRef nucIni() As Byte, _
                                  ByRef nucFin() As Byte, _
                                  ByRef nNuc As Byte)

    Dim i As Byte, L As Byte
    Dim c1 As String, c2 As String, c3 As String

    L = Len(t)
    ReDim nucIni(1 To L)
    ReDim nucFin(1 To L)
    nNuc = 0

    If DebugMotor Then
        addLog
        addLog "---------------------------------------"
        addLog " Procedimiento LocalizarNucleosOrtog_CA"
    End If

    i = 1
    Do While i <= L
        c1 = Mid$(t, i, 1)

        If EsVocal_CA(c1) Then

            ' Intentar triptongo
            If i + 2 <= L Then
                c2 = Mid$(t, i + 1, 1)
                c3 = Mid$(t, i + 2, 1)
                If EsTriptongo_CA(c1, c2, c3) Then
                    nNuc = nNuc + 1
                    nucIni(nNuc) = i
                    nucFin(nNuc) = i + 2
                    If DebugMotor Then
                        addLog "Triptongo: " & c1 & c2 & c3
                    End If
                    i = i + 3
                    GoTo Siguiente
                End If
            End If

            ' Intentar diptongo
            If i + 1 <= L Then
                c2 = Mid$(t, i + 1, 1)
                If EsDiptongo_CA(c1, c2) Then
                    nNuc = nNuc + 1
                    nucIni(nNuc) = i
                    nucFin(nNuc) = i + 1
                    If DebugMotor Then
                        addLog "Diptongo: " & c1 & c2
                    End If
                    i = i + 2
                    GoTo Siguiente
                End If
            End If

            ' Vocal sola
            nNuc = nNuc + 1
            nucIni(nNuc) = i
            nucFin(nNuc) = i
            If DebugMotor Then
                addLog "Vocal sola: " & c1
            End If
            i = i + 1

        Else
            i = i + 1
        End If

Siguiente:
    Loop

    If DebugMotor Then
        addLog "Total núcleos: " & nNuc
        addLog " Fin LocalizarNucleosOrtog_CA"
        addLog "---------------------------------------"
    End If
End Sub

Private Sub CalcularSilabas_CA(ByVal t As String, _
                            ByRef nucIni() As Byte, _
                            ByRef nucFin() As Byte, _
                            ByVal nNuc As Byte, _
                            ByRef silIni() As Byte, _
                            ByRef silFin() As Byte)

' Con control de dígrafos

    Dim i As Byte, L As Byte
    Dim a As Byte, b As Byte
    Dim k As Byte
    Dim c1 As String, c2 As String, c3 As String, grupo As String

    L = Len(t)
    silIni(1) = 1

    For i = 1 To nNuc - 1
        a = nucFin(i)
        b = nucIni(i + 1)

        k = IIf(b > a + 1, b - a - 1, 0)

        If DebugMotor Then
            addLog
            addLog "---- Frontera entre núcleo " & i & " y " & (i + 1)
            addLog "Consonantes entre medias: " & k
        End If

        Select Case k

            Case 0
                silFin(i) = a
                silIni(i + 1) = a + 1

            Case 1
                silFin(i) = a
                silIni(i + 1) = a + 1

            Case 2
                c1 = Mid$(t, a + 1, 1)
                c2 = Mid$(t, a + 2, 1)
                grupo = c1 & c2

                ' Dígrafos indivisibles
                If grupo = "rr" Or grupo = "ll" Or grupo = "ch" Then
                    silFin(i) = a
                    silIni(i + 1) = a + 1
                    GoTo Siguiente
                End If

                ' tl ? siempre se separa
                If grupo = "tl" Then
                    silFin(i) = a + 1
                    silIni(i + 1) = a + 2
                    GoTo Siguiente
                End If

                ' Grupos de ataque válidos
                If EsGrupoAtaque_CA(grupo) Then
                    silFin(i) = a
                    silIni(i + 1) = a + 1
                Else
                    silFin(i) = a + 1
                    silIni(i + 1) = a + 2
                End If

'            Case 3
'                silFin(i) = a + 1
'                silIni(i + 1) = a + 2

            Case 3
                ' Tres consonantes entre vocales: C1 C2 C3
                ' Regla IEC:
                ' - Si C2 puede cerrar sílaba ? C1C2 | C3
                ' - Si C2 NO puede cerrar sílaba ? C1 | C2C3

                c1 = Mid$(t, a + 1, 1)
                c2 = Mid$(t, a + 2, 1)
                c3 = Mid$(t, a + 3, 1)

                If PuedeCerrarSilaba_CA(c2) Then
                    ' C1C2 | C3  ? ej. Mont-ju-ïc
                    silFin(i) = a + 2
                    silIni(i + 1) = a + 3
                Else
                    ' C1 | C2C3  ? ej. mons-tre
                    silFin(i) = a + 1
                    silIni(i + 1) = a + 2
                End If

'            Case 3
'
'                C1 = Mid$(t, a + 1, 1)
'                C2 = Mid$(t, a + 2, 1)
'                C3 = Mid$(t, a + 3, 1)
'
'                ' --- REGLA ESPECIAL PARA L·L + VOCAL ---
'                ' Si PreferirIlLu = True, agrupar siempre "l·lu" como una sola síl·laba
'                If PreferirIlLu Then
'                    If C1 = "l" And C2 = "·" And C3 = "l" Then
'                        ' La siguiente letra es vocal?
'                        If b <= Len(t) Then
'                            Dim C4 As String
'                            C4 = Mid$(t, a + 4, 1)
'                            If EsVocal_CA(C4) Then
'                                ' Forzar il·lu como una sola síl·laba
'                                silFin(i) = a + 4
'                                silIni(i + 1) = a + 5
'                                GoTo Siguiente
'                            End If
'                        End If
'                    End If
'                End If
'                ' --- FIN REGLA ESPECIAL ---
'
'                ' REGLA GENERAL IEC PARA 3 CONSONANTES
'                If PuedeCerrarSilaba_CA(C2) Then
'                    silFin(i) = a + 2
'                    silIni(i + 1) = a + 3
'                Else
'                    silFin(i) = a + 1
'                    silIni(i + 1) = a + 2
'                End If


            Case Else
                silFin(i) = a + 2
                silIni(i + 1) = a + 3

        End Select

Siguiente:
    Next i

    silFin(nNuc) = L
End Sub

' ============================================================
'   FUNCIONES DE VOCAL — CATALÁN (versión robusta y coherente)
' ============================================================

' Vocal general (todas las vocales catalanas, incluidas variantes)
Private Function EsVocal_CA(ByVal c As String) As Boolean
    ' Incluye ä para que no rompa el silabeo
    EsVocal_CA = InStr("aeiouàèéíïòóúüä", LCase$(c)) > 0
End Function

' Vocal fuerte (a, e, o abiertas o cerradas, incluidas variantes)
Private Function EsVocalForta_CA(ByVal c As String) As Boolean
    ' Añadimos ä porque es una variante de a abierta
    EsVocalForta_CA = InStr("aàäeèéoò", LCase$(c)) > 0
End Function

' Vocal débil (i, u y variantes)
Private Function EsVocalDebil_CA(ByVal c As String) As Boolean
    EsVocalDebil_CA = InStr("iíïuúü", LCase$(c)) > 0
End Function

' Vocal débil tónica (solo í, ú)
Private Function EsVocalDebilTonica_CA(ByVal c As String) As Boolean
    EsVocalDebilTonica_CA = InStr("íú", LCase$(c)) > 0
End Function

Private Function EsDiptongo_CA(ByVal v1 As String, ByVal v2 As String) As Boolean

    If Not EsVocal_CA(v1) Or Not EsVocal_CA(v2) Then Exit Function

    ' Dos fortes ? hiat
    If EsVocalForta_CA(v1) And EsVocalForta_CA(v2) Then Exit Function

    ' Si la segunda vocal tiene dièresi, nunca forma diptongo, es hiato
    If v2 = "ï" Or v2 = "ü" Then
        EsDiptongo_CA = False
        Exit Function
    End If

    ' Dos dèbils ? diptong
    If EsVocalDebil_CA(v1) And EsVocalDebil_CA(v2) Then
        EsDiptongo_CA = True
        Exit Function
    End If

    ' Dèbil tònica + forta ? hiat
    If (EsVocalDebilTonica_CA(v1) And EsVocalForta_CA(v2)) _
    Or (EsVocalForta_CA(v1) And EsVocalDebilTonica_CA(v2)) Then
        Exit Function
    End If

    ' Resta ? diptong
    EsDiptongo_CA = True

End Function

Private Function EsTriptongo_CA(ByVal v1 As String, ByVal v2 As String, ByVal v3 As String) As Boolean
    If EsVocalDebil_CA(v1) And Not EsVocalDebilTonica_CA(v1) _
       And EsVocalForta_CA(v2) _
       And EsVocalDebil_CA(v3) And Not EsVocalDebilTonica_CA(v3) Then
        EsTriptongo_CA = True
    End If
End Function

Private Function EsGrupoAtaque_CA(ByVal g As String) As Boolean
    Dim AC As Variant
    AC = Array("pr", "br", "tr", "dr", "cr", "gr", "fr", _
               "pl", "bl", "cl", "gl", "fl")
    EsGrupoAtaque_CA = (UBound(Filter(AC, g)) >= 0)
End Function

Private Function PuedeCerrarSilaba_CA(ByVal c As String) As Boolean
    ' Consonantes que NO pueden cerrar sílaba en catalán ortográfico
    ' (r y l cuando forman dígrafo, h muda)
    PuedeCerrarSilaba_CA = Not (c = "r" Or c = "l" Or c = "h")
End Function

' Vocal con tilde (para detectar sílaba tónica)
Private Function TieneTilde_CA(ByVal silaba As String) As Boolean
    ' Añadimos ä si la usas como marca de acento (como en "mònicä")
    TieneTilde_CA = (InStr(silaba, "à") > 0 Or _
                     InStr(silaba, "è") > 0 Or _
                     InStr(silaba, "é") > 0 Or _
                     InStr(silaba, "í") > 0 Or _
                     InStr(silaba, "ï") > 0 Or _
                     InStr(silaba, "ò") > 0 Or _
                     InStr(silaba, "ó") > 0 Or _
                     InStr(silaba, "ú") > 0 Or _
                     InStr(silaba, "ä") > 0)
End Function

'=================================================================
'=================================================================
'                 SECCIÓN MÓDULO ACENTOS
'=================================================================
'=================================================================

' ============================================================
'   DETECTAR SÍLABAS TÓNICAS
' ============================================================
Private Sub CalcularTonicas_CA()

    Dim tGlobal As New Collection
    Dim elementos As Collection
    Set elementos = ObtenerPalabrasDesdeSilabasAuto_CA()

    Dim globalIndex As Byte
    Dim i As Byte

    globalIndex = 0

    For i = 1 To elementos.count

        If TypeName(elementos(i)) = "Collection" Then
            ' palabra real
            Dim w As Collection
            Set w = elementos(i)

            Dim tLocal As Long
            tLocal = DetectarTonica_CA(w)

            If tLocal > 0 Then
                tGlobal.Add globalIndex + tLocal
            End If

            globalIndex = globalIndex + w.count

        Else
            ' HUECO
            globalIndex = globalIndex + 1
        End If

    Next i

    ObjDTO.SilabasTonicas = JoinCollection_CA(tGlobal)

End Sub


' ============================================================
'   DETECTAR SÍLABAS SECUNDARIAS (pueden ser varias)
' ============================================================
Private Sub CalcularSecundarias_CA()

    Dim sGlobal As New Collection
    Dim elementos As Collection
    Set elementos = ObtenerPalabrasDesdeSilabasAuto_CA()

    Dim globalIndex As Byte
    Dim i As Byte
    Dim tLocal As Byte

    globalIndex = 0

    For i = 1 To elementos.count

        If TypeName(elementos(i)) = "Collection" Then

            Dim w As Collection
            Set w = elementos(i)

            ' Detectar tónica local
            tLocal = DetectarTonica_CA(w)

            ' Detectar secundarias locales
            Dim secs As Collection
            Set secs = DetectarSecundarias_CA(w, tLocal)

            ' Convertir secundarias locales a globales
            Dim x As Variant
            For Each x In secs
                sGlobal.Add globalIndex + CByte(x)
            Next x

            globalIndex = globalIndex + w.count

        Else
            ' HUECO
            globalIndex = globalIndex + 1
        End If

    Next i

    ObjDTO.SilabasSecundarias = JoinCollection_CA(sGlobal)

End Sub


' ============================================================
'   MARCAR TÓNICAS Y SECUNDARIAS EN LA CADENA FINAL
' ============================================================
Private Sub MarcarTonicaYSecundariaEnCadena_CA()

    Dim sils As Variant
    Dim i As Byte
    Dim out() As String

    Dim t As Variant
    Dim x As Variant

    sils = Split(ObjDTO.SilabasAuto, " | ")

    ReDim out(LBound(sils) To UBound(sils))

    For i = LBound(sils) To UBound(sils)
        out(i) = sils(i)   ' copia sin marcar
    Next i

    ' 1) TÓNICAS
    If ObjDTO.SilabasTonicas <> "" Then

        t = Split(ObjDTO.SilabasTonicas, ",")

        For Each x In t
            Dim idx As Long
            idx = CByte(x) - 1   ' arrays base 0

            If idx >= LBound(out) And idx <= UBound(out) Then
                If Trim$(out(idx)) <> "" Then
                    out(idx) = "( " & out(idx) & " )"
                End If
            End If
        Next x
    End If

    ' 2) SECUNDARIAS
    If ObjDTO.SilabasSecundarias <> "" Then

        Dim s As Variant
        Dim y As Variant
        Dim idx2 As Byte

        s = Split(ObjDTO.SilabasSecundarias, ",")

        For Each y In s

            idx2 = CByte(y) - 1

            If idx2 >= LBound(out) And idx2 <= UBound(out) Then
                If Trim$(out(idx2)) <> "" Then
                    out(idx2) = "[ " & out(idx2) & " ]"
                End If
            End If
        Next y
    End If

    ' 3) UNIR RESULTADO
    ObjDTO.SilabasAcentuadas = Join(out, " | ")

End Sub

'-------------------------------------------------------------
'             AUXILIARES TÓNICAS Y SECUNDARIAS
'-------------------------------------------------------------
Private Function DetectarTonica_CA(w As Collection) As Byte

    Dim ultima As String
    Dim palabra As String
    Dim terminaLlana As Boolean
    Dim i As Byte
    
    palabra = ""

    ' 1) Si alguna síl·laba té accent gràfic ? tònica directa
    
    For i = 1 To w.count
        If TieneTilde_CA(w(i)) Then
            DetectarTonica_CA = i
            Exit Function
        End If
    Next i

    ' 2) Sense accent: aplicar regla general catalana
    'palabra = JoinCollection_CA(w) '.ToArray, "")
    ' Reconstruir la paraula a partir de les síl·labes
    For i = 1 To w.count
        palabra = palabra & w(i)
    Next i

    ultima = Right$(palabra, 1)

    terminaLlana = False

    ' Vocal o terminacions planes catalanes
    If InStr("aeiouàèéíïòóúü", ultima) > 0 Then terminaLlana = True
    If LCase$(Right$(palabra, 2)) Like "*as" Then terminaLlana = True
    If LCase$(Right$(palabra, 2)) Like "*es" Then terminaLlana = True
    If LCase$(Right$(palabra, 2)) Like "*is" Then terminaLlana = True
    If LCase$(Right$(palabra, 2)) Like "*os" Then terminaLlana = True
    If LCase$(Right$(palabra, 2)) Like "*us" Then terminaLlana = True
    If LCase$(Right$(palabra, 2)) Like "*en" Then terminaLlana = True
    If LCase$(Right$(palabra, 2)) Like "*in" Then terminaLlana = True

    If terminaLlana And w.count >= 2 Then
        DetectarTonica_CA = w.count - 1   ' penúltima síl·laba
    Else
        DetectarTonica_CA = w.count       ' última síl·laba
    End If

End Function

Private Function DetectarSecundarias_CA(w As Collection, tPos As Byte) As Collection

    Dim secs As New Collection
    Dim n As Byte
    Dim pos2 As Byte

    n = w.count

    ' Palabras de 1–3 sílabas ? sin secundaria
    If n < 4 Then
        Set DetectarSecundarias_CA = secs
        Exit Function
    End If

    ' Primera secundaria SIEMPRE en la sílaba 1
    secs.Add 1

    ' Palabras de 6+ sílabas ? segunda secundaria
    If n >= 6 Then
        pos2 = tPos - 2   ' dos antes de la tónica

        If pos2 > 1 Then
            secs.Add pos2
        End If
    End If

    Set DetectarSecundarias_CA = secs

End Function


' ============================================================
'   OBTENER PALABRAS DESDE SILABAS AUTO
' ============================================================
Private Function ObtenerPalabrasDesdeSilabasAuto_CA() As Collection

    Dim resultado As New Collection
    Dim palabraActual As New Collection

    Dim sils As Variant
    sils = Split(ObjDTO.SilabasAuto, " | ")

    Dim i As Byte
    For i = LBound(sils) To UBound(sils)

        If Trim$(sils(i)) = "" Then
            ' Es un hueco
            If palabraActual.count > 0 Then
                resultado.Add palabraActual
                Set palabraActual = New Collection
            End If
            resultado.Add "HUECO"
        Else
            ' Es una sílaba real
            palabraActual.Add sils(i)
        End If

    Next i

    ' Última palabra
    If palabraActual.count > 0 Then resultado.Add palabraActual

    Set ObtenerPalabrasDesdeSilabasAuto_CA = resultado

End Function


' ============================================================
'   JOIN COLLECTION
' ============================================================
Private Function JoinCollection_CA(col As Collection) As String

    Dim arr() As String
    Dim i As Byte

    If col Is Nothing Then
        JoinCollection_CA = ""
        Exit Function
    End If

    If col.count = 0 Then
        JoinCollection_CA = ""
        Exit Function
    End If

    ReDim arr(1 To col.count)

    For i = 1 To col.count
        arr(i) = CStr(col(i))
    Next i

    JoinCollection_CA = Join(arr, ",")

End Function




