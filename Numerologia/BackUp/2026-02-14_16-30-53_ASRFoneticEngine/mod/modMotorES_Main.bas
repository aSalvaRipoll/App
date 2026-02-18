Attribute VB_Name = "modMotorES_Main"
'Option Compare Database
'Option Explicit
'
'Private IndiceSilabaActual As Long
'Private EsTonicaActual As Boolean
'Private GrafAnterior As String
'Private GrafActual As String
'Private GrafSiguiente As String
'
'Private objDTO As clsMotorFonetico
'
'' ============================================================
''   ENTRADA PRINCIPAL DEL MOTOR
'' ============================================================
'Public Function EntradaMotor_ES(texto As String) As String
'
'    Set objDTO = New clsMotorFonetico
'    objDTO.TextoOriginal = texto
'
'    Call NormalizarEntrada
'    Call SilabearAuto
'    Call MF_DebugDTO("SilabearAuto")
'    GoTo Fin
'
'    Call DetectarTonica
'    Call MarcarSilabasTonicas
'    Call DetectarSecundarias
'    Call MarcarSilabasSecundarias
'
'
'    Call RevisionSilabeo
'    Call ConvertirSilabasAFonemas
'
'Fin:
'    EntradaMotor_ES = Replace(objDTO.TextoFinal, " ", "")
'
'End Function
'
'' ============================================================
''   1- NORMALIZACIÓN
'' ============================================================
'Private Sub NormalizarEntrada()
'
'    Dim s As String
'
'    s = objDTO.TextoOriginal
'
'    Do While InStr(s, "  ") > 0
'        s = Replace(s, "  ", " ")
'    Loop
'
'    s = Replace(s, vbTab, " ")
'    s = Replace(s, vbCr, "")
'    s = Replace(s, vbLf, "")
'
'    s = Replace(s, "–", "-")
'    s = Replace(s, "—", "-")
'    s = Replace(s, "“", """")
'    s = Replace(s, "”", """")
'
'    s = Replace(s, " -", "-")
'    s = Replace(s, "- ", "-")
'
'    s = LCase$(Trim$(s))
'
'    objDTO.TextoNormalizado = s
'
'End Sub
'
'' ============================================================
''   2- SILABEO AUTOMÁTICO (VERSIÓN MODULAR) V X22
'' ============================================================
'Private Sub SilabearAuto()
'
'    Dim texto As String
'    Dim arr() As String
'    Dim salida As String
'
'
'    Debug.Print
'
'    texto = Trim$(objDTO.TextoNormalizado)
'    If Len(texto) = 0 Then
'        objDTO.SilabasAuto = ""
'        Exit Sub
'    End If
'
'    ' 1) SEGMENTACIÓN BASE
'    Call SilabearAuto_Base(texto, arr)
'    Debug.Print "SilabearAuto_Base finalizado"
'
'    ' 2) CORRECCIONES INTERNAS
'    Call SilabearAuto_Correcciones(arr)
'    Debug.Print "SilabearAuto_Correcciones finalizado"
'
'    ' 3) ENSAMBLADO FINAL
'    Call SilabearAuto_EnsambladoFinal(arr, salida)
'    Debug.Print "SilabearAuto_EnsambladoFinal finalizado"
'
'    objDTO.SilabasAuto = salida
'
'End Sub
'
'' ============================================================
''   MÓDULO 1 — SEGMENTACIÓN BASE (SilabearAuto_Base) V X21
''   Orden fonológico correcto:
''   1) Triptongo
''   2) Hiato
''   3) Diptongo
''   4) Grupos inseparables (dr, gr, tr…)
''   5) CCC
''   6) VCV
''   7) CCV
'' ============================================================
'Private Sub SilabearAuto_Base(ByVal texto As String, ByRef arr() As String)
'
'    Dim col As New Collection
'    Dim i As Long, ini As Long
'    Dim c1 As String, c2 As String, c3 As String
'    Dim par As String
'
'    ini = 1
'
'    For i = 2 To Len(texto)
'
'        c1 = Mid$(texto, i - 1, 1)
'        c2 = Mid$(texto, i, 1)
'        If i < Len(texto) Then c3 = Mid$(texto, i + 1, 1) Else c3 = ""
'        par = c1 & c2
'
'        ' --------------------------------------------------------
'        ' 0. ESPACIOS
'        ' --------------------------------------------------------
'        If c1 = " " Then
'            If i - 2 >= ini Then col.Add Array(ini, i - 2)
'            ini = i
'            GoTo siguiente
'        End If
'
'        If c2 = " " Then
'            col.Add Array(ini, i - 1)
'            col.Add Array(i, i)
'            ini = i + 1
'            GoTo siguiente
'        End If
'
'        ' --------------------------------------------------------
'        ' 1. DÍGRAFOS
'        ' --------------------------------------------------------
'        If par = "ch" Or par = "ll" Or par = "rr" Then GoTo siguiente
'
'        ' --------------------------------------------------------
'        ' 2. NSP / NST
'        ' --------------------------------------------------------
'        If c1 = "n" And c2 = "s" Then
'            If c3 = "p" Or c3 = "t" Then
'                col.Add Array(ini, i)
'                ini = i + 1
'                GoTo siguiente
'            End If
'        End If
'
'        ' --------------------------------------------------------
'        ' 3. TRIPTONGO (PRIORIDAD MÁXIMA)
'        ' --------------------------------------------------------
'        If i < Len(texto) - 1 Then
'            If EsTriptongo(c1, c2, Mid$(texto, i + 1, 1)) Then GoTo siguiente
'        End If
'
'        ' --------------------------------------------------------
'        ' 4. HIATOS
'        ' --------------------------------------------------------
'        If EsVocal_ES(c1) And EsVocal_ES(c2) Then
'            If EsDiptongo(c1, c2) Then GoTo siguiente
'            If EsHiato(c1, c2) Then
'                col.Add Array(ini, i - 1)
'                ini = i
'                GoTo siguiente
'            End If
'            col.Add Array(ini, i - 1)
'            ini = i
'            GoTo siguiente
'        End If
'
'        ' --------------------------------------------------------
'        ' 5. DIPTONGO
'        ' --------------------------------------------------------
'        If EsDiptongo(c1, c2) Then GoTo siguiente
'
'        ' --------------------------------------------------------
'        ' 6. GRUPOS INSEPARABLES (dr, gr, tr…)
'        ' --------------------------------------------------------
'        If EsConsonante_ES(c1) And EsConsonante_ES(c2) Then
'            If EsGrupoInseparable_ES(par) Then GoTo siguiente
'        End If
'
'        ' --------------------------------------------------------
'        ' 7. CCC
'        ' --------------------------------------------------------
'        If EsCCC(c1, c2, c3) Then
'            col.Add Array(ini, i - 1)
'            ini = i
'            GoTo siguiente
'        End If
'
'        ' --------------------------------------------------------
'        ' 8. V + C + V TILDADA (ción)
'        ' --------------------------------------------------------
'        If EsVocal_ES(c1) And EsConsonante_ES(c2) And EsVocal_ES(c3) Then
'            If c3 Like "[áéíóú]" Then GoTo siguiente
'        End If
'
'        ' --------------------------------------------------------
'        ' 9. V + RR
'        ' --------------------------------------------------------
'        If EsVocal_ES(c1) And c2 = "r" Then
'            If i < Len(texto) Then
'                If Mid$(texto, i + 1, 1) = "r" Then
'                    col.Add Array(ini, i - 1)
'                    ini = i
'                    GoTo siguiente
'                End If
'            End If
'        End If
'
'        ' --------------------------------------------------------
'        ' 10. ST después de NS
'        ' --------------------------------------------------------
'        If c1 = "s" And c2 = "t" Then
'            If i > 2 Then
'                If Mid$(texto, i - 2, 1) = "n" Then GoTo siguiente
'            End If
'        End If
'
'        ' --------------------------------------------------------
'        ' 11. VCV ? V | CV
'        ' --------------------------------------------------------
'        If EsVocal_ES(c1) And EsConsonante_ES(c2) And EsVocal_ES(c3) Then
'            col.Add Array(ini, i - 1)
'            ini = i
'            GoTo siguiente
'        End If
'
'        ' --------------------------------------------------------
'        ' 12. CCV ? C | CV
'        ' --------------------------------------------------------
'        If EsConsonante_ES(c1) And EsConsonante_ES(c2) And EsVocal_ES(c3) Then
'            If Not EsGrupoInseparable_ES(par) Then
'                col.Add Array(ini, i - 1)
'                ini = i
'                GoTo siguiente
'            End If
'        End If
'
'siguiente:
'    Next i
'
'    If ini <= Len(texto) Then col.Add Array(ini, Len(texto))
'
'    ReDim arr(1 To col.Count)
'    For i = 1 To col.Count
'        arr(i) = Mid$(texto, col(i)(0), col(i)(1) - col(i)(0) + 1)
'    Next i
'
'End Sub
'
''' ============================================================
'''   MÓDULO 1 — SEGMENTACIÓN BASE (SilabearAuto_Base)
''' ============================================================
''Private Sub SilabearAuto_Base(ByVal texto As String, ByRef arr() As String)
''
''    Dim col As New Collection
''    Dim i As Long, ini As Long
''    Dim c1 As String, c2 As String, c3 As String
''    Dim par As String
''
''    ini = 1
''
''    For i = 2 To Len(texto)
''
''        c1 = Mid$(texto, i - 1, 1)
''        c2 = Mid$(texto, i, 1)
''        If i < Len(texto) Then c3 = Mid$(texto, i + 1, 1) Else c3 = ""
''        par = c1 & c2
''
''        ' 0. ESPACIOS
''        If c1 = " " Then
''            If i - 2 >= ini Then col.Add Array(ini, i - 2)
''            ini = i
''            GoTo siguiente
''        End If
''
''        If c2 = " " Then
''            col.Add Array(ini, i - 1)
''            col.Add Array(i, i)
''            ini = i + 1
''            GoTo siguiente
''        End If
''
''        ' 1. DÍGRAFOS
''        If par = "ch" Or par = "ll" Or par = "rr" Then GoTo siguiente
''
''        ' 2. NSP / NST
''        If c1 = "n" And c2 = "s" Then
''            If c3 = "p" Or c3 = "t" Then
''                col.Add Array(ini, i)
''                ini = i + 1
''                GoTo siguiente
''            End If
''        End If
''
''' 3. GRUPOS INSEPARABLES (pero respetando diptongos previos)
''If EsConsonante_ES(c1) And EsConsonante_ES(c2) Then
''
''    ' Si es grupo inseparable…
''    If EsGrupoInseparable_ES(par) Then
''
''        ' …pero la vocal anterior forma diptongo con la vocal previa ? NO saltar
''        If i > 2 Then
''            Dim vPrev As String
''            vPrev = Mid$(texto, i - 2, 1)
''
''            ' Si vPrev + vocal actual forman diptongo ? NO saltar
''            If EsDiptongo(vPrev, Mid$(texto, i - 1, 1)) Then
''                ' No hacemos GoTo siguiente
''            Else
''                GoTo siguiente
''            End If
''        Else
''            GoTo siguiente
''        End If
''    End If
''End If
''
'''        ' 3. GRUPOS INSEPARABLES
'''        If EsConsonante_ES(c1) And EsConsonante_ES(c2) Then
'''            If EsGrupoInseparable_ES(par) Then GoTo siguiente
'''        End If
''
''        ' 4. CCC
''        If EsCCC(c1, c2, c3) Then
''            col.Add Array(ini, i - 1)
''            ini = i
''            GoTo siguiente
''        End If
''
''        ' 5. V + C + V TILDADA (ción)
''        If EsVocal_ES(c1) And EsConsonante_ES(c2) And EsVocal_ES(c3) Then
''            If c3 Like "[áéíóú]" Then GoTo siguiente
''        End If
''
''        ' 6. TRIPTONGO
''        If i < Len(texto) - 1 Then
''            If EsTriptongo(c1, c2, Mid$(texto, i + 1, 1)) Then GoTo siguiente
''        End If
''
''        ' 7. HIATOS
''        If EsVocal_ES(c1) And EsVocal_ES(c2) Then
''            If EsDiptongo(c1, c2) Then GoTo siguiente
''            If EsHiato(c1, c2) Then
''                col.Add Array(ini, i - 1)
''                ini = i
''                GoTo siguiente
''            End If
''            col.Add Array(ini, i - 1)
''            ini = i
''            GoTo siguiente
''        End If
''
''        ' 8. DIPTONGO
''        If EsDiptongo(c1, c2) Then GoTo siguiente
''
''        ' 9. V + RR
''        If EsVocal_ES(c1) And c2 = "r" Then
''            If i < Len(texto) Then
''                If Mid$(texto, i + 1, 1) = "r" Then
''                    col.Add Array(ini, i - 1)
''                    ini = i
''                    GoTo siguiente
''                End If
''            End If
''        End If
''
''        ' 10. ST después de NS
''        If c1 = "s" And c2 = "t" Then
''            If i > 2 Then
''                If Mid$(texto, i - 2, 1) = "n" Then GoTo siguiente
''            End If
''        End If
''
''' 11. CCV ? C | CV (solo si c1 y c2 son consonantes)
''If EsConsonante_ES(c1) And EsConsonante_ES(c2) Then
''    If i < Len(texto) Then
''        If EsVocal_ES(c3) Then
''
''            ' *** NUEVO: NO cortar si la vocal anterior forma diptongo ***
''            If i > 2 Then
''                Dim prevV As String
''                prevV = Mid$(texto, i - 2, 1)
''
''                ' Si prevV + c1 es diptongo ? NO cortar
''                If EsDiptongo(prevV, c1) Then GoTo siguiente
''            End If
''
''            If Not EsGrupoInseparable_ES(par) Then
''                col.Add Array(ini, i - 1)
''                ini = i
''                GoTo siguiente
''            End If
''        End If
''    End If
''End If
''
'''        ' 11. CCV ? C | CV (si NO es grupo inseparable)
'''        If EsConsonante_ES(c1) And EsConsonante_ES(c2) Then
'''            If i < Len(texto) Then
'''                If EsVocal_ES(c3) Then
'''                    If Not EsGrupoInseparable_ES(par) Then
'''                        col.Add Array(ini, i - 1)
'''                        ini = i
'''                        GoTo siguiente
'''                    End If
'''                End If
'''            End If
'''        End If
''
''        ' 12. VCV ? V | CV
''        If EsVocal_ES(c1) And EsConsonante_ES(c2) Then
''            If i < Len(texto) Then
''                If EsVocal_ES(c3) Then
''                    col.Add Array(ini, i - 1)
''                    ini = i
''                    GoTo siguiente
''                End If
''            End If
''        End If
''
''siguiente:
''    Next i
''
''    If ini <= Len(texto) Then col.Add Array(ini, Len(texto))
''
''    ReDim arr(1 To col.Count)
''    For i = 1 To col.Count
''        arr(i) = Mid$(texto, col(i)(0), col(i)(1) - col(i)(0) + 1)
''    Next i
''
''End Sub
'
'' ============================================================
''   MÓDULO 2 — CORRECCIONES INTERNAS (V X21)
''   - Hiatos acentuales
''   - Ataques complejos
''   - Diptongo creciente incompleto
'' ============================================================
'Private Sub SilabearAuto_Correcciones(ByRef arr() As String)
'
'    Dim arr2() As String
'    Dim arr3() As String
'    Dim arr4() As String
'
'    Dim i As Long, j As Long, k As Long, n As Long
'    Dim s As String, t As String, curr As String, prev As String
'    Dim v2 As String, vAbierta As String
'    Dim posV1 As Long, posCons As Long, posV2 As Long
'    Dim posV As Long, posC1 As Long, posC2 As Long
'    Dim parCC As String
'    Dim dipt As Variant, d As Variant
'    Dim tieneDiptongo As Boolean
'
'
'' ============================================================
''   1) CORRECCIÓN DE HIATOS ACENTUALES
'' ============================================================
'    ReDim arr2(1 To UBound(arr) * 2)
'    n = 1
'
'    For i = 1 To UBound(arr)
'        s = arr(i)
'
'        posV1 = 0: posCons = 0: posV2 = 0
'
'        ' Buscar patrón V C V´ dentro de la sílaba
'        For j = 1 To Len(s)
'            If EsVocal_ES(Mid$(s, j, 1)) Then
'                If posV1 = 0 Then
'                    posV1 = j
'                ElseIf posCons > 0 And posV2 = 0 Then
'                    posV2 = j
'                End If
'            ElseIf EsConsonante_ES(Mid$(s, j, 1)) Then
'                If posV1 > 0 And posCons = 0 Then
'                    posCons = j
'                End If
'            End If
'        Next j
'
'        ' Si hay patrón V C V´ y la segunda vocal está acentuada
'        If posV1 > 0 And posCons > 0 And posV2 > 0 Then
'            v2 = Mid$(s, posV2, 1)
'            If v2 Like "[áéíóú]" Then
'                arr2(n) = Left$(s, posV1)
'                n = n + 1
'                arr2(n) = Mid$(s, posCons)
'                n = n + 1
'                GoTo siguiente_hiato
'            End If
'        End If
'
'        arr2(n) = s
'        n = n + 1
'
'siguiente_hiato:
'    Next i
'
'    ReDim arr(1 To n - 1)
'    For i = 1 To n - 1
'        arr(i) = arr2(i)
'    Next i
'
'' --- CORRECCIÓN DE ATAQUES COMPLEJOS (sa | gra | do, pie | dra, a | grio...) ---
'ReDim arr3(1 To UBound(arr) * 2)
'n = 1
'
'For i = 1 To UBound(arr)
'    t = arr(i)
'
'    posV = 0: posC1 = 0: posC2 = 0: posV2 = 0
'
'    ' Buscar patrón V C C V dentro de la sílaba
'    For k = 1 To Len(t)
'        If EsVocal_ES(Mid$(t, k, 1)) Then
'            If posV = 0 Then
'                posV = k
'            ElseIf posC1 > 0 And posC2 > 0 And posV2 = 0 Then
'                posV2 = k
'            End If
'        ElseIf EsConsonante_ES(Mid$(t, k, 1)) Then
'            If posV > 0 And posC1 = 0 Then
'                posC1 = k
'            ElseIf posC1 > 0 And posC2 = 0 Then
'                posC2 = k
'            End If
'        End If
'    Next k
'
'    ' Si hay V C C V y CC es grupo inseparable, comprobar diptongo previo
'    If posV > 0 And posC1 > 0 And posC2 > 0 And posV2 > 0 Then
'
'        parCC = Mid$(t, posC1, 2)
'
'        If EsGrupoInseparable_ES(parCC) Then
'
'            ' --- NUEVO: comprobar si entre V y C1 hay una vocal que forme diptongo ---
'            Dim m As Long, vNucleo As String, rompeDiptongo As Boolean
'            rompeDiptongo = False
'
'            For m = posV + 1 To posC1 - 1
'                If EsVocal_ES(Mid$(t, m, 1)) Then
'                    vNucleo = Mid$(t, m, 1)
'                    If EsDiptongo(Mid$(t, posV, 1), vNucleo) Then
'                        rompeDiptongo = True
'                        Exit For
'                    End If
'                End If
'            Next m
'
'            ' Si hay diptongo ? NO cortar
'            If rompeDiptongo Then
'                arr3(n) = t
'                n = n + 1
'                GoTo siguienteT
'            End If
'
'            ' Si no hay diptongo ? cortar V | CCV
'            arr3(n) = Left$(t, posV)
'            n = n + 1
'            arr3(n) = Mid$(t, posC1)
'            n = n + 1
'            GoTo siguienteT
'        End If
'    End If
'
'    ' Si no hay división, copiar tal cual
'    arr3(n) = t
'    n = n + 1
'
'siguienteT:
'Next i
'
''' ============================================================
'''   2) CORRECCIÓN DE ATAQUES COMPLEJOS (V C C V ? V | CCV)
''' ============================================================
''    ReDim arr3(1 To UBound(arr) * 2)
''    n = 1
''
''    For i = 1 To UBound(arr)
''        t = arr(i)
''
''        posV = 0: posC1 = 0: posC2 = 0: posV2 = 0
''
''        ' Buscar patrón V C C V dentro de la sílaba
''        For k = 1 To Len(t)
''            If EsVocal_ES(Mid$(t, k, 1)) Then
''                If posV = 0 Then
''                    posV = k
''                ElseIf posC1 > 0 And posC2 > 0 And posV2 = 0 Then
''                    posV2 = k
''                End If
''            ElseIf EsConsonante_ES(Mid$(t, k, 1)) Then
''                If posV > 0 And posC1 = 0 Then
''                    posC1 = k
''                ElseIf posC1 > 0 And posC2 = 0 Then
''                    posC2 = k
''                End If
''            End If
''        Next k
''
''    ' Si hay V C C V y CC es grupo inseparable ? dividir V | CCV
''    If posV > 0 And posC1 > 0 And posC2 > 0 And posV2 > 0 Then
''        parCC = Mid$(t, posC1, 2)
''
''        ' --- NUEVO: comprobar si entre V y C1 hay una vocal que forme diptongo con V ---
''        Dim m As Long
''        Dim vNucleo As String
''        Dim rompeDiptongo As Boolean
''
''        rompeDiptongo = False
''
''        For m = posV + 1 To posC1 - 1
''            If EsVocal_ES(Mid$(t, m, 1)) Then
''                vNucleo = Mid$(t, m, 1)
''                If EsDiptongo(Mid$(t, posV, 1), vNucleo) Then
''                    rompeDiptongo = True
''                    Exit For
''                End If
''            End If
''        Next m
''
''        ' Si rompería un diptongo (como i + e en "piedra"), NO dividir
''        If rompeDiptongo Then
''            ' no hacemos nada, se copia tal cual más abajo
''        ElseIf EsGrupoInseparable_ES(parCC) Then
''            arr3(n) = Left$(t, posV)
''            n = n + 1
''            arr3(n) = Mid$(t, posC1)
''            n = n + 1
''            GoTo siguiente_ataque
''        End If
''    End If
''
'''        ' Si hay V C C V y CC es grupo inseparable ? dividir V | CCV
'''        If posV > 0 And posC1 > 0 And posC2 > 0 And posV2 > 0 Then
'''            parCC = Mid$(t, posC1, 2)
'''            If EsGrupoInseparable_ES(parCC) Then
'''                arr3(n) = Left$(t, posV)
'''                n = n + 1
'''                arr3(n) = Mid$(t, posC1)
'''                n = n + 1
'''                GoTo siguiente_ataque
'''            End If
'''        End If
''
''        arr3(n) = t
''        n = n + 1
''
''siguiente_ataque:
''    Next i
'
'    ReDim arr(1 To n - 1)
'    For i = 1 To n - 1
'        arr(i) = arr3(i)
'    Next i
'
'
'' ============================================================
''   3) CORRECCIÓN DE DIPTONGO CRECIENTE INCOMPLETO
'' ============================================================
'dipt = Array("ia", "ie", "io", "ua", "ue", "uo")
'
'ReDim arr4(1 To UBound(arr))
'n = 1
'
'For i = 1 To UBound(arr)
'    curr = arr(i)
'
'    If i > 1 Then
'
'        ' Detectar vocal abierta aunque vaya seguida de consonante
'        vAbierta = Left$(curr, 1)
'
'        If vAbierta Like "[aeo]" Then
'            prev = arr4(n - 1)
'
'            ' La sílaba anterior debe terminar en vocal cerrada
'            If Right$(prev, 1) Like "[iu]" Then
'
'                ' Comprobar si forman diptongo creciente válido
'                tieneDiptongo = False
'                For Each d In dipt
'                    If Right$(prev, 1) & vAbierta = d Then
'                        tieneDiptongo = True
'                        Exit For
'                    End If
'                Next d
'
'                If tieneDiptongo Then
'                    ' Fusionar la vocal abierta con la sílaba anterior
'                    arr4(n - 1) = prev & vAbierta
'                    curr = Mid$(curr, 2)
'                End If
'            End If
'        End If
'    End If
'
'    arr4(n) = curr
'    n = n + 1
'Next i
'
'ReDim arr(1 To n - 1)
'For i = 1 To n - 1
'    arr(i) = arr4(i)
'Next i
'
''' ============================================================
'''   3) CORRECCIÓN DE DIPTONGO CRECIENTE INCOMPLETO
'''      (piedra ? pie | dra, viento ? vien | to)
''' ============================================================
''    dipt = Array("ia", "ie", "io", "ua", "ue", "uo")
''
''    ReDim arr4(1 To UBound(arr))
''    n = 1
''
''    For i = 1 To UBound(arr)
''        curr = arr(i)
''
''        If i > 1 Then
''            vAbierta = Left$(curr, 1)
''
''            If vAbierta Like "[aeo]" Then
''                prev = arr4(n - 1)
''
''                If Right$(prev, 1) Like "[iu]" Then
''
''                    tieneDiptongo = False
''                    For Each d In dipt
''                        If Right$(prev, 1) & vAbierta = d Then
''                            tieneDiptongo = True
''                            Exit For
''                        End If
''                    Next d
''
''                    If tieneDiptongo Then
''                        arr4(n - 1) = prev & vAbierta
''                        curr = Mid$(curr, 2)
''                    End If
''                End If
''            End If
''        End If
''
''        arr4(n) = curr
''        n = n + 1
''    Next i
''
''    ReDim arr(1 To n - 1)
''    For i = 1 To n - 1
''        arr(i) = arr4(i)
''    Next i
'
'End Sub
'
'' ============================================================
''   MÓDULO 3 — ENSAMBLADO FINAL (V X21)
''   - Fusión de "s" suelta
''   - Construcción de la cadena final
'' ============================================================
'Private Sub SilabearAuto_EnsambladoFinal(ByRef arr() As String, ByRef salidaFinal As String)
'
'    Dim salida() As String
'    Dim i As Long, n As Long
'    Dim tx As String
'
'    ' --- FUSIÓN DE S SUELTA ---
'    ReDim salida(UBound(arr))
'    n = 0
'
'    For i = 1 To UBound(arr)
'
'        If arr(i) = "s" Then
'            ' Fusionar con la sílaba anterior
'            If n > 0 Then
'                salida(n - 1) = salida(n - 1) & "s"
'                n = n - 1
'            Else
'                salida(n) = "s"
'            End If
'
'        Else
'            salida(n) = arr(i)
'        End If
'
'        n = n + 1
'    Next i
'
'    ' --- CONSTRUCCIÓN DE LA CADENA FINAL ---
'    tx = ""
'    For i = 0 To n - 1
'        If Trim$(salida(i)) <> "" Then
'            tx = tx & salida(i) & " | "
'        End If
'    Next i
'
'    ' Quitar la última barra
'    tx = Trim$(tx)
'    If Right$(tx, 1) = "|" Then tx = Trim$(Left$(tx, Len(tx) - 1))
'
'    salidaFinal = tx
'
'End Sub
'
'
''' ============================================================
'''   2- SILABEO AUTOMÁTICO (VERSIÓN MODULAR) V X20
''' ============================================================
''Private Sub SilabearAuto()
''
''    Dim texto As String
''    Dim col As New Collection
''    Dim i As Long, ini As Long
''    Dim c1 As String, c2 As String, c3 As String
''    Dim par As String
''    Dim arr() As String
''
'''    Debug.Print
'''    Debug.Print ">>> SilabearAuto Iniciando V10"
''
''    texto = Trim$(objDTO.TextoNormalizado)
''    If Len(texto) = 0 Then
''        objDTO.SilabasAuto = ""
''        Exit Sub
''    End If
''
''    ini = 1
''
''    For i = 2 To Len(texto)
''
''        c1 = Mid$(texto, i - 1, 1)
''        c2 = Mid$(texto, i, 1)
''        If i < Len(texto) Then c3 = Mid$(texto, i + 1, 1) Else c3 = ""
''        par = c1 & c2
''
''        ' --------------------------------------------------------
''        ' 0. ESPACIOS
''        ' --------------------------------------------------------
''        If c1 = " " Then
''            If i - 2 >= ini Then col.Add Array(ini, i - 2)
''            ini = i
''            GoTo siguiente
''        End If
''
''        If c2 = " " Then
''            col.Add Array(ini, i - 1)
''            col.Add Array(i, i)
''            ini = i + 1
''            GoTo siguiente
''        End If
''
''        ' --------------------------------------------------------
''        ' 1. DÍGRAFOS
''        ' --------------------------------------------------------
''        If par = "ch" Or par = "ll" Or par = "rr" Then GoTo siguiente
''
''        ' --------------------------------------------------------
''        ' 2. NSP / NST
''        ' --------------------------------------------------------
''        If c1 = "n" And c2 = "s" Then
''            If c3 = "p" Or c3 = "t" Then
''                col.Add Array(ini, i)
''                ini = i + 1
''                GoTo siguiente
''            End If
''        End If
''
''        ' --------------------------------------------------------
''        ' 3. GRUPOS INSEPARABLES
''        ' --------------------------------------------------------
''        If EsConsonante_ES(c1) And EsConsonante_ES(c2) Then
''            If EsGrupoInseparable_ES(par) Then GoTo siguiente
''        End If
''
''        ' --------------------------------------------------------
''        ' 4. CCC
''        ' --------------------------------------------------------
''        If EsCCC(c1, c2, c3) Then
''            col.Add Array(ini, i - 1)
''            ini = i
''            GoTo siguiente
''        End If
''
''        ' --------------------------------------------------------
''        ' 5. EXCEPCIÓN: V + C + V TILDADA (ción)
''        ' --------------------------------------------------------
''        If EsVocal_ES(c1) And EsConsonante_ES(c2) And EsVocal_ES(c3) Then
''            If c3 Like "[áéíóú]" Then GoTo siguiente
''        End If
''
''        ' --------------------------------------------------------
''        ' 6. TRIPTONGO
''        ' --------------------------------------------------------
''        If i < Len(texto) - 1 Then
''            If EsTriptongo(c1, c2, Mid$(texto, i + 1, 1)) Then GoTo siguiente
''        End If
''
''        ' --------------------------------------------------------
''        ' 7. HIATOS (nuevo bloque)
''        ' --------------------------------------------------------
''        If EsVocal_ES(c1) And EsVocal_ES(c2) Then
''
''            ' No separar si es diptongo
''            If EsDiptongo(c1, c2) Then GoTo siguiente
''
''            ' Separar si es hiato
''            If EsHiato(c1, c2) Then
''                col.Add Array(ini, i - 1)
''                ini = i
''                GoTo siguiente
''            End If
''
''            ' Separación por defecto
''            col.Add Array(ini, i - 1)
''            ini = i
''            GoTo siguiente
''        End If
''
''        ' --------------------------------------------------------
''        ' 8. DIPTONGO
''        ' --------------------------------------------------------
''        If EsDiptongo(c1, c2) Then GoTo siguiente
''
''        ' --------------------------------------------------------
''        ' 9. V + RR
''        ' --------------------------------------------------------
''        If EsVocal_ES(c1) And c2 = "r" Then
''            If i < Len(texto) Then
''                If Mid$(texto, i + 1, 1) = "r" Then
''                    col.Add Array(ini, i - 1)
''                    ini = i
''                    GoTo siguiente
''                End If
''            End If
''        End If
''
''        ' --------------------------------------------------------
''        ' 10. ST después de NS
''        ' --------------------------------------------------------
''        If c1 = "s" And c2 = "t" Then
''            If i > 2 Then
''                If Mid$(texto, i - 2, 1) = "n" Then GoTo siguiente
''            End If
''        End If
''
''        ' --------------------------------------------------------
''        ' 11. CCV ? C | CV
''        ' --------------------------------------------------------
''        If EsConsonante_ES(c1) And EsConsonante_ES(c2) Then
''            If i < Len(texto) Then
''                If EsVocal_ES(c3) Then
''                    If Not EsGrupoInseparable_ES(par) Then
''
''                        ' <<< PARCHE DE SEGURIDAD >>>
''                        If i - 1 >= ini Then
''                            col.Add Array(ini, i - 1)
''                        End If
''
''                        ini = i
''                        GoTo siguiente
''                    End If
''                End If
''            End If
''        End If
''
'''        If EsConsonante_ES(c1) And EsConsonante_ES(c2) Then
'''            If i < Len(texto) Then
'''                If EsVocal_ES(c3) Then
'''                    If Not EsGrupoInseparable_ES(par) Then
'''                        col.Add Array(ini, i - 1)
'''                        ini = i
'''                        GoTo siguiente
'''                    End If
'''                End If
'''            End If
'''        End If
''
''        ' --------------------------------------------------------
''        ' 12. VCV ? V | CV
''        ' --------------------------------------------------------
''        If EsVocal_ES(c1) And EsConsonante_ES(c2) Then
''            If i < Len(texto) Then
''                If EsVocal_ES(c3) Then
''                    col.Add Array(ini, i - 1)
''                    ini = i
''                    GoTo siguiente
''                End If
''            End If
''        End If
''
''siguiente:
''    Next i
''
''    If ini <= Len(texto) Then col.Add Array(ini, Len(texto))
''
'''    ' >>> AQUI, JUSTO AQUI <<<
'''    Debug.Print "---- Detalle índices col ----"
'''    For i = 1 To col.Count
'''        Debug.Print i, "ini:", col(i)(0), "fin:", col(i)(1), _
'''                     "texto:", Mid$(texto, col(i)(0), col(i)(1) - col(i)(0) + 1)
'''    Next i
'''    Debug.Print "-----------------------------"
''
''    ReDim arr(1 To col.Count)
''    For i = 1 To col.Count
''        arr(i) = Mid$(texto, col(i)(0), col(i)(1) - col(i)(0) + 1)
''    Next i
''
''' --- CORRECCIÓN DE HIATOS ACENTUALES (María, Lucía, frío, río...) ---
''Dim arr2() As String
''Dim s As String, j As Long
''Dim v2 As String
''Dim posV1 As Long, posCons As Long, posV2 As Long
''Dim n
''
''Dim arr3() As String
''Dim t As String
''Dim k As Long
''Dim posV As Long, posC1 As Long, posC2 As Long
''Dim parCC As String
''
''' IMPORTANTE: reservar espacio suficiente (el doble)
''ReDim arr2(1 To UBound(arr) * 2)
''n = 1
''
''For i = 1 To UBound(arr)
''    s = arr(i)
''
''    posV1 = 0: posCons = 0: posV2 = 0
''
''    ' Buscar patrón V C V´ dentro de la sílaba
''    For j = 1 To Len(s)
''        If EsVocal_ES(Mid$(s, j, 1)) Then
''            If posV1 = 0 Then
''                posV1 = j
''            ElseIf posCons > 0 And posV2 = 0 Then
''                posV2 = j
''            End If
''        ElseIf EsConsonante_ES(Mid$(s, j, 1)) Then
''            If posV1 > 0 And posCons = 0 Then
''                posCons = j
''            End If
''        End If
''    Next j
''
''    ' Si hay patrón V C V´ y la segunda vocal está acentuada
''    If posV1 > 0 And posCons > 0 And posV2 > 0 Then
''        v2 = Mid$(s, posV2, 1)
''        If v2 Like "[áéíóú]" Then
''            ' Dividir: V | C V´
''            arr2(n) = Left$(s, posV1)
''            n = n + 1
''            arr2(n) = Mid$(s, posCons)
''            n = n + 1
''            GoTo siguienteS
''        End If
''    End If
''
''    ' Si no hay división, copiar tal cual
''    arr2(n) = s
''    n = n + 1
''
''siguienteS:
''Next i
''
''' Redimensionar arr con el nuevo contenido
''ReDim arr(1 To n - 1)
''For i = 1 To n - 1
''    arr(i) = arr2(i)
''Next i
''
''' --- CORRECCIÓN DE ATAQUES COMPLEJOS (sa | gra | do, pie | dra, a | grio...) ---
''ReDim arr3(1 To UBound(arr) * 2)
''n = 1
''
''For i = 1 To UBound(arr)
''    t = arr(i)
''
''    posV = 0: posC1 = 0: posC2 = 0: posV2 = 0
''
''    ' Buscar patrón V C C V dentro de la sílaba
''    For k = 1 To Len(t)
''        If EsVocal_ES(Mid$(t, k, 1)) Then
''            If posV = 0 Then
''                posV = k
''            ElseIf posC1 > 0 And posC2 > 0 And posV2 = 0 Then
''                posV2 = k
''            End If
''        ElseIf EsConsonante_ES(Mid$(t, k, 1)) Then
''            If posV > 0 And posC1 = 0 Then
''                posC1 = k
''            ElseIf posC1 > 0 And posC2 = 0 Then
''                posC2 = k
''            End If
''        End If
''    Next k
''
''    ' Si hay V C C V y CC es grupo inseparable ? dividir V | CCV
''    If posV > 0 And posC1 > 0 And posC2 > 0 And posV2 > 0 Then
''        parCC = Mid$(t, posC1, 2)
''        If EsGrupoInseparable_ES(parCC) Then
''            arr3(n) = Left$(t, posV)
''            n = n + 1
''            arr3(n) = Mid$(t, posC1)
''            n = n + 1
''            GoTo siguienteT
''        End If
''    End If
''
''    ' Si no hay división, copiar tal cual
''    arr3(n) = t
''    n = n + 1
''
''siguienteT:
''Next i
''
''' Redimensionar arr con el nuevo contenido
''ReDim arr(1 To n - 1)
''For i = 1 To n - 1
''    arr(i) = arr3(i)
''Next i
''
''
''' --- FUSIÓN DE S SUELTA ---
''Dim salida() As String
''Dim tx As String
''
''
''ReDim salida(UBound(arr))
''n = 0
''
''For i = 1 To UBound(arr)
''
''    If arr(i) = "s" Then
''        ' Fusionar con el anterior (si existe)
''        If n > 0 Then
''            salida(n - 1) = salida(n - 1) & "s"
''            n = n - 1
''        Else
''            ' Caso imposible en español, pero por seguridad
''            salida(n) = "s"
''        End If
''
''    Else
''        salida(n) = arr(i)
''    End If
''
''    n = n + 1
''Next i
''
''' Construir la salida final
''tx = ""
''For i = 0 To n - 1
''    If Trim$(salida(i)) <> "" Then
''        tx = tx & salida(i) & " | "
''    End If
''Next i
''
''' Quitar la última barra y espacios
''tx = Trim$(tx)
''If Right$(tx, 1) = "|" Then tx = Trim$(Left$(tx, Len(tx) - 1))
''
''
''    objDTO.SilabasAuto = tx ' Join(arr, " | ")
''
'''    Debug.Print ">>> SilabearAuto ejecutado"
''
''End Sub
'
'Private Function EsNSP(c1 As String, c2 As String, c3 As String) As Boolean
'    EsNSP = (c1 = "n" And c2 = "s" And c3 = "p")
'End Function
'
'Private Function EsNST(c1 As String, c2 As String, c3 As String) As Boolean
'    EsNST = (c1 = "n" And c2 = "s" And c3 = "t")
'End Function
'
'Private Function EsCCC(c1 As String, c2 As String, c3 As String) As Boolean
'
'    If Not (EsConsonante_ES(c1) And EsConsonante_ES(c2) And EsConsonante_ES(c3)) Then
'        EsCCC = False
'        Exit Function
'    End If
'
'    ' No tocar NSP/NST
'    If c1 = "n" And c2 = "s" Then
'        EsCCC = False
'        Exit Function
'    End If
'
'    ' c2c3 inseparable ? C | CC
'    If EsGrupoInseparable_ES(c2 & c3) Then
'        EsCCC = True
'        Exit Function
'    End If
'
'    ' c1c2 inseparable ? CC | C
'    If EsGrupoInseparable_ES(c1 & c2) Then
'        EsCCC = True
'        Exit Function
'    End If
'
'    ' General ? C | CC
'    EsCCC = True
'
'End Function
'
'' ============================================================
''   3- DETECCIÓN DE ACENTUACIÓN
'' ============================================================
'' ------------------------------------------------------------
''   3.1- DETECCIÓN DE TÓNICA
'' ------------------------------------------------------------
'Private Sub DetectarTonica()
'
'    Dim partes() As String
'    Dim palabras As New Collection
'    Dim palabraActual As New Collection
'    Dim i As Long, j As Long
'    Dim s As String
'    Dim idx As Long
'    Dim offset As Long
'    Dim arrTonica() As String
'    Dim countTonica As Long
'
'    ' Si no hay sílabas, no hay nada que hacer
'    If Len(objDTO.SilabasAuto) = 0 Then
'        objDTO.SilabaTonica = ""
'        Exit Sub
'    End If
'
'    ' 1. Dividir sílabas globales
'    partes = Split(objDTO.SilabasAuto, " | ")
'
'    ' 2. Reconstruir palabras a partir de la sílaba vacía " "
'    For i = LBound(partes) To UBound(partes)
'        s = partes(i)
'
'        If Trim$(s) = "" Then
'            ' Fin de palabra
'            If palabraActual.Count > 0 Then
'                palabras.Add palabraActual
'                Set palabraActual = New Collection
'            End If
'        Else
'            palabraActual.Add s
'        End If
'    Next i
'
'    ' Añadir la última palabra si existe
'    If palabraActual.Count > 0 Then palabras.Add palabraActual
'
'    ' 3. Preparar array dinámico para tónicas
'    ReDim arrTonica(0 To 0)
'    countTonica = 0
'    offset = 0
'
'    ' 4. Procesar palabra por palabra
'    For i = 1 To palabras.Count
'
'        ' Detectar tónica de esta palabra (tu función existente)
'        idx = DetectarTonicaDeUnaPalabra_DesdeSilabas(palabras(i))
'
'        If idx > 0 Then
'            countTonica = countTonica + 1
'            ReDim Preserve arrTonica(0 To countTonica - 1)
'            arrTonica(countTonica - 1) = CStr(offset + idx)
'        End If
'
'        ' Avanzar offset global (+1 por la sílaba vacía entre palabras)
'        offset = offset + palabras(i).Count + 1
'    Next i
'
'    ' 5. Guardar resultado final
'    If countTonica = 0 Then
'        objDTO.SilabaTonica = ""
'    Else
'        objDTO.SilabaTonica = Join(arrTonica, ",")
'    End If
'
'End Sub
'
'' ------------------------------------------------------------
''   3.1.1- DETECTAR SÍLABA TÓNICA
'' ------------------------------------------------------------
'Private Function DetectarTonicaDeUnaPalabra_DesdeSilabas(ByVal colSilabas As Collection) As Long
'    Dim i As Long
'    Dim sil As String
'
'    ' 1. Buscar tilde explícita
'    For i = 1 To colSilabas.Count
'        sil = colSilabas(i)
'        If TieneTilde(sil) Then
'            DetectarTonicaDeUnaPalabra_DesdeSilabas = i
'            Exit Function
'        End If
'    Next i
'
'    ' 2. Si no hay tilde, aplicar reglas generales:
'    '    - aguda si termina en vocal, n o s
'    '    - llana en caso contrario
'
'    Dim ultima As String
'    ultima = colSilabas(colSilabas.Count)
'
'    If TerminaEnVocalNoSNoN(ultima) Then
'        ' Palabra llana ? tónica en penúltima
'        If colSilabas.Count >= 2 Then
'            DetectarTonicaDeUnaPalabra_DesdeSilabas = colSilabas.Count - 1
'        Else
'            DetectarTonicaDeUnaPalabra_DesdeSilabas = colSilabas.Count
'        End If
'    Else
'        ' Palabra aguda ? tónica en última
'        DetectarTonicaDeUnaPalabra_DesdeSilabas = colSilabas.Count
'    End If
'End Function
'
'' ------------------------------------------------------------
''   3.1.1.1- DETECTAR TILDE
'' ------------------------------------------------------------
'Private Function TieneTilde(ByVal silaba As String) As Boolean
'    TieneTilde = (InStr(silaba, "á") > 0 Or _
'                  InStr(silaba, "é") > 0 Or _
'                  InStr(silaba, "í") > 0 Or _
'                  InStr(silaba, "ó") > 0 Or _
'                  InStr(silaba, "ú") > 0)
'End Function
'
'' ------------------------------------------------------------
''   3.1.2- MARCAR TÓNICA
'' ------------------------------------------------------------
'Private Sub MarcarSilabasTonicas()
'
'Dim marcadas As String
'
'    marcadas = MarcarTonicas(objDTO.SilabasAuto, objDTO.SilabaTonica)
'    objDTO.SilabasAuto = marcadas
'
'End Sub
'
'' ------------------------------------------------------------
''   3.2- DETECCIÓN DE SECUNDARIA
'' ------------------------------------------------------------
'Public Sub DetectarSecundarias()
'
'    Dim partes() As String
'    Dim palabras As New Collection
'    Dim palabraActual As New Collection
'    Dim i As Long
'    Dim s As String
'    Dim offset As Long
'    Dim resultado As String
'    Dim idxTonicaLocal As Long
'    Dim idxSecRel As String
'    Dim partesSec() As String
'    Dim p As Variant
'    Dim idxGlobal As Long
'
'    resultado = ""
'
'    ' ------------------------------------------------------------
'    ' 1) Si no hay sílabas, no hay nada que hacer
'    ' ------------------------------------------------------------
'    If Len(objDTO.SilabasAuto) = 0 Then
'        objDTO.SilabaSecundaria = ""
'        Exit Sub
'    End If
'
'    ' ------------------------------------------------------------
'    ' 2) Dividir sílabas globales
'    ' ------------------------------------------------------------
'    partes = Split(objDTO.SilabasAuto, " | ")
'
'    ' ------------------------------------------------------------
'    ' 3) Reconstruir palabras a partir de la sílaba vacía " "
'    ' ------------------------------------------------------------
'    For i = LBound(partes) To UBound(partes)
'        s = partes(i)
'
'        If Trim$(s) = "" Then
'            ' Fin de palabra
'            If palabraActual.Count > 0 Then
'                palabras.Add palabraActual
'                Set palabraActual = New Collection
'            End If
'        Else
'            palabraActual.Add s
'        End If
'    Next i
'
'    ' Añadir la última palabra si existe
'    If palabraActual.Count > 0 Then palabras.Add palabraActual
'
'    ' ------------------------------------------------------------
'    ' 4) Procesar palabra por palabra
'    ' ------------------------------------------------------------
'    offset = 0
'
'    For i = 1 To palabras.Count
'
'        ' Obtener tónica local (ya es global en DetectarTonica)
'        idxTonicaLocal = DetectarTonicaDeUnaPalabra_DesdeSilabas(palabras(i))
'
'        ' --------------------------------------------------------
'        ' 4.1 Detectar secundarias relativas de esta palabra
'        ' --------------------------------------------------------
'        idxSecRel = DetectarSecundariasDeUnaPalabra(palabras(i), idxTonicaLocal)
'
'        If idxSecRel <> "" Then
'
'            partesSec = Split(idxSecRel, ",")
'
'            ' ----------------------------------------------------
'            ' 4.2 Convertir cada índice relativo ? índice global
'            ' ----------------------------------------------------
'            For Each p In partesSec
'
'                idxGlobal = offset + CLng(p)
'
'                If resultado = "" Then
'                    resultado = CStr(idxGlobal)
'                Else
'                    resultado = resultado & "," & CStr(idxGlobal)
'                End If
'
'            Next p
'
'        End If
'
'        ' --------------------------------------------------------
'        ' 4.3 Avanzar offset global (+1 por la sílaba vacía)
'        ' --------------------------------------------------------
'        offset = offset + palabras(i).Count + 1
'
'    Next i
'
'    ' ------------------------------------------------------------
'    ' 5) Guardar resultado final
'    ' ------------------------------------------------------------
'    objDTO.SilabaSecundaria = resultado
'
'End Sub
'
'' ------------------------------------------------------------
''   3.2.1- DETECTAR SECUNDARIAS DE PALABRA
'' ------------------------------------------------------------
'' ============================================================
''   DetectarSecundariaDeUnaPalabra
''
''   Entrada:
''       colSilabas  -> Collection con las sílabas de UNA palabra
''       idxTonica   -> Índice (1-based) de la sílaba tónica
''
''   Salida:
''       Índice (1-based) de la sílaba con acento secundario
''       o 0 si no existe
''
''   Reglas implementadas (versión 2.3):
''
''             - patrones prosódicos más finos
''
''       1) Solo palabras de 4 sílabas o más pueden tener secundaria.
''
''       2) secundaria inicial (solo palabras largas)
''
''       3) La secundaria suele caer 2 sílabas antes de la tónica si es sílaba fuerte.
''
''       3) Si no es válida, se prueba 3 sílabas antes de la tónica si es fuerte
''
''       4) Nunca puede coincidir con la tónica ni con la sílaba
''          inmediatamente posterior.
''
''
''
''       5) Si ninguna opción es válida, devuelve 0.
''
''   NOTA:
''       Esta es la tercera versión.
'
'' ============================================================
'Private Function DetectarSecundariasDeUnaPalabra( _
'        colSilabas As Collection, _
'        idxTonica As Long) As String
'
'    Dim numSilabas As Long
'    Dim cand2 As Long, cand3 As Long
'    Dim resultado As String
'    Dim fuerte2 As Boolean, fuerte3 As Boolean
'
'    numSilabas = colSilabas.Count
'    resultado = ""
'
'    ' 1) Secundaria inicial (palabras largas)
'    If numSilabas >= 6 Then
'        resultado = "1"
'    End If
'
'    ' 2) Validación de tónica
'    If idxTonica < 1 Or idxTonica > numSilabas Then
'        DetectarSecundariasDeUnaPalabra = resultado
'        Exit Function
'    End If
'
'    ' 3) Candidatos intermedios
'    cand2 = idxTonica - 2
'    cand3 = idxTonica - 3
'
'    If cand2 >= 1 Then fuerte2 = EsSilabaFuerte(colSilabas(cand2))
'    If cand3 >= 1 Then fuerte3 = EsSilabaFuerte(colSilabas(cand3))
'
'    ' 4) Selección prosódica real
'    Dim elegido As Long: elegido = 0
'
'    If cand2 >= 1 And fuerte2 Then
'        elegido = cand2
'    ElseIf cand3 >= 1 And fuerte3 Then
'        elegido = cand3
'    ElseIf cand2 >= 1 Then
'        elegido = cand2
'    End If
'
'    ' 5) Añadir secundaria intermedia
'    If elegido > 0 Then
'        If resultado = "" Then
'            resultado = CStr(elegido)
'        Else
'            resultado = resultado & "," & CStr(elegido)
'        End If
'    End If
'
'    DetectarSecundariasDeUnaPalabra = resultado
'End Function
'
'Private Function EsSilabaFuerte(silaba As String) As Boolean
'    Dim v As String
'    v = LCase$(silaba)
'
'    ' Vocales fuertes
'    If InStr(v, "a") > 0 Or InStr(v, "e") > 0 Or InStr(v, "o") > 0 Then
'        EsSilabaFuerte = True
'        Exit Function
'    End If
'
'    ' Diptongos con vocal fuerte dominante
'    If InStr(v, "ai") > 0 Or InStr(v, "au") > 0 _
'    Or InStr(v, "ei") > 0 Or InStr(v, "eu") > 0 _
'    Or InStr(v, "oi") > 0 Or InStr(v, "ou") > 0 Then
'        EsSilabaFuerte = True
'        Exit Function
'    End If
'
'    ' Sílaba débil típica
'    Select Case v
'        Case "de", "le", "lo", "me", "te", "se", "que"
'            EsSilabaFuerte = False
'            Exit Function
'    End Select
'
'    ' Vocal cerrada sola ? débil
'    If v = "i" Or v = "u" Then
'        EsSilabaFuerte = False
'        Exit Function
'    End If
'
'    ' Por defecto, consideramos fuerte
'    EsSilabaFuerte = True
'End Function
'
'' ------------------------------------------------------------
''   3.2.2- MARCAR SECUNDARIAS
'' ------------------------------------------------------------
'Private Sub MarcarSilabasSecundarias()
'
'Dim marcadas As String
'
'    marcadas = MarcarSecundarias(objDTO.SilabasAuto, objDTO.SilabaSecundaria)
'    objDTO.SilabasAuto = marcadas
'
'End Sub
'
'' ============================================================
''   4- REVISIÓN MANUAL
'' ============================================================
'Private Sub RevisionSilabeo()
'
'    Dim texto As String
'    Dim s As String
'    Dim partes() As String
'    Dim resultado As String
'    Dim idxTonica As New Collection
'    Dim raw As String
'    Dim i As Long
'
'    'texto = objDTO.TextoNormalizado
'    texto = objDTO.TextoOriginal
'
'    ' SilabasAuto ya es un string con guiones
'    s = objDTO.SilabasAuto
'
'    ' Abrir formulario de revisión
'    s = RevisarSilabas_EnFormulario(texto, s)
'
'    ' Si el usuario cancela, conservar lo anterior
'    If s = "" Then
'        objDTO.SilabasFinal = objDTO.SilabasAuto
'        objDTO.SilabaTonica = objDTO.SilabaTonica
'        Exit Sub
'    End If
'
'    ' Dividir sílabas revisadas
'    partes = Split(s, " | ")
'
'    ' Procesar cada sílaba
'    For i = LBound(partes) To UBound(partes)
'
'        raw = partes(i)
'
'        If EsSilabaMarcada(raw) Then
'            idxTonica.Add i + 1
'            partes(i) = LimpiarSilabaMarcada(raw)
'        Else
'            partes(i) = raw
'        End If
'
'    Next i
'
'    ' Reconstruir SilabasFinal como string
'    resultado = Join(partes, " | ")
'    objDTO.SilabasFinal = resultado
'
'    ' Reconstruir SilabaTonica como string
'    If idxTonica.Count > 0 Then
'        Dim arr() As String
'        ReDim arr(0 To idxTonica.Count - 1)
'
'        For i = 1 To idxTonica.Count
'            arr(i - 1) = CStr(idxTonica(i))
'        Next i
'
'        objDTO.SilabaTonica = Join(arr, ",")
'    End If
'
'End Sub
'
'
''======================================================================================
'' Conversor Fonético
''======================================================================================
'
'' ============================================================
''   PROCEDIMIENTO: ConvertirSilabasAFonemas
''   Ensambla todas las sílabas en una secuencia de fonemas
''   Entrada: array de sílabas (strings)
''   Salida: string con IDs separados por comas
'' ============================================================
'Public Sub ConvertirSilabasAFonemas()
'
'    Dim listaSilabas() As String
'    Dim listaTonicas() As String
'    Dim idx As Byte
'    Dim salida As String
'    Dim bloque As String
'    Dim esTonica As Boolean
'
'    Dim idxTonica As Byte
'
'    salida = ""
'    idxTonica = 0
'
'    ' 1. Separar sílabas de toda la frase
'    listaSilabas = Split(objDTO.SilabasFinal, " | ")
'
'    ' 2. Separar índices de sílabas tónicas (globales)
'    listaTonicas = Split(objDTO.SilabaTonica, ",")
'
'    ' 3. Procesar cada sílaba
'    For idx = 0 To UBound(listaSilabas)
'
'        ' ¿Es esta sílaba tónica?
'        esTonica = False 'EsIndiceTonica(idx + 1, listaTonicas)
'        If idxTonica <= UBound(listaTonicas) Then
'            If (idx + 1) = CByte(listaTonicas(idxTonica)) Then
'                esTonica = True
'                idxTonica = idxTonica + 1
'            End If
'        End If
'
'        ' ¿Es un espacio?
'        If Trim(listaSilabas(idx)) = "" Then
'            bloque = "0"
'        Else
'            Debug.Print Chr(34); listaSilabas(idx); Chr(34)
'            ' Convertir sílaba a fonemas
'            bloque = Conv_Silaba(listaSilabas(idx))
'
'            Debug.Print bloque
'            ' Insertar acento si corresponde
'            If esTonica Then
'                bloque = "61," & bloque
'            End If
'        End If
'
'        ' Añadir al resultado final
'        If salida = "" Then
'            salida = bloque
'        Else
'            salida = salida & " | " & bloque
'        End If
'
'    Next idx
'
'    objDTO.TextoFinal = salida
'
'End Sub
'
'
''======================================================================================
'' Conversor
''======================================================================================
'' ============================================================
''   MÓDULO PRINCIPAL: Conv_Silaba
''   Orquesta la conversión grafema ? fonema
''   Devuelve:
''       - Collection de fonemas
''       - String de fonemas
'' ============================================================
'
'Private Function Conv_Silaba(silaba As String) As String
'
'    'Dim col As New Collection
'    Dim i As Long
'    Dim g As String, g2 As String
'    Dim anterior As String, siguiente As String
'    Dim resultado As String
'    Dim f As Variant
'
'    Dim asTemp As String
'
'    asTemp = ""
'    silaba = NormalizaVocales(LCase$(silaba))
'
'    For i = 1 To Len(silaba)
'
'        g = Mid$(silaba, i, 1)
'
'        anterior = ""
'        If i > 1 Then anterior = Mid$(silaba, i - 1, 1)
'
'        siguiente = ""
'        If i < Len(silaba) Then siguiente = Mid$(silaba, i + 1, 1)
'
'        ' ============================================
'        '   1. DÍGRAFOS (2 grafemas ? 1 fonema)
'        ' ============================================
'        If i < Len(silaba) Then
'            g2 = Mid$(silaba, i, 2)
'            resultado = Conv_Digrafos(g2, siguiente)
'
'            If resultado <> "" Then
'                asTemp = asTemp & "  |  " & resultado
'                'col.Add CLng(resultado)
'                i = i + 1   ' Consumimos 2 grafemas
'                GoTo SiguienteGrafema
'            End If
'        End If
'
'        ' ============================================
'        '   2. VOCALES / DIPTONGOS / HIATOS
'        ' ============================================
'        resultado = Conv_Vocales(g, siguiente)
'
'        If resultado <> "" Then
'            asTemp = asTemp & ", " & resultado
'            'If InStr(resultado, ",") > 0 Then
'                'For Each f In Split(resultado, ",")
'                '    col.Add CLng(f)
'                'Next f
'            'Else
'            '    col.Add CLng(resultado)
'            'End If
'            GoTo SiguienteGrafema
'        End If
'
'        ' ============================================
'        '   3. REGLAS CONTEXTUALES
'        ' ============================================
'        resultado = Conv_Contexto(g, anterior, siguiente)
'
'        If resultado <> "" Then
'            asTemp = asTemp & ", " & resultado
'            'col.Add CLng(resultado)
'            GoTo SiguienteGrafema
'        End If
'
'        ' ============================================
'        '   4. MONÓGRAFOS
'        ' ============================================
'        resultado = Conv_Monografos(g)
'
'        If resultado <> "" Then
'            asTemp = asTemp & ", " & resultado
'            'col.Add CLng(resultado)
'            GoTo SiguienteGrafema
'        End If
'
'        ' ============================================
'        '   5. Si nada aplica ? placeholder
'        ' ============================================
'        'col.Add 99
'        asTemp = asTemp & ", " & resultado
'SiguienteGrafema:
'    Next i
'
'    Conv_Silaba = Trim(Mid(asTemp, 2))
'    'Set Conv_Silaba = col
'
'End Function
'
'' ============================================================
''   PROCEDIMIENTO: Conv_Digrafos
''   Convierte dígrafos (2 grafemas ? 1 fonema)
''   Devuelve:
''       - ID fonema (Long)
''       - Null si no aplica
'' ============================================================
'Private Function Conv_Digrafos(g2 As String, siguiente As String) As String
'
'    g2 = LCase$(g2)
'    siguiente = LCase$(siguiente)
'
'    Select Case g2
'
'        ' ===== Dígrafos reales =====
'
'        Case "ch"
'            Conv_Digrafos = "41"      ' /t?/
'
'        Case "ll"
'            Conv_Digrafos = "42"      ' /?/ o /?/
'
'        Case "rr"
'            Conv_Digrafos = "43"      ' vibrante múltiple
'
'
'        ' ===== Dígrafos ortográficos =====
'
'        Case "gu"
'            If siguiente = "e" Or siguiente = "i" Then
'                Conv_Digrafos = "24"  ' /g/
'            Else
'                Conv_Digrafos = ""
'            End If
'
'        Case "qu"
'            If siguiente = "e" Or siguiente = "i" Then
'                Conv_Digrafos = "31"  ' /k/
'            Else
'                Conv_Digrafos = ""
'            End If
'
'
'        ' ===== No reconocido =====
'        Case Else
'            Conv_Digrafos = ""
'
'    End Select
'
'End Function
'
'' ============================================================
''   PROCEDIMIENTO: Conv_Vocales
''   Vocales, diptongos, hiatos, semivocales
''   Devuelve:
''       - String con IDs separados por comas
''       - "" si no aplica
'' ============================================================
'Public Function Conv_Vocales(g As String, siguiente As String) As String
'
'    ' --- Diptongos crecientes con U ---
'    If g = "u" Then
'        If siguiente = "a" Or siguiente = "e" Or siguiente = "i" Or siguiente = "o" Then
'            Conv_Vocales = "14"   ' fonema del diptongo "ua/ue/ui/uo"
'            Exit Function
'        End If
'    End If
'
'    ' --- Diptongos decrecientes ---
'    If g = "a" And siguiente = "i" Then
'        Conv_Vocales = "12"
'        Exit Function
'    End If
'
'    If g = "a" And siguiente = "u" Then
'        Conv_Vocales = "16"
'        Exit Function
'    End If
'
'
'
'    If g = "e" And (siguiente = "i" Or siguiente = "u") Then
'        Conv_Vocales = 13 ' "14"
'        Exit Function
'    End If
'
'    If g = "o" And (siguiente = "i" Or siguiente = "u") Then
'        Conv_Vocales = "14"
'        Exit Function
'    End If
'
'    ' --- Vocal simple ---
'    Select Case g
'        Case "a": Conv_Vocales = "12"
'        Case "e": Conv_Vocales = "13"
'        Case "i": Conv_Vocales = "14"
'        Case "o": Conv_Vocales = "15"
'        Case "u": Conv_Vocales = "16"
'    End Select
'
'End Function
'
'' ============================================================
''   Función auxiliar: VocalID
'' ============================================================
'Private Function VocalID(v As String) As Byte
'    Select Case v
'        Case "a": VocalID = 12
'        Case "e": VocalID = 13
'        Case "i": VocalID = 14
'        Case "o": VocalID = 15
'        Case "u": VocalID = 16
'    End Select
'End Function
'
'' ============================================================
''   PROCEDIMIENTO: Conv_Contexto
''   Reglas contextuales (c, g, y, x...)
''   Devuelve:
''       - String con ID fonema
''       - "" si no aplica
'' ============================================================
'Private Function Conv_Contexto(g As String, anterior As String, siguiente As String) As String
'
'    g = LCase$(g)
'    anterior = LCase$(anterior)
'    siguiente = LCase$(siguiente)
'
'    ' ============================
'    '   1. C ? /k/ o /?/ (/s/)
'    ' ============================
'    If g = "c" Then
'        If siguiente = "e" Or siguiente = "i" Then
'            Conv_Contexto = "33"   ' /?/ o /s/
'        Else
'            Conv_Contexto = "31"   ' /k/
'        End If
'        Exit Function
'    End If
'
'    ' ============================
'    '   2. G ? /g/, /x/ o /gw/
'    ' ============================
'    If g = "j" Then
'        Conv_Contexto = "34"    ' /x/
'        Exit Function
'    End If
'
'    If g = "g" Then
'
'        ' --- g + ü ? /gw/ ---
'        If siguiente = "ü" Then
'            Conv_Contexto = "57"   ' /gw/
'            Exit Function
'        End If
'
'        ' --- g + e/i ? /x/ ---
'        If siguiente = "e" Or siguiente = "i" Then
'            Conv_Contexto = "34"   ' /x/
'            Exit Function
'        End If
'
'        ' --- g + a/o/u ? /g/ ---
'        Conv_Contexto = "24"
'        Exit Function
'    End If
'
'    ' ============================
'    '   3. Y ? /i/ o /?/
'    ' ============================
'    If g = "y" Then
'        If siguiente = "" Then
'            Conv_Contexto = "14"   ' /i/ final
'        Else
'            Conv_Contexto = "35"   ' /?/
'        End If
'        Exit Function
'    End If
'
'    ' ============================
'    '   4. X ? /ks/ o /x/
'    ' ============================
'    If g = "x" Then
'        If siguiente Like "[aeiou]" Then
'            Conv_Contexto = "36"   ' /ks/
'        Else
'            Conv_Contexto = "32"   ' /x/
'        End If
'        Exit Function
'    End If
'
'    ' ============================
'    '   5. S ? /s/
'    ' ============================
'    If g = "s" Then
'        Conv_Contexto = "29"
'        Exit Function
'    End If
'
'    ' ============================
'    '   6. No aplica
'    ' ============================
'    Conv_Contexto = ""
'
'End Function
'
'' ============================================================
''   PROCEDIMIENTO: Conv_Monografos
''   Convierte monógrafos (1 grafema ? 1 fonema)
''   Devuelve:
''       - String con ID fonema
''       - "" si no aplica
'' ============================================================
'Private Function Conv_Monografos(g As String) As String
'
'    g = LCase$(g)
'
'    Select Case g
'
'        ' ===== Vocales =====
'        Case "a": Conv_Monografos = "12"
'        Case "e": Conv_Monografos = "13"
'        Case "i": Conv_Monografos = "14"
'        Case "o": Conv_Monografos = "15"
'        Case "u": Conv_Monografos = "16"
'
'        ' ===== Consonantes =====
'        Case "m": Conv_Monografos = "21"
'        Case "n": Conv_Monografos = "22"
'        Case "p": Conv_Monografos = "23"
'        Case "b": Conv_Monografos = "24"
'        Case "d": Conv_Monografos = "25"
'        Case "f": Conv_Monografos = "26"
'        Case "l": Conv_Monografos = "27"
'        Case "r": Conv_Monografos = "28"
'        Case "s": Conv_Monografos = "29"
'        Case "t": Conv_Monografos = "30"
'        Case "k": Conv_Monografos = "31"
'        Case "x": Conv_Monografos = "32"
'        Case "z": Conv_Monografos = "33"
'
'        ' ===== No reconocido =====
'        Case Else
'            Conv_Monografos = ""
'
'    End Select
'
'End Function
'
'' ============================================================
''   PROCEDIMIENTO: NormalizaVocales
''   Convierte las vocales de una sílaba eliminando acentos
''   Devuelve:
''       - String con las vocales de la sílaba normalizadas
'' ============================================================
'Private Function NormalizaVocales(ByVal texto As String) As String
'
'    Dim t As String
'    Dim n As Integer
'    Dim salida As String
'
'    salida = ""
'
'    For n = 1 To Len(texto)
'
'        t = Mid(texto, n, 1)
'
'        Select Case t
'            Case "á", "à", "ä", "â"
'                t = "a"
'
'            Case "é", "è", "ë", "ê"
'                t = "e"
'
'            Case "í", "ì", "ï", "î"
'                t = "i"
'
'            Case "ó", "ò", "ö", "ô"
'                t = "o"
'
'            Case "ú", "ù", "û"   ' ü se mantiene
'                t = "u"
'
'            Case Else
'
'        End Select
'
'        salida = salida & t
'    Next
'
'    NormalizaVocales = salida
'
'End Function
'
'
''=================================================================================================
'
''-----------------
'' Auxiliares
''-----------------
'
'Private Function ContarSilabasDePalabra(ByVal palabra As String, ByRef sils() As String, ByVal offset As Long) As Long
'    Dim total As Long
'    Dim i As Long
'    Dim lenAcum As Long
'
'    total = 0
'    lenAcum = 0
'
'    For i = offset To UBound(sils)
'        lenAcum = lenAcum + Len(sils(i))
'
'        If lenAcum <= Len(palabra) Then
'            total = total + 1
'        Else
'            Exit For
'        End If
'    Next i
'
'    ContarSilabasDePalabra = total
'End Function
'
'Private Function DetectarTonicaDeUnaPalabra(ByVal palabra As String, ByVal numSilabas As Long) As Long
'
'    ' 1. Si hay tilde explícita ? esa sílaba es la tónica
'    Dim i As Long
'    Dim sils() As String
'    sils = Split(objDTO.SilabasAuto, " | ")
'
'    For i = 0 To numSilabas - 1
'        If TieneTilde(sils(i)) Then
'            DetectarTonicaDeUnaPalabra = i + 1
'            Exit Function
'        End If
'    Next i
'
'    ' 2. Si no hay tilde ? aplicar reglas generales
'    ' Palabras terminadas en vocal, n, s ? llana
'    Dim ultima As String
'    ultima = Right$(palabra, 1)
'
'If ultima Like "[aeiousn]" Then
'    If numSilabas > 1 Then
'        DetectarTonicaDeUnaPalabra = numSilabas - 1
'    Else
'        DetectarTonicaDeUnaPalabra = 1
'    End If
'Else
'    DetectarTonicaDeUnaPalabra = numSilabas
'End If
'
'End Function
'
'
'Private Function TerminaEnVocalNoSNoN(ByVal silaba As String) As Boolean
'    Dim c As String
'    c = Right$(silaba, 1)
'    TerminaEnVocalNoSNoN = (c Like "[aeiouáéíóú]")
'End Function
'
'
''Public Function EsSilabaMarcada(ByVal s As String) As Boolean
''    Dim t As String
''    t = Trim$(s)
''    EsSilabaMarcada = (Left$(t, 1) = "*" And Right$(t, 1) = "*")
''End Function
'
''Public Function LimpiarSilabaMarcada(ByVal s As String) As String
''    Dim t As String
''    t = Trim$(s)
''    t = Mid$(t, 2, Len(t) - 2)   ' quitar los dos *
''    LimpiarSilabaMarcada = Trim$(t)
''End Function
'
'
'Public Function EsSilabaMarcada(ByVal s As String) As Boolean
'    Dim t As String
'    Dim i As Long
'    Dim leadStars As Long
'    Dim trailStars As Long
'    Dim iFirstNonStar As Long
'    Dim iLastNonStar As Long
'    Dim contenido As String
'
'    t = Trim$(s)
'    If t = "" Then Exit Function
'
'    ' Contar asteriscos iniciales
'    For i = 1 To Len(t)
'        If Mid$(t, i, 1) = "(" Then
'            leadStars = leadStars + 1
'        Else
'            Exit For
'        End If
'    Next i
'
'    ' Contar asteriscos finales
'    For i = Len(t) To 1 Step -1
'        If Mid$(t, i, 1) = ")" Then
'            trailStars = trailStars + 1
'        Else
'            Exit For
'        End If
'    Next i
'
'    ' Debe haber al menos uno al principio y al final, y ser iguales
'    If leadStars = 0 Or trailStars = 0 Then Exit Function
'    If leadStars <> trailStars Then Exit Function
'
'    ' Posiciones de contenido
'    iFirstNonStar = leadStars + 1
'    iLastNonStar = Len(t) - trailStars
'
'    If iFirstNonStar > iLastNonStar Then Exit Function
'
'    contenido = Mid$(t, iFirstNonStar, iLastNonStar - iFirstNonStar + 1)
'    contenido = Trim$(contenido)
'
'    ' Contenido no vacío y sin asteriscos internos
'    If contenido = "" Then Exit Function
'    If InStr(contenido, "(") > 0 Or InStr(contenido, ")") > 0 Then Exit Function
'
'    EsSilabaMarcada = True
'End Function
'
'
'Public Function LimpiarSilabaMarcada(ByVal s As String) As String
'    Dim t As String
'    Dim i As Long
'    Dim leadStars As Long
'    Dim trailStars As Long
'    Dim iFirstNonStar As Long
'    Dim iLastNonStar As Long
'    Dim contenido As String
'
'    t = Trim$(s)
'    If t = "" Then Exit Function
'
'    ' Contar asteriscos iniciales
'    For i = 1 To Len(t)
'        If Mid$(t, i, 1) = "(" Then
'            leadStars = leadStars + 1
'        Else
'            Exit For
'        End If
'    Next i
'
'    ' Contar asteriscos finales
'    For i = Len(t) To 1 Step -1
'        If Mid$(t, i, 1) = ")" Then
'            trailStars = trailStars + 1
'        Else
'            Exit For
'        End If
'    Next i
'
'    ' Posiciones de contenido (aunque la marca sea inválida, limpiamos igual)
'    iFirstNonStar = leadStars + 1
'    iLastNonStar = Len(t) - trailStars
'
'    If iFirstNonStar > iLastNonStar Then
'        LimpiarSilabaMarcada = ""
'    Else
'        contenido = Mid$(t, iFirstNonStar, iLastNonStar - iFirstNonStar + 1)
'        LimpiarSilabaMarcada = Trim$(contenido)
'    End If
'End Function
'
'' ============================================================
''   AUXILIARES FONÉTICAS (ES)
'' ============================================================
'
'Private Function EsVocal_ES(c As String) As Boolean
'    EsVocal_ES = (InStr("aeiouáéíóú", c) > 0)
'End Function
'
'Private Function EsConsonante_ES(c As String) As Boolean
'    EsConsonante_ES = (c <> " " And Not EsVocal_ES(c))
'End Function
'
'Private Function EsVocalFuerte(c As String) As Boolean
'    EsVocalFuerte = (InStr("aeoáéó", c) > 0)
'End Function
'
'Private Function EsVocalDebil(c As String) As Boolean
'    EsVocalDebil = (InStr("iuíú", c) > 0)
'End Function
'
'Private Function EsDiptongo(c1 As String, c2 As String) As Boolean
'
'    If EsVocalDebil(c1) And EsVocalDebil(c2) Then
'        If c1 <> "í" And c1 <> "ú" And c2 <> "í" And c2 <> "ú" Then
'            EsDiptongo = True
'            Exit Function
'        End If
'    End If
'
'    If EsVocalDebil(c1) And EsVocalFuerte(c2) Then
'        If c1 <> "í" And c1 <> "ú" Then
'            EsDiptongo = True
'            Exit Function
'        End If
'    End If
'
'    If EsVocalFuerte(c1) And EsVocalDebil(c2) Then
'        If c2 <> "í" And c2 <> "ú" Then
'            EsDiptongo = True
'            Exit Function
'        End If
'    End If
'
'End Function
'
'Private Function EsTriptongo(c1 As String, c2 As String, c3 As String) As Boolean
'    If EsVocalDebil(c1) And EsVocalFuerte(c2) And EsVocalDebil(c3) Then
'        If c1 <> "í" And c1 <> "ú" And c3 <> "í" And c3 <> "ú" Then
'            EsTriptongo = True
'        End If
'    End If
'End Function
'
'Private Function EsHiatoFuerteFuerte(c1 As String, c2 As String) As Boolean
'    If EsVocalFuerte(c1) And EsVocalFuerte(c2) Then
'        EsHiatoFuerteFuerte = True
'    End If
'End Function
'
'Private Function EsGrupoInseparable_ES(par As String) As Boolean
'
'    Dim g As Variant
'    Dim lista As Variant
'
'    'lista = Array("br", "bl", "cr", "cl", "dr", "fr", "gr", "gl", "pr", "pl", "tr")
'    lista = Array("br", "bl", "cr", "cl", "dr", "fr", "fl", "gr", "gl", "pr", "pl", "tr")
'
'    For Each g In lista
'        If par = g Then
'            EsGrupoInseparable_ES = True
'            Exit Function
'        End If
'    Next g
'
'End Function
'
'' ============================================================
''   HIATOS (V10)
'' ============================================================
'Private Function EsHiato(v1 As String, v2 As String) As Boolean
'
'    ' 1) Fuerte + fuerte
'    If EsVocalFuerte(v1) And EsVocalFuerte(v2) Then
'        EsHiato = True
'        Exit Function
'    End If
'
'    ' 2) Débil tónica + fuerte
'    If (v1 = "í" Or v1 = "ú") And EsVocalFuerte(v2) Then
'        EsHiato = True
'        Exit Function
'    End If
'
'    ' 3) Fuerte + débil tónica
'    If EsVocalFuerte(v1) And (v2 = "í" Or v2 = "ú") Then
'        EsHiato = True
'        Exit Function
'    End If
'
'    EsHiato = False
'
'End Function
'
' ============================================================
' ============================================================
' ============================================================

' ============================================================
'   Rutina auxiliar de diagnóstico del motor
'   Imprime el estado completo del DTO
' ============================================================
Public Sub MF_DebugDTO(Proc As String)

    strDebug = ""

    If ObjDTO Is Nothing Then
        strDebug = strDebug & vbCrLf & "DTO no inicializado."
        Exit Sub
    End If

    strDebug = strDebug & vbCrLf
    strDebug = strDebug & vbCrLf & "==============================="
    strDebug = strDebug & vbCrLf & "   ESTADO DEL MOTOR"
    strDebug = strDebug & vbCrLf & "==============================="

    strDebug = strDebug & vbCrLf
    strDebug = strDebug & vbCrLf & "-------------------------------"
    strDebug = strDebug & vbCrLf & " Proc.: " & Proc
    strDebug = strDebug & vbCrLf & "-------------------------------"
    strDebug = strDebug & vbCrLf

    strDebug = strDebug & vbCrLf & "Texto original:        " & ObjDTO.TextoOriginal
    strDebug = strDebug & vbCrLf & "Texto Corregido:       " & ObjDTO.TextoCorregido
    strDebug = strDebug & vbCrLf & "Texto normalizado:     " & ObjDTO.TextoNormalizado
    strDebug = strDebug & vbCrLf
    strDebug = strDebug & vbCrLf & "SilabasAuto:           " & ObjDTO.SilabasAuto
    strDebug = strDebug & vbCrLf & "SilabasAcentuadas:     " & ObjDTO.SilabasAcentuadas
    strDebug = strDebug & vbCrLf & "SilabasFinal:          " & ObjDTO.SilabasFinal
    strDebug = strDebug & vbCrLf
    strDebug = strDebug & vbCrLf & "SilabaTonica:          " & ObjDTO.SilabasTonicas
    strDebug = strDebug & vbCrLf & "SilabaSecundaria:      " & ObjDTO.SilabasSecundarias
    strDebug = strDebug & vbCrLf
    strDebug = strDebug & vbCrLf & "TextoFinal (fonemas):  " & ObjDTO.FonemasFinal

    strDebug = strDebug & vbCrLf & "-------------------------------"
    strDebug = strDebug & vbCrLf & "   Detalles internos"
    strDebug = strDebug & vbCrLf & "-------------------------------"

    strDebug = strDebug & vbCrLf & "Num sílabas auto:      " & CountItems(ObjDTO.SilabasAuto, " | ") + 1
    strDebug = strDebug & vbCrLf & "Num sílabas final:     " & CountItems(ObjDTO.SilabasFinal, " | ") + 1

    strDebug = strDebug & vbCrLf & "==============================="
    strDebug = strDebug & vbCrLf   'vbCrLf

    'Stop
    
    Debug.Print strDebug
    
'    Open CurrentProject.Path & "\Debug.txt" For Output As #1
'    Print #1, strDebug
'    Close (1)
'
'    Shell "explorer " & CurrentProject.Path & "\Debug.txt"
    
End Sub

' ============================================================
' Contador auxiliar para separar elementos
' ============================================================
Private Function CountItems(ByVal s As String, ByVal sep As String) As Long
    If Len(Trim$(s)) = 0 Then
        CountItems = 0
    Else
        CountItems = UBound(Split(s, sep))
    End If
End Function


' ============================================================
' ============================================================
' ============================================================
'
