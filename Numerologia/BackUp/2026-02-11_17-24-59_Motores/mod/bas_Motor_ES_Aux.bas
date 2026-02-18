Attribute VB_Name = "bas_Motor_ES_Aux"
Option Compare Database
Option Explicit










'' ============================================================
''   2- SILABEO AUTOMÁTICO (VERSIÓN MODULAR) V X19
'' ============================================================
'Private Sub SilabearAuto()
'
'    Dim Texto As String
'    Dim col As New Collection
'    Dim i As Long, ini As Long
'    Dim c1 As String, c2 As String, c3 As String
'    Dim par As String
'    Dim arr() As String
'
''    Debug.Print
''    Debug.Print ">>> SilabearAuto Iniciando V10"
'
'    Texto = Trim$(ObjDTO.TextoNormalizado)
'    If Len(Texto) = 0 Then
'        ObjDTO.SilabasAuto = ""
'        Exit Sub
'    End If
'
'    ini = 1
'
'    For i = 2 To Len(Texto)
'
'        c1 = Mid$(Texto, i - 1, 1)
'        c2 = Mid$(Texto, i, 1)
'        If i < Len(Texto) Then c3 = Mid$(Texto, i + 1, 1) Else c3 = ""
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
'        ' 3. GRUPOS INSEPARABLES
'        ' --------------------------------------------------------
'        If EsConsonante_ES(c1) And EsConsonante_ES(c2) Then
'            If EsGrupoInseparable_ES(par) Then GoTo siguiente
'        End If
'
'        ' --------------------------------------------------------
'        ' 4. CCC
'        ' --------------------------------------------------------
'        If EsCCC(c1, c2, c3) Then
'            col.Add Array(ini, i - 1)
'            ini = i
'            GoTo siguiente
'        End If
'
'        ' --------------------------------------------------------
'        ' 5. EXCEPCIÓN: V + C + V TILDADA (ción)
'        ' --------------------------------------------------------
'        If EsVocal_ES(c1) And EsConsonante_ES(c2) And EsVocal_ES(c3) Then
'            If c3 Like "[áéíóú]" Then GoTo siguiente
'        End If
'
'        ' --------------------------------------------------------
'        ' 6. TRIPTONGO
'        ' --------------------------------------------------------
'        If i < Len(Texto) - 1 Then
'            If EsTriptongo(c1, c2, Mid$(Texto, i + 1, 1)) Then GoTo siguiente
'        End If
'
'        ' --------------------------------------------------------
'        ' 7. HIATOS (nuevo bloque)
'        ' --------------------------------------------------------
'        If EsVocal_ES(c1) And EsVocal_ES(c2) Then
'
'            ' No separar si es diptongo
'            If EsDiptongo(c1, c2) Then GoTo siguiente
'
'            ' Separar si es hiato
'            If EsHiato(c1, c2) Then
'                col.Add Array(ini, i - 1)
'                ini = i
'                GoTo siguiente
'            End If
'
'            ' Separación por defecto
'            col.Add Array(ini, i - 1)
'            ini = i
'            GoTo siguiente
'        End If
'
'        ' --------------------------------------------------------
'        ' 8. DIPTONGO
'        ' --------------------------------------------------------
'        If EsDiptongo(c1, c2) Then GoTo siguiente
'
'        ' --------------------------------------------------------
'        ' 9. V + RR
'        ' --------------------------------------------------------
'        If EsVocal_ES(c1) And c2 = "r" Then
'            If i < Len(Texto) Then
'                If Mid$(Texto, i + 1, 1) = "r" Then
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
'                If Mid$(Texto, i - 2, 1) = "n" Then GoTo siguiente
'            End If
'        End If
'
'        ' --------------------------------------------------------
'        ' 11. CCV ? C | CV
'        ' --------------------------------------------------------
'        If EsConsonante_ES(c1) And EsConsonante_ES(c2) Then
'            If i < Len(Texto) Then
'                If EsVocal_ES(c3) Then
'                    If Not EsGrupoInseparable_ES(par) Then
'
'                        ' <<< PARCHE DE SEGURIDAD >>>
'                        If i - 1 >= ini Then
'                            col.Add Array(ini, i - 1)
'                        End If
'
'                        ini = i
'                        GoTo siguiente
'                    End If
'                End If
'            End If
'        End If

'        If EsConsonante_ES(c1) And EsConsonante_ES(c2) Then
'            If i < Len(texto) Then
'                If EsVocal_ES(c3) Then
'                    If Not EsGrupoInseparable_ES(par) Then
'                        col.Add Array(ini, i - 1)
'                        ini = i
'                        GoTo siguiente
'                    End If
'                End If
'            End If
'        End If
'
'        ' --------------------------------------------------------
'        ' 12. VCV ? V | CV
'        ' --------------------------------------------------------
'        If EsVocal_ES(c1) And EsConsonante_ES(c2) Then
'            If i < Len(Texto) Then
'                If EsVocal_ES(c3) Then
'                    col.Add Array(ini, i - 1)
'                    ini = i
'                    GoTo siguiente
'                End If
'            End If
'        End If
'
'siguiente:
'    Next i
'
'    If ini <= Len(Texto) Then col.Add Array(ini, Len(Texto))
'
''    ' >>> AQUI, JUSTO AQUI <<<
''    Debug.Print "---- Detalle índices col ----"
''    For i = 1 To col.Count
''        Debug.Print i, "ini:", col(i)(0), "fin:", col(i)(1), _
''                     "texto:", Mid$(texto, col(i)(0), col(i)(1) - col(i)(0) + 1)
''    Next i
''    Debug.Print "-----------------------------"
'
'    ReDim arr(1 To col.Count)
'    For i = 1 To col.Count
'        arr(i) = Mid$(Texto, col(i)(0), col(i)(1) - col(i)(0) + 1)
'    Next i
'
'' --- FUSIÓN DE S SUELTA ---
'Dim salida() As String
'Dim tx As String
'Dim n As Long
'
'ReDim salida(UBound(arr))
'n = 0
'
'For i = 1 To UBound(arr)
'
'    If arr(i) = "s" Then
'        ' Fusionar con el anterior (si existe)
'        If n > 0 Then
'            salida(n - 1) = salida(n - 1) & "s"
'            n = n - 1
'        Else
'            ' Caso imposible en español, pero por seguridad
'            salida(n) = "s"
'        End If
'
'    Else
'        salida(n) = arr(i)
'    End If
'
'    n = n + 1
'Next i
'
'' Construir la salida final
'tx = ""
'For i = 0 To n - 1
'    If Trim$(salida(i)) <> "" Then
'        tx = tx & salida(i) & " | "
'    End If
'Next i
'
'' Quitar la última barra y espacios
'tx = Trim$(tx)
'If Right$(tx, 1) = "|" Then tx = Trim$(Left$(tx, Len(tx) - 1))
'
'
'    ObjDTO.SilabasAuto = tx ' Join(arr, " | ")
'
''    Debug.Print ">>> SilabearAuto ejecutado"
'
'End Sub

