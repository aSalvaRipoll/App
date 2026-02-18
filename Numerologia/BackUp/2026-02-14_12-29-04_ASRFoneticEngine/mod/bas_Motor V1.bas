Attribute VB_Name = "bas_Motor V1"

'' ============================================================
'' Nombre:    bas-Motor_ES_Main
'' Tipo:      Módulo
'' Propósito: Módulo del motor fonético aplicado al Español (Castellano)
'' Autor:     Alba Salvá
'' Fecha:     11/02/2026
'' ============================================================
'Option Compare Database
'Option Explicit
'
'Public ObjDTO As clsDTO_Motor
'
'' ============================================================
''   ENTRADA PRINCIPAL DEL MOTOR (ESPAÑOL)
'' ============================================================
'' ----------------------------------------------------------------
'' Procedimiento: Entrada_Motor_ES
'' Propósito:     Punto de entrada al motor fonético del español general
'' Tipo proc.:    Function
'' Acceso proc.:  Public
'
'' Parameter Texto
''    (String): Texto que se recibe (nombre o apellido español general)
'
'' Tipo retorno: String -> Texto que contiene la lista de fonemas
''   resultado de la conversión
'
'' Autor:        Alba Salvá
'' Fecha:        11/02/2026
'' ----------------------------------------------------------------
'Public Function Entrada_Motor_ES(Texto As String) As String
'
'    Set ObjDTO = New clsDTO_Motor
'
'    ' 1) Asignamos el texto recibido y
'    '    Normalización (dentro del DTO)
'    ObjDTO.TextoOriginal = Texto
'    ObjDTO.NormalizaEntrada
'
'    ' 2) Silabeo automático
'    Call Silabear
'
'    Call MF_DebugDTO("Silabear")
'
'    ' 3) Devolver resultado
'    Entrada_Motor_ES = ObjDTO.SilabasAuto
'
'End Function
'
'
'' ============================================================
''  1.- SILABEO AUTOMÁTICO
'' ============================================================
'' ----------------------------------------------------------------
'' Procedimiento: Silabear (MÓDULO ÚNICO)
'' Propósito:     Convierte las palabras ya normalizadas en sílabas.
'' Tipo proc.:    Sub
'' Acceso proc.:  Private
'
'' Tipo retorno: Ninguno
'
'' Autor:        Alba Salvá
'' Fecha:        11/02/2026
'' ----------------------------------------------------------------
'Private Sub Silabear()
'
'    Dim Texto As String
'    Dim i As Long
'    Dim c As String, prev As String, nextC As String
'    Dim silaba As String
'    Dim resultado As Collection
'Dim next2 As String
'
'    Set resultado = New Collection
'
'    Texto = ObjDTO.TextoNormalizado
'
'    If Len(Texto) = 0 Then
'        ObjDTO.SilabasAuto = ""
'        Exit Sub
'    End If
'
'    silaba = ""
'
'    For i = 1 To Len(Texto)
'
'        c = Mid$(Texto, i, 1)
'
'        prev = ""
'        If i > 1 Then
'            prev = Mid$(Texto, i - 1, 1)
'        End If
'
'        nextC = ""
'        If i < Len(Texto) Then
'            nextC = Mid$(Texto, i + 1, 1)
'        End If
'
'        next2 = ""
'        If i < Len(Texto) - 1 Then
'            next2 = Mid$(Texto, i + 2, 1)
'        End If
'
'        ' --- Espacio ---
'        If c = " " Then
'            If silaba <> "" Then resultado.Add silaba
'            resultado.Add " "
'            silaba = ""
'            GoTo siguiente
'        End If
'
'        ' --- DECISIÓN DE CORTE (MODULAR) ---
'        If silaba <> "" Then
'            Dim ult As String
'            ult = Right$(silaba, 1)
'
'            'If DebeCortar(ult, c, nextC, next2) Then
'            If DebeCortar(ult, c, nextC, next2, silaba) Then
'
'                resultado.Add silaba
'                silaba = ""
'            End If
'        End If
'
'        silaba = silaba & c
'
'siguiente:
'    Next i
'
'    If silaba <> "" Then resultado.Add silaba
'
'    ' --- Construcción final ---
'    Dim out As String
'    out = ""
'
'    For i = 1 To resultado.Count
'        out = out & resultado(i)
'        If i < resultado.Count Then out = out & " | "
'    Next i
'
'    ObjDTO.SilabasAuto = out
'
'End Sub
'
''=================================================================
''       REGLAS SILABEADOR
''=================================================================
'Private Function DebeCortar(ult As String, c As String, nextC As String, next2 As String, silaba As String) As Boolean
''Private Function DebeCortar(ult As String, c As String, nextC As String, next2 As String) As Boolean
'
'' === REGLAS VOCÁLICAS (primero SIEMPRE) ===
'If Regla_Hiato(ult, c, nextC) Then GoTo cortar
'If Regla_Triptongo(ult, c, nextC) Then GoTo seguir
'If Regla_Diptongo(ult, c, nextC) Then GoTo seguir
'If Regla_VocalTonica(ult, c, nextC) Then GoTo cortar
'
'' === REGLAS DE PREFIJOS ===
'If Regla_Prefijos(silaba, c) Then GoTo cortar
'
'' === REGLAS CONSONÁNTICAS (después SIEMPRE) ===
'If Regla_NSP_NST(ult, c, nextC) Then GoTo cortar
'If Regla_SConsonante(ult, c, nextC) Then GoTo cortar
'If Regla_ClustersS(ult, c, nextC) Then GoTo cortar
'
'If Regla_VCV(ult, c, nextC) Then GoTo cortar
'If Regla_CCV(ult, c, nextC) Then GoTo cortar
'If Regla_CCC(ult, c, nextC) Then GoTo cortar
'
'' === REGLAS AVANZADAS VC + CCV / SCV / SCC ===
'If Regla_VC_CCV(ult, c, nextC, next2) Then GoTo cortar
'If Regla_VC_SCV(ult, c, nextC, next2) Then GoTo cortar
'If Regla_VC_SC(ult, c, nextC) Then GoTo cortar
'If Regla_VC_SCV2(ult, c, nextC, next2) Then GoTo cortar
'If Regla_VC_SCC(ult, c, nextC, next2) Then GoTo cortar
'
'
''    ' === REGLAS BÁSICAS ===
''    If Regla_Hiato(ult, c, nextC) Then GoTo cortar
''    If Regla_Triptongo(ult, c, nextC) Then GoTo seguir
''    If Regla_Diptongo(ult, c, nextC) Then GoTo seguir
''
''    If Regla_VCV(ult, c, nextC) Then GoTo cortar
''    If Regla_CCV(ult, c, nextC) Then GoTo cortar
''    If Regla_CCC(ult, c, nextC) Then GoTo cortar
''
''    ' === REGLAS AVANZADAS (FASE 1) ===
''    If Regla_NSP_NST(ult, c, nextC) Then GoTo cortar
''    If Regla_SConsonante(ult, c, nextC) Then GoTo cortar
''    If Regla_ClustersS(ult, c, nextC) Then GoTo cortar
''    If Regla_VocalTonica(ult, c, nextC) Then GoTo cortar
''
''    ' === REGLAS AVANZADAS (FASE 2) ===
''    If Regla_VC_CCV(ult, c, nextC, next2) Then GoTo cortar
''    If Regla_VC_SCV(ult, c, nextC, next2) Then GoTo cortar
''    If Regla_VC_SC(ult, c, nextC) Then GoTo cortar
''    If Regla_VC_SCV2(ult, c, nextC, next2) Then GoTo cortar
''    If Regla_VC_SCC(ult, c, nextC, next2) Then GoTo cortar
''    If Regla_Prefijos(silaba, c) Then GoTo cortar
'
'seguir:
'    DebeCortar = False
'    Exit Function
'
'cortar:
'    DebeCortar = True
'
'End Function
'
''-----------------------------------------------------------------
'' 1.- REGLAS BÁSICAS
''-----------------------------------------------------------------
'Private Function Regla_Hiato(ult As String, c As String, nextC As String) As Boolean
'
'    Debug.Print
'    Debug.Print "Inicio Regla_Hiato"; " --> "; ult; " - "; c; " - "; nextC
'
'    If EsVocal(ult) And EsVocal(c) Then
'        If EsHiato(ult, c) Then Regla_Hiato = True
'    End If
'
'    Debug.Print "Fin Regla_Hiato"; " >>> "; Regla_Hiato
'
'End Function
'
'Private Function Regla_Triptongo(ult As String, c As String, nextC As String) As Boolean
'
'    Debug.Print
'    Debug.Print "Inicio Regla_Triptongo"; " --> "; ult; " - "; c; " - "; nextC
'
'    If EsVocalDebil(ult) And EsVocalFuerte(c) And EsVocalDebil(nextC) Then
'        Regla_Triptongo = True
'    End If
'
'    Debug.Print "Fin Regla_Triptongo"; " >>> "; Regla_Triptongo
'
'End Function
'
'Private Function Regla_Diptongo(ult As String, c As String, nextC As String) As Boolean
'
'    Debug.Print
'    Debug.Print "Inicio Regla_Diptongo"; " --> "; ult; " - "; c; " - "; nextC
'
'    If EsDiptongo(ult, c) Then Regla_Diptongo = True
'
'    Debug.Print "Fin Regla_Diptongo"; " >>> "; Regla_Diptongo
'
'End Function
'
'Private Function Regla_VCV(ult As String, c As String, nextC As String) As Boolean
'
'    Debug.Print
'    Debug.Print "Inicio Regla_VCV"; " --> "; ult; " - "; c; " - "; nextC
'
'    ' 1. Si la vocal siguiente es tónica ? NO cortar
'    If nextC Like "[áéíóú]" Then Exit Function
'
'    ' 2. Si hay hiato ? NO cortar (ya lo gestiona Regla_Hiato)
'    If EsHiato(ult, nextC) Then Exit Function
'
'    ' 3. Si hay diptongo ? NO cortar (ya lo gestiona Regla_Diptongo)
'    If EsDiptongo(ult, nextC) Then Exit Function
'
'    ' 4. Si la consonante forma ataque complejo ? NO cortar
'    If EsGrupoInseparable(c & nextC) Then Exit Function
'
'    ' 5. VCV normal ? cortar
'    If EsVocal(ult) And EsConsonante(c) And EsVocal(nextC) Then
'        Regla_VCV = True
'    End If
'
'    Debug.Print "Fin Regla_VCV"; " >>> "; Regla_VCV
'
'End Function
'
''Private Function Regla_VCV(ult As String, c As String, nextC As String) As Boolean
''
''    ' No cortar si la vocal siguiente es tónica
''    If nextC Like "[áéíóú]" Then Exit Function
''
''    ' Patrón VCV normal
''    If EsVocal(ult) And EsConsonante(c) And EsVocal(nextC) Then
''        Regla_VCV = True
''    End If
''
''End Function
'
''Private Function Regla_VCV(ult As String, c As String, nextC As String) As Boolean
''    If EsVocal(ult) And EsConsonante(c) And EsVocal(nextC) Then
''        Regla_VCV = True
''    End If
''End Function
'
'Private Function Regla_CCV(ult As String, c As String, nextC As String) As Boolean
'
'    Debug.Print
'    Debug.Print "Inicio Regla_CCV"; " --> "; ult; " - "; c; " - "; nextC
'
'    If EsConsonante(ult) And EsConsonante(c) And EsVocal(nextC) Then
'        If Not EsGrupoInseparable(ult & c) Then
'            Regla_CCV = True
'        End If
'    End If
'
'    Debug.Print "Fin Regla_CCV"; " >>> "; Regla_CCV
'
'End Function
'
'Private Function Regla_CCC(ult As String, c As String, nextC As String) As Boolean
'
'    Debug.Print
'    Debug.Print "Inicio Regla_CCC"; " --> "; ult; " - "; c; " - "; nextC
'
'    If EsConsonante(ult) And EsConsonante(c) And EsConsonante(nextC) Then
'        Regla_CCC = True
'    End If
'
'    Debug.Print "Fin Regla_CCC"; " >>> "; Regla_CCC
'
'End Function
'
''-----------------------------------------------------------------
'' 2.- REGLAS AVANZADAS (FASE 1)
''-----------------------------------------------------------------
'
'Private Function Regla_NSP_NST(ult As String, c As String, nextC As String) As Boolean
'
'    Debug.Print
'    Debug.Print "Inicio Regla_NSP_NST"; " --> "; ult; " - "; c; " - "; nextC
'
'    If ult = "n" And c = "s" Then
'        If nextC = "p" Or nextC = "t" Then
'            Regla_NSP_NST = True
'        End If
'    End If
'
'    Debug.Print "Fin Regla_NSP_NST"; " >>> "; Regla_NSP_NST
'
'End Function
'
'Private Function Regla_SConsonante(ult As String, c As String, nextC As String) As Boolean
'
'    Debug.Print
'    Debug.Print "Inicio Regla_SConsonante"; " --> "; ult; " - "; c; " - "; nextC
'
'    If ult = "s" And EsConsonante(c) Then
'        If EsVocal(nextC) Then
'            Regla_SConsonante = True
'        End If
'    End If
'
'    Debug.Print "Fin Regla_SConsonante"; " >>> "; Regla_SConsonante
'
'End Function
'
'Private Function Regla_ClustersS(ult As String, c As String, nextC As String) As Boolean
'
'    Debug.Print
'    Debug.Print "Inicio Regla_ClustersS"; " --> "; ult; " - "; c; " - "; nextC
'
'    Dim par As String
'    par = c & nextC
'
'    If ult = "s" Then
'        Select Case par
'            Case "tr", "pr", "pl", "cr", "cl", "gr", "fr"
'                Regla_ClustersS = True
'        End Select
'    End If
'
'    Debug.Print "Fin Regla_ClustersS"; " >>> "; Regla_ClustersS
'
'End Function
'
'Private Function Regla_VocalTonica(ult As String, c As String, nextC As String) As Boolean
'
'    Debug.Print
'    Debug.Print "Inicio Regla_VocalTonica"; " --> "; ult; " - "; c; " - "; nextC
'
'    ' No cortar si ult forma parte de un ataque complejo
'    If EsGrupoInseparable(ult & c) Then Exit Function
'
'    ' Cortar si consonante + vocal tónica
'    If EsConsonante(ult) And c Like "[áéíóú]" Then
'        Regla_VocalTonica = True
'    End If
'
'    Debug.Print "Fin Regla_VocalTonica"; " >>> "; Regla_VocalTonica
'
'End Function
'
''Private Function Regla_VocalTonica(ult As String, c As String, nextC As String) As Boolean
''    If EsConsonante(ult) And c Like "[áéíóú]" Then
''        Regla_VocalTonica = True
''    End If
''End Function
'
''-----------------------------------------------------------------
'' 3.- REGLAS AVANZADAS (FASE 2)
''-----------------------------------------------------------------
'
'' V C + C C V
'Private Function Regla_VC_CCV(ult As String, c As String, nextC As String, next2 As String) As Boolean
'    ' ult = vocal
'    ' c = consonante
'    ' nextC = consonante
'    ' next2 = vocal
'
'    Debug.Print
'    Debug.Print "Inicio Regla_VC_CCV"; " --> "; ult; " - "; c; " - "; nextC; " - "; next2
'
'    If EsVocal(ult) And EsConsonante(c) And EsConsonante(nextC) And EsVocal(next2) Then
'        Regla_VC_CCV = True
'    End If
'
'    Debug.Print "Fin Regla_VC_CCV"; " >>> "; Regla_VC_CCV
'
'End Function
'
'' V C + S C V
'Private Function Regla_VC_SCV(ult As String, c As String, nextC As String, next2 As String) As Boolean
'
'    Debug.Print
'    Debug.Print "Inicio Regla_VC_SCV"; " --> "; ult; " - "; c; " - "; nextC; " - "; next2
'
'    If EsVocal(ult) And EsConsonante(c) And nextC = "s" And EsConsonante(next2) Then
'        Regla_VC_SCV = True
'    End If
'
'    Debug.Print "Fin Regla_VC_SCV"; " >>> "; Regla_VC_SCV
'
'End Function
'
'' V C + S + C
'Private Function Regla_VC_SC(ult As String, c As String, nextC As String) As Boolean
'
'    Debug.Print
'    Debug.Print "Inicio Regla_VC_SC"; " --> "; ult; " - "; c; " - "; nextC '; " - "; next2
'
'    If EsVocal(ult) And EsConsonante(c) And nextC = "s" Then
'        Regla_VC_SC = True
'    End If
'
'    Debug.Print "Fin Regla_VC_SC"; " >>> "; Regla_VC_SC
'
'End Function
'
'' V C + S + C + V
'Private Function Regla_VC_SCV2(ult As String, c As String, nextC As String, next2 As String) As Boolean
'
'    Debug.Print
'    Debug.Print "Inicio Regla_VC_SCV2"; " --> "; ult; " - "; c; " - "; nextC; " - "; next2
'
'    If EsVocal(ult) And EsConsonante(c) And nextC = "s" And EsConsonante(next2) Then
'        Regla_VC_SCV2 = True
'    End If
'
'    Debug.Print "Fin Regla_VC_SCV2"; " >>> "; Regla_VC_SCV2
'
'End Function
'
''V C + S + C + C
'Private Function Regla_VC_SCC(ult As String, c As String, nextC As String, next2 As String) As Boolean
'
'    Debug.Print
'    Debug.Print "Inicio Regla_VC_SCC"; " --> "; ult; " - "; c; " - "; nextC; " - "; next2
'
'    If EsVocal(ult) And EsConsonante(c) And nextC = "s" And EsConsonante(next2) Then
'        Regla_VC_SCC = True
'    End If
'
'    Debug.Print "Fin Regla_VC_SCC"; " >>> "; Regla_VC_SCC
'
'End Function
'
'' Prefijos comunes: anti-, intro-, trans-, contra-, extra-, pre-,  pro-, sub-
'Private Function Regla_Prefijos(silaba As String, c As String) As Boolean
'
'    Debug.Print
'    Debug.Print "Inicio Regla_Prefijos"; " --> "; silaba; " - "; c '; " - "; nextC; " - "; next2
'
'    Dim prefijos
'    prefijos = Array("anti", "intro", "trans", "contra", "extra", "pre", "pro", "sub")
'
'    Dim p As Variant
'    For Each p In prefijos
'        If silaba = p Then
'            Regla_Prefijos = True
'            Exit Function
'        End If
'    Next p
'
'    Debug.Print "Fin Regla_Prefijos"; " >>> "; Regla_Prefijos
'
'End Function
'
'
'
'
'
'
''=================================================================
''=================================================================
''       FUNCIONES AUXILIARES GENERALES
''=================================================================
''=================================================================
'' ----------------------------------------------------------------
'' Procedimiento: EsVocal
'' Propósito:     Indica si una letra es vocal o no
'' Tipo proc.:    Function
'' Acceso proc.:  Private
'
'' Parameter: c (String): Letra a validar
'
'
'' Tipo retorno: Boolean
'
'' Autor:        Alba Salvá
'' Fecha:        11/02/2026
'' ----------------------------------------------------------------
'Private Function EsVocal(c As String) As Boolean
'    EsVocal = (c Like "[aeiouáéíóú]")
'End Function
'
'Private Function EsVocalFuerte(c As String) As Boolean
'    EsVocalFuerte = (c Like "[aeoáéó]")
'End Function
'
'Private Function EsVocalDebil(c As String) As Boolean
'    EsVocalDebil = (c Like "[iuíú]")
'End Function
'
'Private Function EsConsonante(c As String) As Boolean
'    EsConsonante = (Not EsVocal(c) And c <> " ")
'End Function
'
'Private Function EsDiptongo(v1 As String, v2 As String) As Boolean
'    If EsVocalDebil(v1) And EsVocalDebil(v2) Then
'        EsDiptongo = True
'    ElseIf EsVocalFuerte(v1) And EsVocalDebil(v2) Then
'        EsDiptongo = True
'    ElseIf EsVocalDebil(v1) And EsVocalFuerte(v2) Then
'        EsDiptongo = True
'    End If
'End Function
'
'Private Function EsHiato(v1 As String, v2 As String) As Boolean
'    If EsVocalFuerte(v1) And EsVocalFuerte(v2) Then
'        EsHiato = True
'    ElseIf v1 Like "[íú]" And EsVocalFuerte(v2) Then
'        EsHiato = True
'    End If
'End Function
'
'Private Function EsGrupoInseparable(par As String) As Boolean
'    Select Case par
'        Case "dr", "tr", "gr", "pr", "pl", "cl", "fr", "fl", "br", "bl"
'            EsGrupoInseparable = True
'    End Select
'End Function
'
'_____________________________________________________________________________________________________________________

'Private Sub MarcarTonicaYSecundariaEnCadena()
'
'    Dim sils As Variant
'    Dim palabras As New Collection
'    Dim palabraActual As New Collection
'    Dim i As Long, globalIndex As Long
'    Dim arrT() As String
'    Dim arrS() As String
'
'
'    sils = Split(ObjDTO.SilabasAuto, " | ")
'
'    ' 1. Separar sílabas por palabra
'    For i = LBound(sils) To UBound(sils)
'
'        If Trim$(sils(i)) = "" Then
'            If palabraActual.Count > 0 Then palabras.Add palabraActual
'            Set palabraActual = New Collection
'        Else
'            palabraActual.Add sils(i)
'        End If
'
'    Next i
'
'    If palabraActual.Count > 0 Then palabras.Add palabraActual
'
'    ' 2. Calcular índices globales
'    Dim tGlobal As New Collection
'    Dim sGlobal As New Collection
'
'    globalIndex = 0
'
'    Dim p As Long
'    For p = 1 To palabras.Count
'
'        Dim w As Collection
'        Set w = palabras(p)
'
'        Dim tLocal As Long
'        Dim secs As Collection
'
'        tLocal = DetectarTonica(w)
'        Set secs = DetectarSecundarias(w, tLocal)
'
'        ' Tónica global
'        tGlobal.Add globalIndex + tLocal
'
'        ' Secundarias globales
'        Dim s As Variant
'        For Each s In secs
'            sGlobal.Add globalIndex + CLng(s)
'        Next s
'
'        globalIndex = globalIndex + w.Count
'
'    Next p
'
'    ' 3. Reconstruir cadena marcada
'    Dim out As String
'    out = ""
'
'    Dim g As Long
'    g = 1
'
'    For i = LBound(sils) To UBound(sils)
'
'        If Trim$(sils(i)) = "" Then
'            out = out & " | "
'        Else
'            Dim marcado As String
'            marcado = sils(i)
'
'            Dim esTonica As Boolean: esTonica = False
'            Dim esSec As Boolean: esSec = False
'
'            ' ¿Es tónica?
'            Dim x As Variant
'            For Each x In tGlobal
'                If x = g Then esTonica = True
'            Next x
'
'            ' ¿Es secundaria?
'            For Each x In sGlobal
'                If x = g Then esSec = True
'            Next x
'
'            If esTonica Then
'                marcado = "( " & marcado & " )"
'            ElseIf esSec Then
'                marcado = "[ " & marcado & " ]"
'            End If
'
'            out = out & marcado
'
'            If i < UBound(sils) Then out = out & " | "
'
'            g = g + 1
'        End If
'
'    Next i
'
'' Guardar índices globales en el DTO
'If tGlobal.Count > 0 Then
'    ReDim arrT(1 To tGlobal.Count)
'
'    For i = 1 To tGlobal.Count
'        arrT(i) = CStr(tGlobal(i))
'    Next i
'    ObjDTO.SilabaTonica = Join(arrT, ",")
'Else
'    ObjDTO.SilabaTonica = ""
'End If
'
'If sGlobal.Count > 0 Then
'    ReDim arrS(1 To sGlobal.Count)
'
'    For i = 1 To sGlobal.Count
'        arrS(i) = CStr(sGlobal(i))
'    Next i
'    ObjDTO.SilabaSecundaria = Join(arrS, ",")
'Else
'    ObjDTO.SilabaSecundaria = ""
'End If
'
'    ObjDTO.SilabasFinal = out
'
'End Sub


