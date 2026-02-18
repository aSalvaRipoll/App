Attribute VB_Name = "bas-Motor_ES_Main"
' ============================================================
' Nombre:    bas-Motor_ES_Main
' Tipo:      Módulo
' Propósito: Motor de silabeo ortográfico del español
' Autor:     Alba Salvá
' Fecha:     11/02/2026
' Versión:   3 (Depuración activada)
' ============================================================

Option Compare Database
Option Explicit

Public ObjDTO As clsDTO_Motor

Public strDebug As String


' ============================================================
'   ENTRADA PRINCIPAL DEL MOTOR (ESPAÑOL)
' ============================================================
Public Function Entrada_Motor_ES(texto As String) As String

        
    Set ObjDTO = New clsDTO_Motor

    ObjDTO.TextoOriginal = texto
    ObjDTO.NormalizaEntrada

    Call Silabear
    
    Call CalcularTonicas
    Call CalcularSecundarias
    Call MarcarTonicaYSecundariaEnCadena
    Call ConstruirCadenaFonemas
    Call MF_DebugDTO("MarcarTonicaYSecundariaEnCadena")

    Entrada_Motor_ES = ObjDTO.SilabasAuto

    
End Function


' ============================================================
'   1.- SILABEO AUTOMÁTICO
' ============================================================
Private Sub Silabear()

    Dim texto As String
    Dim i As Long
    Dim c As String, prev As String, nextC As String, next2 As String, next3 As String
    Dim silaba As String
    Dim resultado As Collection

    Set resultado = New Collection
    texto = ObjDTO.TextoNormalizado

    If Len(texto) = 0 Then
        ObjDTO.SilabasAuto = ""
        Exit Sub
    End If

    silaba = ""

    For i = 1 To Len(texto)

        c = Mid$(texto, i, 1)

        prev = ""
        If i > 1 Then prev = Mid$(texto, i - 1, 1)

        nextC = ""
        If i < Len(texto) Then nextC = Mid$(texto, i + 1, 1)

        next2 = ""
        If i < Len(texto) - 1 Then next2 = Mid$(texto, i + 2, 1)


        next3 = ""
        If i < Len(texto) - 2 Then next3 = Mid$(texto, i + 3, 1)

        ' --- Espacio ---
        If c = " " Then
            If silaba <> "" Then resultado.Add silaba
            resultado.Add " "
            silaba = ""
            GoTo siguiente
        End If

        ' --- DECISIÓN DE CORTE ---
        If silaba <> "" Then
            Dim ult As String
            ult = Right$(silaba, 1)

            Select Case DebeCortar(ult, c, nextC, next2, silaba, next3)
                Case 1
                    ' Corte antes de c
                    resultado.Add silaba
                    silaba = ""
    
                Case 2
                    ' Corte entre c y nextC
                    silaba = silaba & c
                    resultado.Add silaba
                    silaba = ""
                    GoTo siguiente   ' NO añadir c otra vez
                Case Else
            
            End Select
        
        End If

        silaba = silaba & c

siguiente:
    Next i

    If silaba <> "" Then resultado.Add silaba

    ' --- Construcción final ---
    Dim out As String
    out = ""

    For i = 1 To resultado.Count
        out = out & resultado(i)
        If i < resultado.Count Then out = out & " | "
    Next i

    ObjDTO.SilabasAuto = out

End Sub


' ============================================================
'   REGLAS DE SILABEO
' ============================================================
Private Function DebeCortar(ult As String, c As String, nextC As String, next2 As String, silaba As String, next3 As String) As Byte

' === REGLAS VOCÁLICAS ===
If Regla_Hiato(ult, c, nextC) Then DebeCortar = 1: Exit Function
If Regla_Triptongo(ult, c, nextC) Then DebeCortar = 0: Exit Function
If Regla_Diptongo(ult, c, nextC) Then DebeCortar = 0: Exit Function
If Regla_VocalTonica(ult, c, nextC) Then DebeCortar = 1: Exit Function

' === PREFIJOS ===
If Regla_Prefijos(silaba, c) Then DebeCortar = 1: Exit Function

' === REGLAS CONSONÁNTICAS ===
If Regla_NSP_NST(ult, c, nextC) Then DebeCortar = 1: Exit Function
If Regla_SConsonante(ult, c, nextC) Then DebeCortar = 1: Exit Function
If Regla_ClustersS(ult, c, nextC) Then DebeCortar = 1: Exit Function

If Regla_VCV(ult, c, nextC) Then DebeCortar = 1: Exit Function

' === REGLAS CONSONÁNTICAS ESPECIALES ===
If Regla_CortarAntesDeGrupo(ult, c, nextC, next2, next3) Then DebeCortar = 1: Exit Function

' === ESTA DEBE IR AQUÍ (ANTES QUE CCV) ===
If Regla_VC_CV(ult, c, nextC, next2) Then DebeCortar = 2: Exit Function

' === AHORA SÍ CCV ===
If Regla_CCV(ult, c, nextC) Then DebeCortar = 1: Exit Function
If Regla_CCC(ult, c, nextC) Then DebeCortar = 1: Exit Function

DebeCortar = 0

'' === REGLAS VOCÁLICAS ===
'If Regla_Hiato(ult, c, nextC) Then DebeCortar = 1: Exit Function
'If Regla_Triptongo(ult, c, nextC) Then DebeCortar = 0: Exit Function
'If Regla_Diptongo(ult, c, nextC) Then DebeCortar = 0: Exit Function
'If Regla_VocalTonica(ult, c, nextC) Then DebeCortar = 1: Exit Function
'
'' === PREFIJOS ===
'If Regla_Prefijos(silaba, c) Then DebeCortar = 1: Exit Function
'
'' === REGLAS CONSONÁNTICAS ===
'If Regla_NSP_NST(ult, c, nextC) Then DebeCortar = 1: Exit Function
'If Regla_SConsonante(ult, c, nextC) Then DebeCortar = 1: Exit Function
'If Regla_ClustersS(ult, c, nextC) Then DebeCortar = 1: Exit Function
'
'' === REGLA VC-CV (corte entre consonantes) ===
'If Regla_VC_CV(ult, c, nextC, next2) Then DebeCortar = 2: Exit Function
'
'If Regla_VCV(ult, c, nextC) Then DebeCortar = 1: Exit Function
'If Regla_CCV(ult, c, nextC) Then DebeCortar = 1: Exit Function
'If Regla_CCC(ult, c, nextC) Then DebeCortar = 1: Exit Function
'
'DebeCortar = 0

End Function


' ============================================================
'   REGLAS BÁSICAS
' ============================================================
Private Function Regla_Hiato(ult As String, c As String, nextC As String) As Boolean
    'strDebug = strDebug & vbCrLf: 'strDebug = strDebug & vbCrLf & "Inicio Regla_Hiato --> " & ult & " - " & c & " - " & nextC
    If EsVocal(ult) And EsVocal(c) Then
        If EsHiato(ult, c) Then Regla_Hiato = True
    End If
    'strDebug = strDebug & vbCrLf & "Fin Regla_Hiato >>> " & Regla_Hiato
End Function

Private Function Regla_Triptongo(ult As String, c As String, nextC As String) As Boolean
    
    'strDebug = strDebug & vbCrLf
    'strDebug = strDebug & vbCrLf & "Inicio Regla_Triptongo --> " & ult & " - " & c & " - " & nextC
    If EsVocalDebil(ult) And EsVocalFuerte(c) And EsVocalDebil(nextC) Then Regla_Triptongo = True
    'strDebug = strDebug & vbCrLf & "Fin Regla_Triptongo >>> " & Regla_Triptongo
    
End Function

Private Function Regla_Diptongo(ult As String, c As String, nextC As String) As Boolean
    
    'strDebug = strDebug & vbCrLf
    'strDebug = strDebug & vbCrLf & "Inicio Regla_Diptongo --> " & ult & " - " & c & " - " & nextC
    If EsDiptongo(ult, c) Then Regla_Diptongo = True
    'strDebug = strDebug & vbCrLf & "Fin Regla_Diptongo >>> " & Regla_Diptongo
    
End Function

' ============================================================
'   REGLA VCV (ORTOGRÁFICA)
' ============================================================
Private Function Regla_VCV(ult As String, c As String, nextC As String) As Boolean

    'strDebug = strDebug & vbCrLf
    'strDebug = strDebug & vbCrLf & "Inicio Regla_VCV --> " & ult & " - " & c & " - " & nextC

    ' ORTOGRÁFICO: VCV SIEMPRE CORTA
    ' salvo que la consonante forme grupo inseparable (pr, tr, cl, etc.)
    If EsGrupoInseparable(c & nextC) Then Exit Function

    If EsVocal(ult) And EsConsonante(c) And EsVocal(nextC) Then
        Regla_VCV = True
    End If

    'strDebug = strDebug & vbCrLf & "Fin Regla_VCV >>> " & Regla_VCV


End Function

' ============================================================
'   REGLA VOCAL TÓNICA (ORTOGRÁFICA)
' ============================================================
Private Function Regla_VocalTonica(ult As String, c As String, nextC As String) As Boolean

    'strDebug = strDebug & vbCrLf
    'strDebug = strDebug & vbCrLf & "Inicio Regla_VocalTonica --> " & ult & " - " & c & " - " & nextC

    ' De momento, para silabeo ortográfico, NO usamos esta regla
    Regla_VocalTonica = False

    'strDebug = strDebug & vbCrLf & "Fin Regla_VocalTonica >>> " & Regla_VocalTonica

End Function


' ============================================================
'   REGLAS CONSONÁNTICAS Y AVANZADAS
' ============================================================
' (Aquí van todas tus reglas NSP/NST, SConsonante, ClustersS,
'  CCV, CCC, VC_CCV, VC_SCV, VC_SC, VC_SCV2, VC_SCC)
'  — Las dejo tal cual las tienes, solo con 'strDebug = strDebug & vbcrlf  activado.

'-----------------------------------------------------------------
' 2.- REGLAS AVANZADAS (FASE 1)
'-----------------------------------------------------------------

Private Function Regla_NSP_NST(ult As String, c As String, nextC As String) As Boolean

    'strDebug = strDebug & vbCrLf
    'strDebug = strDebug & vbCrLf & "Inicio Regla_NSP_NST" & " --> " & ult & " - " & c & " - " & nextC

    If ult = "n" And c = "s" Then
        If nextC = "p" Or nextC = "t" Then
            Regla_NSP_NST = True
        End If
    End If

    'strDebug = strDebug & vbCrLf & "Fin Regla_NSP_NST" & " >>> " & Regla_NSP_NST

End Function

Private Function Regla_SConsonante(ult As String, c As String, nextC As String) As Boolean

    'strDebug = strDebug & vbCrLf
    'strDebug = strDebug & vbCrLf & "Inicio Regla_SConsonante" & " --> " & ult & " - " & c & " - " & nextC

    If ult = "s" And EsConsonante(c) Then
        If EsVocal(nextC) Then
            Regla_SConsonante = True
        End If
    End If

    'strDebug = strDebug & vbCrLf & "Fin Regla_SConsonante" & " >>> " & Regla_SConsonante

End Function

Private Function Regla_ClustersS(ult As String, c As String, nextC As String) As Boolean

    'strDebug = strDebug & vbCrLf
    'strDebug = strDebug & vbCrLf & "Inicio Regla_ClustersS" & " --> " & ult & " - " & c & " - " & nextC

    Dim par As String
    par = c & nextC

    If ult = "s" Then
        Select Case par
            Case "tr", "pr", "pl", "cr", "cl", "gr", "fr"
                Regla_ClustersS = True
        End Select
    End If

    'strDebug = strDebug & vbCrLf & "Fin Regla_ClustersS" & " >>> " & Regla_ClustersS

End Function

Private Function Regla_CCV(ult As String, c As String, nextC As String) As Boolean

    'strDebug = strDebug & vbCrLf
    'strDebug = strDebug & vbCrLf & "Inicio Regla_CCV" & " --> " & ult & " - " & c & " - " & nextC

    If EsConsonante(ult) And EsConsonante(c) And EsVocal(nextC) Then
        If Not EsGrupoInseparable(ult & c) Then
            Regla_CCV = True
        End If
    End If

    'strDebug = strDebug & vbCrLf & "Fin Regla_CCV" & " >>> " & Regla_CCV

End Function

Private Function Regla_CCC(ult As String, c As String, nextC As String) As Boolean

    'strDebug = strDebug & vbCrLf
    'strDebug = strDebug & vbCrLf & "Inicio Regla_CCC" & " --> " & ult & " - " & c & " - " & nextC

    If EsConsonante(ult) And EsConsonante(c) And EsConsonante(nextC) Then
        Regla_CCC = True
    End If

    'strDebug = strDebug & vbCrLf & "Fin Regla_CCC" & " >>> " & Regla_CCC

End Function

'-----------------------------------------------------------------
' 3.- REGLAS AVANZADAS (FASE 2)
'-----------------------------------------------------------------
' V C + C V
Private Function Regla_VC_CV(ult As String, c As String, nextC As String, next2 As String) As Boolean

    'strDebug = strDebug & vbCrLf
    'strDebug = strDebug & vbCrLf & "Inicio Regla_VC_CV --> " & ult & " - " & c & " - " & nextC & " - " & next2

    ' ORTOGRÁFICO RAE:
    ' V + C + C + V ? se corta ENTRE las dos consonantes
    ' salvo que CC sea grupo inseparable (pr, tr, cl, etc.)

    If EsVocal(ult) And EsConsonante(c) And EsConsonante(nextC) And EsVocal(next2) Then
        
        ' Si el grupo CC es inseparable, NO cortar
        If EsGrupoInseparable(c & nextC) Then Exit Function

        ' En todos los demás casos, cortar
        Regla_VC_CV = True
    End If

    'strDebug = strDebug & vbCrLf & "Fin Regla_VC_CV >>> " & Regla_VC_CV

End Function

' V C + C C V
Private Function Regla_VC_CCV(ult As String, c As String, nextC As String, next2 As String) As Boolean
    ' ult = vocal
    ' c = consonante
    ' nextC = consonante
    ' next2 = vocal

    'strDebug = strDebug & vbCrLf
    'strDebug = strDebug & vbCrLf & "Inicio Regla_VC_CCV" & " --> " & ult & " - " & c & " - " & nextC & " - " & next2

    If EsVocal(ult) And EsConsonante(c) And EsConsonante(nextC) And EsVocal(next2) Then
        Regla_VC_CCV = True
    End If

    'strDebug = strDebug & vbCrLf & "Fin Regla_VC_CCV" & " >>> " & Regla_VC_CCV

End Function

' V C + S C V
Private Function Regla_VC_SCV(ult As String, c As String, nextC As String, next2 As String) As Boolean

    'strDebug = strDebug & vbCrLf
    'strDebug = strDebug & vbCrLf & "Inicio Regla_VC_SCV" & " --> " & ult & " - " & c & " - " & nextC & " - " & next2

    If EsVocal(ult) And EsConsonante(c) And nextC = "s" And EsConsonante(next2) Then
        Regla_VC_SCV = True
    End If

    'strDebug = strDebug & vbCrLf & "Fin Regla_VC_SCV" & " >>> " & Regla_VC_SCV

End Function

' V C + S + C
Private Function Regla_VC_SC(ult As String, c As String, nextC As String) As Boolean

    'strDebug = strDebug & vbCrLf
    'strDebug = strDebug & vbCrLf & "Inicio Regla_VC_SC" & " --> " & ult & " - " & c & " - " & nextC    ' &" - " &  next2

    If EsVocal(ult) And EsConsonante(c) And nextC = "s" Then
        Regla_VC_SC = True
    End If

    'strDebug = strDebug & vbCrLf & "Fin Regla_VC_SC" & " >>> " & Regla_VC_SC

End Function

' V C + S + C + V
Private Function Regla_VC_SCV2(ult As String, c As String, nextC As String, next2 As String) As Boolean

    'strDebug = strDebug & vbCrLf
    'strDebug = strDebug & vbCrLf & "Inicio Regla_VC_SCV2" & " --> " & ult & " - " & c & " - " & nextC & " - " & next2

    If EsVocal(ult) And EsConsonante(c) And nextC = "s" And EsConsonante(next2) Then
        Regla_VC_SCV2 = True
    End If

    'strDebug = strDebug & vbCrLf & "Fin Regla_VC_SCV2" & " >>> " & Regla_VC_SCV2

End Function

'V C + S + C + C
Private Function Regla_VC_SCC(ult As String, c As String, nextC As String, next2 As String) As Boolean

    'strDebug = strDebug & vbCrLf
    'strDebug = strDebug & vbCrLf & "Inicio Regla_VC_SCC" & " --> " & ult & " - " & c & " - " & nextC & " - " & next2

    If EsVocal(ult) And EsConsonante(c) And nextC = "s" And EsConsonante(next2) Then
        Regla_VC_SCC = True
    End If

    'strDebug = strDebug & vbCrLf & "Fin Regla_VC_SCC" & " >>> " & Regla_VC_SCC

End Function

' Prefijos comunes: anti-, intro-, trans-, contra-, extra-, pre-,  pro-, sub-
Private Function Regla_Prefijos(silaba As String, c As String) As Boolean

    'strDebug = strDebug & vbCrLf
    'strDebug = strDebug & vbCrLf & "Inicio Regla_Prefijos" & " --> " & silaba & " - " & c    ' &" - " &  nextC &" - " &  next2

    Dim prefijos
    prefijos = Array("anti", "intro", "trans", "contra", "extra", "pre", "pro", "sub")

    Dim p As Variant
    For Each p In prefijos
        If silaba = p Then
            Regla_Prefijos = True
            Exit Function
        End If
    Next p

    'strDebug = strDebug & vbCrLf & "Fin Regla_Prefijos" & " >>> " & Regla_Prefijos

End Function

Private Function Regla_CortarAntesDeGrupo(ult As String, c As String, nextC As String, next2 As String, next3 As String) As Boolean
    ' Detecta patrones tipo: ci | pria | no
    ' C + líquida (r/l) + i + vocal fuerte (a/e/o)

    ' 1) Grupo consonántico válido
    If InStr("ptbcfgd", c) = 0 Then Exit Function
    If nextC <> "r" And nextC <> "l" Then Exit Function

    ' 2) Diptongo creciente: i + vocal fuerte
    If next2 = "i" And InStr("aeo", next3) > 0 Then
        Regla_CortarAntesDeGrupo = True
    End If
End Function

'Private Function Regla_CortarAntesDeGrupo1(ult As String, c As String, nextC As String, next2 As String, next3 As String) As Boolean
'    ' Detecta patrones tipo: ci | pria | no
'
'    ' 1) Grupo consonántico válido
'    If InStr("ptbcfgd", c) = 0 Then Exit Function
'    If nextC <> "r" And nextC <> "l" Then Exit Function
'
'    ' 2) Diptongo creciente: i + vocal fuerte
'    If next2 = "i" Then
'        ' Necesitamos mirar la vocal siguiente a la "i"
'        ' Pero tu función no recibe next3, así que la regla solo se aplica
'        ' cuando la sílaba actual termina en vocal y next2 = "i"
'        Regla_CortarAntesDeGrupo = True
'    End If
'End Function

'Private Function Regla_CortarAntesDeGrupo(ByVal palabra As String, ByVal pos As Long) As Boolean
'    ' pos = índice donde empieza el grupo consonántico (p, t, b, c, f, g, d)
'    ' Ejemplo: "cipriano"
'    ' c i p r i a n o
'    '     ^ pos = 3 (p)
'
'    Dim consonante As String
'    Dim liquida As String
'    Dim v1 As String
'    Dim v2 As String
'
'    ' Comprobamos que hay suficientes caracteres
'    If pos < 3 Or pos + 2 > Len(palabra) Then Exit Function
'
'    consonante = Mid$(palabra, pos, 1)
'    liquida = Mid$(palabra, pos + 1, 1)
'    v1 = Mid$(palabra, pos + 2, 1)
'    v2 = Mid$(palabra, pos + 3, 1)
'
'    ' 1) Grupo consonántico válido (pr, tr, br, cr, fr, gr, dr)
'    If InStr("ptbcfgd", consonante) = 0 Then Exit Function
'    If liquida <> "r" And liquida <> "l" Then Exit Function
'
'    ' 2) Diptongo creciente: i + vocal fuerte (a, e, o)
'    If v1 = "i" And InStr("aeo", v2) > 0 Then
'        Regla_CortarAntesDeGrupo = True
'    End If
'End Function

'=================================================================
'=================================================================
'       FUNCIONES AUXILIARES GENERALES
'=================================================================
'=================================================================
' ----------------------------------------------------------------
' ============================================================
'   AUXILIARES
' ============================================================
Private Function EsVocal(c As String) As Boolean
    EsVocal = (c Like "[aeiouáéíóú]")
End Function

Private Function EsVocalFuerte(c As String) As Boolean
    EsVocalFuerte = (c Like "[aeoáéó]")
End Function

Private Function EsVocalDebil(c As String) As Boolean
    EsVocalDebil = (c Like "[iuíú]")
End Function

Private Function EsConsonante(c As String) As Boolean
    EsConsonante = (Not EsVocal(c) And c <> " ")
End Function

Private Function EsDiptongo(v1 As String, v2 As String) As Boolean
    If EsVocalDebil(v1) And EsVocalDebil(v2) Then EsDiptongo = True
    If EsVocalFuerte(v1) And EsVocalDebil(v2) Then EsDiptongo = True
    If EsVocalDebil(v1) And EsVocalFuerte(v2) Then EsDiptongo = True
End Function

Private Function EsHiato(v1 As String, v2 As String) As Boolean
    If EsVocalFuerte(v1) And EsVocalFuerte(v2) Then EsHiato = True
    If v1 Like "[íú]" And EsVocalFuerte(v2) Then EsHiato = True
End Function

Private Function EsGrupoInseparable(par As String) As Boolean
    Select Case par
        Case "dr", "tr", "gr", "pr", "pl", "cl", "fr", "fl", "br", "bl"
            EsGrupoInseparable = True
    End Select
End Function

'=================================================================
'=================================================================
'                 SECCIÓN MÓDULO FONÉTICO (ACENTOS)
'=================================================================
'=================================================================

' ============================================================
'   DETECTAR SÍLABAS TÓNICAS
' ============================================================
Private Sub CalcularTonicas()

    Dim tGlobal As New Collection
    Dim elementos As Collection
    Set elementos = ObtenerPalabrasDesdeSilabasAuto()

    Dim globalIndex As Byte
    Dim i As Byte
    
    globalIndex = 0

    For i = 1 To elementos.Count

        If TypeName(elementos(i)) = "Collection" Then
            ' palabra real
            Dim w As Collection
            Set w = elementos(i)

            Dim tLocal As Long
            tLocal = DetectarTonica(w)

            If tLocal > 0 Then
                tGlobal.Add globalIndex + tLocal
            End If

            globalIndex = globalIndex + w.Count

        Else
            ' HUECO
            globalIndex = globalIndex + 1
        End If

    Next i

    ObjDTO.SilabasTonicas = JoinCollection(tGlobal)

End Sub

'Private Sub CalcularTonicas()
'
'    Dim tGlobal As New Collection
'    Dim globalIndex As Long
'    Dim palabras As Collection
'
'    Dim w As Collection
'
'
'    Dim p As Byte
'    Dim tLocal As Byte
'
'    globalIndex = 0
'
'    Set palabras = ObtenerPalabrasDesdeSilabasAuto()
'
'    For p = 1 To palabras.Count
'
'        Set w = palabras(p)
'
'        'If w.Count = 1 And Trim(w(1)) <> "" Then
'
'            Debug.Print "'" & w(1) & "'"
'
'            tLocal = DetectarTonica(w)
'
'            tGlobal.Add globalIndex + tLocal
'        'End If
'
'        globalIndex = globalIndex + w.Count
'    Next p
'
'    ObjDTO.SilabasTonicas= JoinCollection(tGlobal)
'
'End Sub

' ============================================================
'   DETECTAR SÍLABAS SECUNDARIAS (pueden ser varias)
' ============================================================
Private Sub CalcularSecundarias()

    Dim sGlobal As New Collection
    Dim elementos As Collection
    Set elementos = ObtenerPalabrasDesdeSilabasAuto()

    Dim globalIndex As Byte
    Dim i As Byte
    Dim tLocal As Byte
    
    globalIndex = 0
    
    For i = 1 To elementos.Count

        ' Si es una palabra real (Collection)
        If TypeName(elementos(i)) = "Collection" Then

            Dim w As Collection
            Set w = elementos(i)

            ' Detectar tónica local
            tLocal = DetectarTonica(w)

            ' Detectar secundarias locales
            Dim secs As Collection
            Set secs = DetectarSecundarias(w, tLocal)

            ' Convertir secundarias locales ? globales
            Dim x As Variant
            For Each x In secs
                sGlobal.Add globalIndex + CByte(x)
            Next x

            ' Avanzar el índice global por el número de sílabas reales
            globalIndex = globalIndex + w.Count

        Else
            ' Es un hueco ? cuenta como 1 posición global
            globalIndex = globalIndex + 1
        End If

    Next i

    ObjDTO.SilabasSecundarias = JoinCollection(sGlobal)

End Sub

'Private Sub CalcularSecundarias()
'
'    Dim sGlobal As New Collection
'    Dim globalIndex As Long
'    globalIndex = 0
'
'    Dim palabras As Collection
'    Set palabras = ObtenerPalabrasDesdeSilabasAuto()
'
'    Dim p As Long
'    For p = 1 To palabras.Count
'
'        Dim w As Collection
'        Set w = palabras(p)
'
'        Dim secs As Collection
'        Set secs = DetectarSecundarias(w, DetectarTonica(w))
'
'        Dim s As Variant
'        For Each s In secs
'            sGlobal.Add globalIndex + CLng(s)
'        Next s
'
'        globalIndex = globalIndex + w.Count
'    Next p
'
'    ObjDTO.SilabasSecundarias = JoinCollection(sGlobal)
'
'End Sub

' ============================================================
'   MARCAR TÓNICAS Y SECUNDARIAS EN LA CADENA FINAL
' ============================================================
Private Sub MarcarTonicaYSecundariaEnCadena()

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

    ' ============================
    '   1) MARCAR TÓNICAS
    ' ============================
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

    ' ============================
    '   2) MARCAR SECUNDARIAS
    ' ============================
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

    ' ============================
    '   3) UNIR RESULTADO
    ' ============================
    ObjDTO.SilabasAcentuadas = Join(out, " | ")

End Sub


'Private Sub MarcarTonicaYSecundariaEnCadena()
'
'    Dim sils As Variant
'
'    Dim t As Variant
'    Dim s As Variant
'
'    Dim out As String
'    Dim i As Byte, g As Byte
'    Dim marcado As String
'    Dim esT As Boolean: esT = False
'    Dim esS As Boolean: esS = False
'
'    Dim x As Variant
'
'    sils = Split(ObjDTO.SilabasAuto, " | ")
'
'    If ObjDTO.SilabasTonicas<> "" Then t = Split(ObjDTO.SilabasTonicas, ",")
'    If ObjDTO.SilabasSecundarias <> "" Then s = Split(ObjDTO.SilabasSecundarias, ",")
'
'    g = 1
'
'    For i = LBound(sils) To UBound(sils)
'
'        marcado = sils(i)
'
'        If Not IsEmpty(t) Then
'        For Each x In t
'            If CByte(x) = g Then
'                esT = True
'            End If
'        Next x
'        End If
'
'        If Not IsEmpty(s) Then
'            For Each x In s
'                If CByte(x) = g Then esS = True
'            Next x
'        End If
'
'        If esT Then
'            marcado = "( " & marcado & " )"
'        ElseIf esS Then
'            marcado = "[ " & marcado & " ]"
'        Else
'
'        End If
'
'        out = out & marcado
'
'
'        If i < UBound(sils) Then out = out & " | "
'
'        g = g + 1
'    Next i
'
'    ObjDTO.SilabasAcentuadas = out
'
'End Sub


'-------------------------------------------------------------
'             AUXILIAREES TÓNICAS Y SECUNDARIAS
'-------------------------------------------------------------

Private Function DetectarTonica(w As Collection) As Byte

    Dim i As Byte
    Dim ultima As String
    
    For i = 1 To w.Count
        If TieneTilde(w(i)) Then
            DetectarTonica = i
            Exit Function
        End If
    Next i

    ultima = Right$(w(w.Count), 1)

    If ultima Like "[aeiouns]" Then
        DetectarTonica = w.Count - 1
    Else
        DetectarTonica = w.Count
    End If

End Function

Private Function TieneTilde(sil As String) As Boolean
    Dim acentos As String
    Dim i As Byte
    
    
    'acentos = "áéíóúÁÉÍÓÚ"
    acentos = "áéíóú"

    For i = 1 To Len(sil)
        If InStr(acentos, Mid$(sil, i, 1)) > 0 Then
            TieneTilde = True
            Exit Function
        End If
    Next i
End Function

Private Function DetectarSecundarias(w As Collection, tPos As Byte) As Collection

    Dim secs As New Collection
    Dim n As Byte
    Dim pos2 As Byte
    
    n = w.Count

    ' Palabras de 1–3 sílabas ? sin secundaria
    If n < 4 Then
        Set DetectarSecundarias = secs
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

    Set DetectarSecundarias = secs

End Function



' ============================================================
'   OBTENER PALABRAS DESDE SILABAS AUTO
'   Devuelve una Collection donde cada elemento es una palabra
'   y cada palabra es una Collection de sílabas
' ============================================================
Private Function ObtenerPalabrasDesdeSilabasAuto() As Collection

    Dim resultado As New Collection
    Dim palabraActual As New Collection

    Dim sils As Variant
    sils = Split(ObjDTO.SilabasAuto, " | ")

    Dim i As Byte
    For i = LBound(sils) To UBound(sils)

        If Trim$(sils(i)) = "" Then
            ' Es un hueco
            If palabraActual.Count > 0 Then
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
    If palabraActual.Count > 0 Then resultado.Add palabraActual

    Set ObtenerPalabrasDesdeSilabasAuto = resultado

End Function

'Private Function ObtenerPalabrasDesdeSilabasAuto() As Collection
'
'    Dim resultado As New Collection
'    Dim palabraActual As New Collection
'
'    Dim sils As Variant
'    Dim i As Long
'
'    sils = Split(ObjDTO.SilabasAuto, " | ")
'
''    Debug.Print
'
'    For i = LBound(sils) To UBound(sils)
'
'        'Debug.Print "'"; sils(i); "' ==> "; Len(sils(i))
'
'
'        If sils(i) = " " Then
'
'            ' Fin de palabra
'            If palabraActual.Count > 0 Then
'                resultado.Add palabraActual
'                Set palabraActual = New Collection
'
'                palabraActual.Add sils(i)
'                resultado.Add palabraActual
'                Set palabraActual = New Collection
'
'
'            End If
'        Else
'            ' Añadir sílaba a la palabra actual
'            palabraActual.Add sils(i)
'        End If
'
'    Next i
'
'    ' Añadir la última palabra si quedó algo pendiente
'    If palabraActual.Count > 0 Then resultado.Add palabraActual
'
'    Set ObtenerPalabrasDesdeSilabasAuto = resultado
'
''    Dim x As Variant
''    Dim y As Variant
''
''    For Each x In resultado
''        For Each y In x
''            Debug.Print y; ",";
''        Next
''        Debug.Print
''    Next
'
'End Function


' ============================================================
'   JOIN COLLECTION
'   Convierte una Collection en una cadena "1,4,7"
' ============================================================
Private Function JoinCollection(col As Collection) As String

    Dim arr() As String
    Dim i As Byte
    
    If col Is Nothing Then
        JoinCollection = ""
        Exit Function
    End If

    If col.Count = 0 Then
        JoinCollection = ""
        Exit Function
    End If

    ReDim arr(1 To col.Count)

    For i = 1 To col.Count
        arr(i) = CStr(col(i))
    Next i

    JoinCollection = Join(arr, ",")

End Function

