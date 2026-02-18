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
Public Function Entrada_Motor_ES(Texto As String) As String

    strDebug = ""
    
    Set ObjDTO = New clsDTO_Motor

    ObjDTO.TextoOriginal = Texto
    ObjDTO.NormalizaEntrada

    Call Silabear
    Call MF_DebugDTO("Silabear")

    Entrada_Motor_ES = ObjDTO.SilabasAuto

End Function


' ============================================================
'   1.- SILABEO AUTOMÁTICO
' ============================================================
Private Sub Silabear()

    Dim Texto As String
    Dim i As Long
    Dim c As String, prev As String, nextC As String, next2 As String
    Dim silaba As String
    Dim resultado As Collection
    Dim tipo As Byte

    Set resultado = New Collection
    Texto = ObjDTO.TextoNormalizado

    If Len(Texto) = 0 Then
        ObjDTO.SilabasAuto = ""
        Exit Sub
    End If

    silaba = ""

    For i = 1 To Len(Texto)

        c = Mid$(Texto, i, 1)

        prev = ""
        If i > 1 Then prev = Mid$(Texto, i - 1, 1)

        nextC = ""
        If i < Len(Texto) Then nextC = Mid$(Texto, i + 1, 1)

        next2 = ""
        If i < Len(Texto) - 1 Then next2 = Mid$(Texto, i + 2, 1)

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

tipo = DebeCortar(ult, c, nextC, next2, silaba)

If tipo = 1 Then
    ' Corte antes de c
    resultado.Add silaba
    silaba = ""

ElseIf tipo = 2 Then
    ' Corte entre c y nextC
    silaba = silaba & c
    resultado.Add silaba
    silaba = ""
    GoTo siguiente   ' NO añadir c otra vez
End If

'            If DebeCortar(ult, c, nextC, next2, silaba) Then
'                resultado.Add silaba
'                silaba = ""
'            End If
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
Private Function DebeCortar(ult As String, c As String, nextC As String, next2 As String, silaba As String) As Byte

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
If Regla_CCV(ult, c, nextC) Then DebeCortar = 1: Exit Function
If Regla_CCC(ult, c, nextC) Then DebeCortar = 1: Exit Function

' === REGLA VC-CV (corte entre consonantes) ===
If Regla_VC_CV(ult, c, nextC, next2) Then DebeCortar = 2: Exit Function

DebeCortar = 0


'    ' === REGLAS VOCÁLICAS (primero SIEMPRE) ===
'    If Regla_Hiato(ult, c, nextC) Then GoTo cortar
'    If Regla_Triptongo(ult, c, nextC) Then GoTo seguir
'    If Regla_Diptongo(ult, c, nextC) Then GoTo seguir
'    If Regla_VocalTonica(ult, c, nextC) Then GoTo cortar
'
'    ' === PREFIJOS ===
'    If Regla_Prefijos(silaba, c) Then GoTo cortar
'
'    ' === REGLAS CONSONÁNTICAS ===
'    If Regla_NSP_NST(ult, c, nextC) Then GoTo cortar
'    If Regla_SConsonante(ult, c, nextC) Then GoTo cortar
'    If Regla_ClustersS(ult, c, nextC) Then GoTo cortar
'
'    If Regla_VCV(ult, c, nextC) Then GoTo cortar
'    If Regla_CCV(ult, c, nextC) Then GoTo cortar
'    If Regla_CCC(ult, c, nextC) Then GoTo cortar
'
'    ' === REGLAS AVANZADAS ===
'    If Regla_VC_CV(ult, c, nextC, next2) Then GoTo cortar
'    'If Regla_VC_CCV(ult, c, nextC, next2) Then GoTo cortar
'    If Regla_VC_SCV(ult, c, nextC, next2) Then GoTo cortar
'    If Regla_VC_SC(ult, c, nextC) Then GoTo cortar
'    If Regla_VC_SCV2(ult, c, nextC, next2) Then GoTo cortar
'    If Regla_VC_SCC(ult, c, nextC, next2) Then GoTo cortar
'
'seguir:
'    DebeCortar = False
'    Exit Function
'
'cortar:
'    DebeCortar = True

End Function


' ============================================================
'   REGLAS BÁSICAS
' ============================================================
Private Function Regla_Hiato(ult As String, c As String, nextC As String) As Boolean
    strDebug = strDebug & vbCrLf: strDebug = strDebug & vbCrLf & "Inicio Regla_Hiato --> " & ult & " - " & c & " - " & nextC
    If EsVocal(ult) And EsVocal(c) Then
        If EsHiato(ult, c) Then Regla_Hiato = True
    End If
    strDebug = strDebug & vbCrLf & "Fin Regla_Hiato >>> " & Regla_Hiato
End Function

Private Function Regla_Triptongo(ult As String, c As String, nextC As String) As Boolean
    
    strDebug = strDebug & vbCrLf
    strDebug = strDebug & vbCrLf & "Inicio Regla_Triptongo --> " & ult & " - " & c & " - " & nextC
    If EsVocalDebil(ult) And EsVocalFuerte(c) And EsVocalDebil(nextC) Then Regla_Triptongo = True
    strDebug = strDebug & vbCrLf & "Fin Regla_Triptongo >>> " & Regla_Triptongo
    
End Function

Private Function Regla_Diptongo(ult As String, c As String, nextC As String) As Boolean
    
    strDebug = strDebug & vbCrLf
    strDebug = strDebug & vbCrLf & "Inicio Regla_Diptongo --> " & ult & " - " & c & " - " & nextC
    If EsDiptongo(ult, c) Then Regla_Diptongo = True
    strDebug = strDebug & vbCrLf & "Fin Regla_Diptongo >>> " & Regla_Diptongo
    
End Function


' ============================================================
'   REGLA VCV (ORTOGRÁFICA)
' ============================================================
Private Function Regla_VCV(ult As String, c As String, nextC As String) As Boolean

    strDebug = strDebug & vbCrLf
    strDebug = strDebug & vbCrLf & "Inicio Regla_VCV --> " & ult & " - " & c & " - " & nextC

    ' ORTOGRÁFICO: VCV SIEMPRE CORTA
    ' salvo que la consonante forme grupo inseparable (pr, tr, cl, etc.)
    If EsGrupoInseparable(c & nextC) Then Exit Function

    If EsVocal(ult) And EsConsonante(c) And EsVocal(nextC) Then
        Regla_VCV = True
    End If

    strDebug = strDebug & vbCrLf & "Fin Regla_VCV >>> " & Regla_VCV


End Function

'Private Function Regla_VCV(ult As String, c As String, nextC As String) As Boolean
'
'    strDebug = strDebug & vbCrLf
'    strDebug = strDebug & vbCrLf & "Inicio Regla_VCV --> " & ult & " - " & c & " - " & nextC
'
'    If nextC Like "[áéíóú]" Then Exit Function
'    If EsHiato(ult, nextC) Then Exit Function
'    If EsDiptongo(ult, nextC) Then Exit Function
'    If EsGrupoInseparable(c & nextC) Then Exit Function
'
'    If EsVocal(ult) And EsConsonante(c) And EsVocal(nextC) Then Regla_VCV = True
'
'    strDebug = strDebug & vbCrLf & "Fin Regla_VCV >>> " & Regla_VCV
'
'End Function


' ============================================================
'   REGLA VOCAL TÓNICA (ORTOGRÁFICA)
' ============================================================
Private Function Regla_VocalTonica(ult As String, c As String, nextC As String) As Boolean

    strDebug = strDebug & vbCrLf
    strDebug = strDebug & vbCrLf & "Inicio Regla_VocalTonica --> " & ult & " - " & c & " - " & nextC

    ' De momento, para silabeo ortográfico, NO usamos esta regla
    Regla_VocalTonica = False

    strDebug = strDebug & vbCrLf & "Fin Regla_VocalTonica >>> " & Regla_VocalTonica

End Function


'Private Function Regla_VocalTonica(ult As String, c As String, nextC As String) As Boolean
'
'    strDebug = strDebug & vbCrLf
'    strDebug = strDebug & vbCrLf & "Inicio Regla_VocalTonica --> " & ult & " - " & c & " - " & nextC
'
'    ' No cortar si forma grupo inseparable
'    If EsGrupoInseparable(ult & c) Then Exit Function
'
'    ' ORTOGRÁFICO: la r NO se une a la vocal siguiente
'    If ult = "r" Then Exit Function
'
'    ' ORTOGRÁFICO: NO cortar si estamos en un patrón VCV
'    ' (porque VCV ya decide el corte)
'    If EsVocal(ult) And EsConsonante(c) And EsVocal(nextC) Then Exit Function
'
'    ' Cortar si consonante + vocal tónica
'    If EsConsonante(ult) And c Like "[áéíóú]" Then
'        Regla_VocalTonica = True
'    End If
'
'    strDebug = strDebug & vbCrLf & "Fin Regla_VocalTonica >>> " & Regla_VocalTonica
'
'End Function

'Private Function Regla_VocalTonica(ult As String, c As String, nextC As String) As Boolean
'
'    strDebug = strDebug & vbCrLf
'    strDebug = strDebug & vbCrLf & "Inicio Regla_VocalTonica --> " & ult & " - " & c & " - " & nextC
'
'    ' No cortar si forma grupo inseparable
'    If EsGrupoInseparable(ult & c) Then Exit Function
'
'    ' ORTOGRÁFICO: la r NO se une a la vocal siguiente
'    If ult = "r" Then Exit Function
'
'    ' ORTOGRÁFICO: NO cortar si estamos en un patrón VCV
'    ' (porque VCV ya decide el corte)
'    If EsVocal(ult) And EsConsonante(c) And EsVocal(nextC) Then Exit Function
'
'    ' Cortar si consonante + vocal tónica
'    If EsConsonante(ult) And c Like "[áéíóú]" Then
'        Regla_VocalTonica = True
'    End If
'
'    strDebug = strDebug & vbCrLf & "Fin Regla_VocalTonica >>> " & Regla_VocalTonica
'
'End Function

'Private Function Regla_VocalTonica(ult As String, c As String, nextC As String) As Boolean
'
'    strDebug = strDebug & vbCrLf
'    strDebug = strDebug & vbCrLf & "Inicio Regla_VocalTonica --> " & ult & " - " & c & " - " & nextC
'
'    ' No cortar si ult forma parte de un ataque complejo
'    If EsGrupoInseparable(ult & c) Then Exit Function
'
'    ' ORTOGRÁFICO: la r NO se une a la vocal siguiente
'    If ult = "r" Then Exit Function
'
'    ' Cortar si consonante + vocal tónica
'    If EsConsonante(ult) And c Like "[áéíóú]" Then
'        Regla_VocalTonica = True
'    End If
'
'    strDebug = strDebug & vbCrLf & "Fin Regla_VocalTonica >>> " & Regla_VocalTonica
'
'End Function

'Private Function Regla_VocalTonica(ult As String, c As String, nextC As String) As Boolean
'
'    strDebug = strDebug & vbCrLf
'    strDebug = strDebug & vbCrLf & "Inicio Regla_VocalTonica --> " & ult & " - " & c & " - " & nextC
'
'    If EsGrupoInseparable(ult & c) Then Exit Function
'    If ult = "r" Then Exit Function
'
'    If EsConsonante(ult) And c Like "[áéíóú]" Then Regla_VocalTonica = True
'
'    strDebug = strDebug & vbCrLf & "Fin Regla_VocalTonica >>> " & Regla_VocalTonica
'
'End Function


' ============================================================
'   REGLAS CONSONÁNTICAS Y AVANZADAS
' ============================================================
' (Aquí van todas tus reglas NSP/NST, SConsonante, ClustersS,
'  CCV, CCC, VC_CCV, VC_SCV, VC_SC, VC_SCV2, VC_SCC)
'  — Las dejo tal cual las tienes, solo con strDebug = strDebug & vbcrlf  activado.

'-----------------------------------------------------------------
' 2.- REGLAS AVANZADAS (FASE 1)
'-----------------------------------------------------------------

Private Function Regla_NSP_NST(ult As String, c As String, nextC As String) As Boolean

    strDebug = strDebug & vbCrLf
    strDebug = strDebug & vbCrLf & "Inicio Regla_NSP_NST" & " --> " & ult & " - " & c & " - " & nextC

    If ult = "n" And c = "s" Then
        If nextC = "p" Or nextC = "t" Then
            Regla_NSP_NST = True
        End If
    End If

    strDebug = strDebug & vbCrLf & "Fin Regla_NSP_NST" & " >>> " & Regla_NSP_NST

End Function

Private Function Regla_SConsonante(ult As String, c As String, nextC As String) As Boolean

    strDebug = strDebug & vbCrLf
    strDebug = strDebug & vbCrLf & "Inicio Regla_SConsonante" & " --> " & ult & " - " & c & " - " & nextC

    If ult = "s" And EsConsonante(c) Then
        If EsVocal(nextC) Then
            Regla_SConsonante = True
        End If
    End If

    strDebug = strDebug & vbCrLf & "Fin Regla_SConsonante" & " >>> " & Regla_SConsonante

End Function

Private Function Regla_ClustersS(ult As String, c As String, nextC As String) As Boolean

    strDebug = strDebug & vbCrLf
    strDebug = strDebug & vbCrLf & "Inicio Regla_ClustersS" & " --> " & ult & " - " & c & " - " & nextC

    Dim par As String
    par = c & nextC

    If ult = "s" Then
        Select Case par
            Case "tr", "pr", "pl", "cr", "cl", "gr", "fr"
                Regla_ClustersS = True
        End Select
    End If

    strDebug = strDebug & vbCrLf & "Fin Regla_ClustersS" & " >>> " & Regla_ClustersS

End Function

Private Function Regla_CCV(ult As String, c As String, nextC As String) As Boolean

    strDebug = strDebug & vbCrLf
    strDebug = strDebug & vbCrLf & "Inicio Regla_CCV" & " --> " & ult & " - " & c & " - " & nextC

    If EsConsonante(ult) And EsConsonante(c) And EsVocal(nextC) Then
        If Not EsGrupoInseparable(ult & c) Then
            Regla_CCV = True
        End If
    End If

    strDebug = strDebug & vbCrLf & "Fin Regla_CCV" & " >>> " & Regla_CCV

End Function

Private Function Regla_CCC(ult As String, c As String, nextC As String) As Boolean

    strDebug = strDebug & vbCrLf
    strDebug = strDebug & vbCrLf & "Inicio Regla_CCC" & " --> " & ult & " - " & c & " - " & nextC

    If EsConsonante(ult) And EsConsonante(c) And EsConsonante(nextC) Then
        Regla_CCC = True
    End If

    strDebug = strDebug & vbCrLf & "Fin Regla_CCC" & " >>> " & Regla_CCC

End Function

'-----------------------------------------------------------------
' 3.- REGLAS AVANZADAS (FASE 2)
'-----------------------------------------------------------------
' V C + C V
Private Function Regla_VC_CV(ult As String, c As String, nextC As String, next2 As String) As Boolean

    strDebug = strDebug & vbCrLf
    strDebug = strDebug & vbCrLf & "Inicio Regla_VC_CV --> " & ult & " - " & c & " - " & nextC & " - " & next2

    ' ORTOGRÁFICO RAE:
    ' V + C + C + V ? se corta ENTRE las dos consonantes
    ' salvo que CC sea grupo inseparable (pr, tr, cl, etc.)

    If EsVocal(ult) And EsConsonante(c) And EsConsonante(nextC) And EsVocal(next2) Then
        
        ' Si el grupo CC es inseparable, NO cortar
        If EsGrupoInseparable(c & nextC) Then Exit Function

        ' En todos los demás casos, cortar
        Regla_VC_CV = True
    End If

    strDebug = strDebug & vbCrLf & "Fin Regla_VC_CV >>> " & Regla_VC_CV

End Function

'' V C + C V
'Private Function Regla_VC_CV(ult As String, c As String, nextC As String, next2 As String) As Boolean
'
'    strDebug = strDebug & vbCrLf
'    strDebug = strDebug & vbCrLf & "Inicio Regla_VC_CV --> " & ult & " - " & c & " - " & nextC & " - " & next2
'
'    ' ORTOGRÁFICO: V C + C V corta entre las dos consonantes
'    If EsVocal(ult) And EsConsonante(c) And EsConsonante(nextC) And EsVocal(next2) Then
'
'        ' Si el grupo CC es inseparable (pr, tr, cl, etc.), NO cortar
'        If EsGrupoInseparable(c & nextC) Then Exit Function
'
'        ' En todos los demás casos, cortar
'        Regla_VC_CV = True
'    End If
'
'    strDebug = strDebug & vbCrLf & "Fin Regla_VC_CV >>> " & Regla_VC_CV
'
'End Function

'' V C + C C V
'Private Function Regla_VC_CCV(ult As String, c As String, nextC As String, next2 As String) As Boolean
'    ' ult = vocal
'    ' c = consonante
'    ' nextC = consonante
'    ' next2 = vocal
'
'    strDebug = strDebug & vbCrLf
'    strDebug = strDebug & vbCrLf & "Inicio Regla_VC_CCV" & " --> " & ult & " - " & c & " - " & nextC & " - " & next2
'
'    If EsVocal(ult) And EsConsonante(c) And EsConsonante(nextC) And EsVocal(next2) Then
'        Regla_VC_CCV = True
'    End If
'
'    strDebug = strDebug & vbCrLf & "Fin Regla_VC_CCV" & " >>> " & Regla_VC_CCV
'
'End Function

' V C + S C V
Private Function Regla_VC_SCV(ult As String, c As String, nextC As String, next2 As String) As Boolean

    strDebug = strDebug & vbCrLf
    strDebug = strDebug & vbCrLf & "Inicio Regla_VC_SCV" & " --> " & ult & " - " & c & " - " & nextC & " - " & next2

    If EsVocal(ult) And EsConsonante(c) And nextC = "s" And EsConsonante(next2) Then
        Regla_VC_SCV = True
    End If

    strDebug = strDebug & vbCrLf & "Fin Regla_VC_SCV" & " >>> " & Regla_VC_SCV

End Function

' V C + S + C
Private Function Regla_VC_SC(ult As String, c As String, nextC As String) As Boolean

    strDebug = strDebug & vbCrLf
    strDebug = strDebug & vbCrLf & "Inicio Regla_VC_SC" & " --> " & ult & " - " & c & " - " & nextC    ' &" - " &  next2

    If EsVocal(ult) And EsConsonante(c) And nextC = "s" Then
        Regla_VC_SC = True
    End If

    strDebug = strDebug & vbCrLf & "Fin Regla_VC_SC" & " >>> " & Regla_VC_SC

End Function

' V C + S + C + V
Private Function Regla_VC_SCV2(ult As String, c As String, nextC As String, next2 As String) As Boolean

    strDebug = strDebug & vbCrLf
    strDebug = strDebug & vbCrLf & "Inicio Regla_VC_SCV2" & " --> " & ult & " - " & c & " - " & nextC & " - " & next2

    If EsVocal(ult) And EsConsonante(c) And nextC = "s" And EsConsonante(next2) Then
        Regla_VC_SCV2 = True
    End If

    strDebug = strDebug & vbCrLf & "Fin Regla_VC_SCV2" & " >>> " & Regla_VC_SCV2

End Function

'V C + S + C + C
Private Function Regla_VC_SCC(ult As String, c As String, nextC As String, next2 As String) As Boolean

    strDebug = strDebug & vbCrLf
    strDebug = strDebug & vbCrLf & "Inicio Regla_VC_SCC" & " --> " & ult & " - " & c & " - " & nextC & " - " & next2

    If EsVocal(ult) And EsConsonante(c) And nextC = "s" And EsConsonante(next2) Then
        Regla_VC_SCC = True
    End If

    strDebug = strDebug & vbCrLf & "Fin Regla_VC_SCC" & " >>> " & Regla_VC_SCC

End Function

' Prefijos comunes: anti-, intro-, trans-, contra-, extra-, pre-,  pro-, sub-
Private Function Regla_Prefijos(silaba As String, c As String) As Boolean

    strDebug = strDebug & vbCrLf
    strDebug = strDebug & vbCrLf & "Inicio Regla_Prefijos" & " --> " & silaba & " - " & c    ' &" - " &  nextC &" - " &  next2

    Dim prefijos
    prefijos = Array("anti", "intro", "trans", "contra", "extra", "pre", "pro", "sub")

    Dim p As Variant
    For Each p In prefijos
        If silaba = p Then
            Regla_Prefijos = True
            Exit Function
        End If
    Next p

    strDebug = strDebug & vbCrLf & "Fin Regla_Prefijos" & " >>> " & Regla_Prefijos

End Function






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


