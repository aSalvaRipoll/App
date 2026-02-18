Attribute VB_Name = "bas_Motor_ES_Main"

' ============================================================
' Nombre:    bas-Motor_ES_Main
' Tipo:      Módulo
' Propósito: Motor de silabeo ortográfico del español (Versión 4)
' Autor:     Alba Salvá
' Fecha:     11/02/2026
' Versión:   4 (Silabeo depurado + reparación de sílabas)
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

    Entrada_Motor_ES = ObjDTO.FonemasFinal

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
                    resultado.Add silaba
                    silaba = ""

                Case 2
                    silaba = silaba & c
                    resultado.Add silaba
                    silaba = ""
                    GoTo siguiente
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
'   REGLAS DE SILABEO (ORDENADAS)
' ============================================================
Private Function DebeCortar(ult As String, c As String, nextC As String, next2 As String, silaba As String, next3 As String) As Integer
    ' 0 = no cortar
    ' 1 = cortar antes de c
    ' 2 = cortar entre c y nextC

    ' --- 1) HIATOS ---
    If EsVocal(ult) And EsVocal(c) Then
        If EsHiato(ult, c) Then
            DebeCortar = 1
            Exit Function
        End If
    End If

    ' --- 2) V C V ---
    If EsVocal(ult) And EsConsonante(c) And EsVocal(nextC) Then
        If EsAtaqueValido(c, nextC) Then
            DebeCortar = 0
        Else
            DebeCortar = 1
        End If
        Exit Function
    End If

    ' --- 3) V C C V ---
    If EsVocal(ult) And EsConsonante(c) And EsConsonante(nextC) And EsVocal(next2) Then
        If EsAtaqueValido(nextC, next2) Then
            DebeCortar = 1
        Else
            DebeCortar = 2
        End If
        Exit Function
    End If

    DebeCortar = 0
End Function

'Private Function DebeCortar(ult As String, c As String, nextC As String, next2 As String, _
'                            silaba As String, next3 As String) As Byte
'
'    ' === REGLAS VOCÁLICAS ===
'    If Regla_Hiato(ult, c, nextC) Then DebeCortar = 1: Exit Function
'    If Regla_Triptongo(ult, c, nextC) Then DebeCortar = 0: Exit Function
'    If Regla_Diptongo(ult, c, nextC) Then DebeCortar = 0: Exit Function
'    If Regla_VocalTonica(ult, c, nextC) Then DebeCortar = 1: Exit Function
'
'    ' === PREFIJOS ===
'    If Regla_Prefijos(silaba, c) Then DebeCortar = 1: Exit Function
'
'    ' === REGLAS CONSONÁNTICAS BÁSICAS ===
'    If Regla_NSP_NST(ult, c, nextC) Then DebeCortar = 1: Exit Function
'    If Regla_SConsonante(ult, c, nextC) Then DebeCortar = 1: Exit Function
'    If Regla_ClustersS(ult, c, nextC) Then DebeCortar = 1: Exit Function
'
'    ' === VCV (ORTOGRÁFICO) ===
'    If Regla_VCV(ult, c, nextC) Then DebeCortar = 1: Exit Function
'
'    ' === REGLAS CONSONÁNTICAS ESPECIALES ===
'    If Regla_CortarAntesDeGrupo(ult, c, nextC, next2, next3) Then DebeCortar = 1: Exit Function
'
'    ' === VC + C V (RAE) ===
'    If Regla_VC_CV(ult, c, nextC, next2) Then DebeCortar = 2: Exit Function
'
'    ' === CCV / CCC ===
'    If Regla_CCV(ult, c, nextC) Then DebeCortar = 1: Exit Function
'    If Regla_CCC(ult, c, nextC) Then DebeCortar = 1: Exit Function
'
'    DebeCortar = 0
'
'End Function


' ============================================================
'   REGLAS BÁSICAS VOCÁLICAS
' ============================================================
Private Function Regla_Hiato(ult As String, c As String, nextC As String) As Boolean

    If EsVocal(ult) And EsVocal(c) Then
        If EsHiato(ult, c) Then Regla_Hiato = True
    End If

End Function

Private Function Regla_Triptongo(ult As String, c As String, nextC As String) As Boolean

    If EsVocalDebil(ult) And EsVocalFuerte(c) And EsVocalDebil(nextC) Then
        Regla_Triptongo = True
    End If

End Function

Private Function Regla_Diptongo(ult As String, c As String, nextC As String) As Boolean

    If EsDiptongo(ult, c) Then Regla_Diptongo = True

End Function


' ============================================================
'   REGLA VCV (ORTOGRÁFICA)
' ============================================================
Private Function Regla_VCV(ult As String, c As String, nextC As String) As Boolean

    ' VCV corta salvo:
    ' - si la consonante forma grupo inseparable con la siguiente
    ' - si la consonante es "h" (no fonética)
    ' - si la consonante es "y" y actúa como semivocal (y + vocal)

    If c = "h" Then Exit Function
    If c = "y" And EsVocal(nextC) Then Exit Function
    If EsGrupoInseparable(c & nextC) Then Exit Function

    If EsVocal(ult) And EsConsonante(c) And EsVocal(nextC) Then
        Regla_VCV = True
    End If

End Function


' ============================================================
'   REGLA VOCAL TÓNICA (RESERVADA)
' ============================================================
Private Function Regla_VocalTonica(ult As String, c As String, nextC As String) As Boolean

    ' Reservada para silabeo prosódico; no se usa en ortográfico
    Regla_VocalTonica = False

End Function


' ============================================================
'   REGLAS CONSONÁNTICAS Y AVANZADAS
' ============================================================
'-----------------------------------------------------------------
' 2.- REGLAS AVANZADAS (FASE 1)
'-----------------------------------------------------------------
Private Function Regla_NSP_NST(ult As String, c As String, nextC As String) As Boolean

    If ult = "n" And c = "s" Then
        If nextC = "p" Or nextC = "t" Then
            Regla_NSP_NST = True
        End If
    End If

End Function

Private Function Regla_SConsonante(ult As String, c As String, nextC As String) As Boolean

    If ult = "s" And EsConsonante(c) Then
        If EsVocal(nextC) Then
            Regla_SConsonante = True
        End If
    End If

End Function

Private Function Regla_ClustersS(ult As String, c As String, nextC As String) As Boolean

    Dim par As String
    par = c & nextC

    If ult = "s" Then
        Select Case par
            Case "tr", "pr", "pl", "cr", "cl", "gr", "fr"
                Regla_ClustersS = True
        End Select
    End If

End Function

Private Function Regla_CCV(ult As String, c As String, nextC As String) As Boolean

    If EsConsonante(ult) And EsConsonante(c) And EsVocal(nextC) Then
        If Not EsGrupoInseparable(ult & c) Then
            Regla_CCV = True
        End If
    End If

End Function

Private Function Regla_CCC(ult As String, c As String, nextC As String) As Boolean

    ' CCC corta salvo que c & nextC formen grupo inseparable
    If EsConsonante(ult) And EsConsonante(c) And EsConsonante(nextC) Then
        If Not EsGrupoInseparable(c & nextC) Then
            Regla_CCC = True
        End If
    End If

End Function


'-----------------------------------------------------------------
' 3.- REGLAS AVANZADAS (FASE 2)
'-----------------------------------------------------------------
' V C + C V
Private Function Regla_VC_CV(ult As String, c As String, nextC As String, next2 As String) As Boolean

    ' ORTOGRÁFICO RAE:
    ' V + C + C + V ? se corta ENTRE las dos consonantes
    ' salvo que CC sea grupo inseparable (pr, tr, cl, etc.)

    If EsVocal(ult) And EsConsonante(c) And EsConsonante(nextC) And EsVocal(next2) Then

        ' Si el grupo CC es inseparable, NO cortar
        If EsGrupoInseparable(c & nextC) Then Exit Function

        ' En todos los demás casos, cortar
        Regla_VC_CV = True
    End If

End Function

' V C + C C V
Private Function Regla_VC_CCV(ult As String, c As String, nextC As String, next2 As String) As Boolean

    If EsVocal(ult) And EsConsonante(c) And EsConsonante(nextC) And EsVocal(next2) Then
        Regla_VC_CCV = True
    End If

End Function

' V C + S C V
Private Function Regla_VC_SCV(ult As String, c As String, nextC As String, next2 As String) As Boolean

    If EsVocal(ult) And EsConsonante(c) And nextC = "s" And EsConsonante(next2) Then
        Regla_VC_SCV = True
    End If

End Function

' V C + S + C
Private Function Regla_VC_SC(ult As String, c As String, nextC As String) As Boolean

    If EsVocal(ult) And EsConsonante(c) And nextC = "s" Then
        Regla_VC_SC = True
    End If

End Function

' V C + S + C + V
Private Function Regla_VC_SCV2(ult As String, c As String, nextC As String, next2 As String) As Boolean

    If EsVocal(ult) And EsConsonante(c) And nextC = "s" And EsConsonante(next2) Then
        Regla_VC_SCV2 = True
    End If

End Function

' V C + S + C + C
Private Function Regla_VC_SCC(ult As String, c As String, nextC As String, next2 As String) As Boolean

    If EsVocal(ult) And EsConsonante(c) And nextC = "s" And EsConsonante(next2) Then
        Regla_VC_SCC = True
    End If

End Function

' Prefijos comunes: anti-, intro-, trans-, contra-, extra-, pre-, pro-, sub-
Private Function Regla_Prefijos(silaba As String, c As String) As Boolean

    Dim prefijos
    prefijos = Array("anti", "intro", "trans", "contra", "extra", "pre", "pro", "sub")

    Dim p As Variant
    For Each p In prefijos
        If silaba = p Then
            Regla_Prefijos = True
            Exit Function
        End If
    Next p

End Function

Private Function Regla_CortarAntesDeGrupo(ult As String, c As String, nextC As String, _
                                          next2 As String, next3 As String) As Boolean
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


'=================================================================
'=================================================================
'       FUNCIONES AUXILIARES GENERALES
'=================================================================
'=================================================================
Private Function EsVocal(c As String) As Boolean
    EsVocal = InStr("aeiouáéíóúü", LCase$(c)) > 0
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
    Dim abiertas As String: abiertas = "aáeéoó"
    Dim cerradasAcent As String: cerradasAcent = "íú"

    If InStr(abiertas, v1) > 0 And InStr(abiertas, v2) > 0 Then
        EsHiato = True
    ElseIf InStr(abiertas, v1) > 0 And InStr(cerradasAcent, v2) > 0 Then
        EsHiato = True
    ElseIf InStr(cerradasAcent, v1) > 0 And InStr(abiertas, v2) > 0 Then
        EsHiato = True
    Else
        EsHiato = False
    End If
End Function

Private Function EsGrupoInseparable(par As String) As Boolean

    Select Case par
        Case "dr", "tr", "gr", "pr", "pl", "cl", "fr", "fl", "br", "bl"
            EsGrupoInseparable = True
    End Select

End Function

Private Function EsAtaqueValido(c1 As String, c2 As String) As Boolean
    Dim ataques As Variant
    ataques = Array("br", "bl", "cr", "cl", "dr", "fr", "fl", "gr", "gl", "pr", "pl", "tr")
    EsAtaqueValido = (UBound(Filter(ataques, LCase$(c1 & c2))) >= 0)
End Function

' ============================================================
'   REPARACIÓN DE SÍLABAS IMPOSIBLES
' ============================================================
Private Sub RepararSilabas(resultado As Collection)

    Dim i As Long
    Dim s As String

    i = 1
    Do While i <= resultado.Count

        s = Trim$(resultado(i))

        ' 1) Consonante sola ? unir con la siguiente sílaba real
        If Len(s) = 1 Then
            If EsConsonante(s) Then

                ' Si hay siguiente sílaba real
                If i < resultado.Count Then
                    If Trim$(resultado(i + 1)) <> "" Then
                        resultado(i + 1) = s & resultado(i + 1)
                        resultado.Remove i
                        GoTo siguienteIter
                    End If
                End If

            End If
        End If

        i = i + 1
siguienteIter:
    Loop

End Sub


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

        If TypeName(elementos(i)) = "Collection" Then

            Dim w As Collection
            Set w = elementos(i)

            ' Detectar tónica local
            tLocal = DetectarTonica(w)

            ' Detectar secundarias locales
            Dim secs As Collection
            Set secs = DetectarSecundarias(w, tLocal)

            ' Convertir secundarias locales a globales
            Dim x As Variant
            For Each x In secs
                sGlobal.Add globalIndex + CByte(x)
            Next x

            globalIndex = globalIndex + w.Count

        Else
            ' HUECO
            globalIndex = globalIndex + 1
        End If

    Next i

    ObjDTO.SilabasSecundarias = JoinCollection(sGlobal)

End Sub


' ============================================================
'   MARCAR TÓNICAS Y SECUNDARIAS EN LA CADENA FINAL
' ============================================================
Private Sub MarcarTonicaYSecundariaEnCadena()

    Dim silabas() As String
    Dim i As Long
    Dim tonica As Long

    ' Dividir sílabas auto en array
    silabas = Split(ObjDTO.SilabasAuto, " | ")

    ' Detectar sílaba tónica (la que contiene vocal acentuada)
    tonica = 0
    For i = LBound(silabas) To UBound(silabas)
        If TieneTilde(silabas(i)) Then
            tonica = i + 1
            Exit For
        End If
    Next i

    ObjDTO.SilabasTonicas = tonica
    ObjDTO.SilabasSecundarias = ""

    ' Construir SilabasAcentuadas SIN modificar las sílabas
    Dim out As String
    out = ""

    For i = LBound(silabas) To UBound(silabas)
        If (i + 1) = tonica Then
            out = out & "( " & silabas(i) & " )"
        Else
            out = out & silabas(i)
        End If

        If i < UBound(silabas) Then out = out & " | "
    Next i

    ObjDTO.SilabasAcentuadas = out

End Sub


'-------------------------------------------------------------
'             AUXILIARES TÓNICAS Y SECUNDARIAS
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


' ============================================================
'   JOIN COLLECTION
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


