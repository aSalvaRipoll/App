Attribute VB_Name = "bas_Motor_IB_Main"

' ============================================================
'   MOTOR FONÈTIC — ILLAS BALEARS (MÒDUL ORTOGRÀFIC)
'   Arquitectura idèntica al motor català
' ============================================================

Option Compare Database
Option Explicit

Private usarSilabeoMorfologico As Boolean
Private modoPrefijosEstrictos As Boolean
Private respetarPrefijos As Boolean

Private prefijosEstrictos_IB As Variant
Private prefijosCargados_IB As Boolean

' ============================================================
'   ENTRADA PRINCIPAL DEL MOTOR (ILLAS BALEARS)
' ============================================================
' ----------------------------------------------------------------
' Procedimiento: Entrada_Motor_IB
' Propósito:     Punto de entrada al motor fonético del Mallorquín
' Tipo proc.:    Function
' Acceso proc.:  Public

' Parameter Texto (String): Texto que se recibe (nombre o apellido)

' Tipo retorno: String -> Texto que contiene la lista de fonemas
'   resultado de la conversión

' Autor:        Alba Salvá
' Fecha:        16/02/2026
' ----------------------------------------------------------------

' ============================================================
'   ENTRADA PRINCIPAL DEL MOTOR IB
' ============================================================
Public Function Entrada_Motor_IB(texto As String) As String

    Set ObjDTO = New clsDTO_Motor

    ' 1) Normalització general (DTO)
    ObjDTO.TextoOriginal = texto
    ObjDTO.NormalizaEntrada

    ' 2) Silabeo automàtic
    Call SilabearFrase_IB

    ' 3) Detectar tònica
    Call CalcularTonicas_IB

    ' 4) Detectar secundàries
    Call CalcularSecundarias_IB

    ' 5) Marcar tònica i secundàries
    Call MarcarTonicaYSecundariaEnCadena_IB

    ObjDTO.SilabasFinal = ObjDTO.SilabasAcentuadas

    ' 6) Fonètica
    Call ConstruirCadenaFonemas_IB

    ' 7) Retorn (igual que català)
    Entrada_Motor_IB = ObjDTO.SilabasAuto

End Function

' ============================================================
'   SILABEO DE FRASE
' ============================================================
Private Sub SilabearFrase_IB()

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
            sil = SilabearPalabra_IB(limpia)
            If resultado = "" Then
                resultado = sil
            Else
                resultado = resultado & " |   | " & sil
            End If
        End If
    Next i

    ObjDTO.SilabasAuto = resultado

End Sub

' ============================================================
'   SILABEO DE PALABRA
' ============================================================
Public Function SilabearPalabra_IB(ByVal texto As String) As String
    Dim t As String

    t = LCase$(Trim$(texto))

    If usarSilabeoMorfologico Then
        SilabearPalabra_IB = SilabearMorfologico_IB(t)
    Else
        SilabearPalabra_IB = SilabearOrtog_IB(t)
    End If

End Function

' ============================================================
'   SILABEO ORTOGRÁFICO IB
' ============================================================
Public Function SilabearOrtog_IB(ByVal t As String) As String
    Dim s As String
    Dim nucIni() As Integer, nucFin() As Integer
    Dim silIni() As Integer, silFin() As Integer
    Dim silabas() As String
    Dim nNuc As Integer
    Dim i As Integer

    ' 1. Normalización ortográfica IB
    s = Normalizar_IB(t)

    ' 2. Arrays para núcleos (máx. 200)
    ReDim nucIni(1 To 200)
    ReDim nucFin(1 To 200)

    ' 3. Localizar núcleos vocálicos IB
    LocalizarNucleosOrtog_IB s, nucIni, nucFin, nNuc

    ' 4. Arrays para sílabas
    ReDim silIni(1 To nNuc)
    ReDim silFin(1 To nNuc)

    ' 5. Calcular sílabas IB
    CalcularSilabas_IB s, nucIni, nucFin, nNuc, silIni, silFin

    ' 6. Construir sílabas
    ReDim silabas(1 To nNuc)
    For i = 1 To nNuc
        silabas(i) = Mid$(s, silIni(i), silFin(i) - silIni(i) + 1)
    Next i

    ' 7. Retornar sílabas separadas por " | "
    SilabearOrtog_IB = Join(silabas, " | ")
End Function

' ============================================================
'   LOCALIZAR NÚCLEOS VOCÁLICOS IB
' ============================================================
Sub LocalizarNucleosOrtog_IB(ByVal t As String, _
                             ByRef nucIni() As Integer, _
                             ByRef nucFin() As Integer, _
                             ByRef nNuc As Integer)

    Dim i As Integer, L As Integer
    Dim c1 As String, c2 As String

    L = Len(t)
    nNuc = 0
    i = 1

    Do While i <= L

        c1 = Mid$(t, i, 1)

        If EsVocal_IB(c1) Then

            ' ¿Diptongo IB?
            If i < L Then
                c2 = Mid$(t, i + 1, 1)

                If EsVocal_IB(c2) Then
                    If EsDiptong_IB(c1, c2) Then
                        nNuc = nNuc + 1
                        nucIni(nNuc) = i
                        nucFin(nNuc) = i + 1
                        i = i + 2
                        GoTo Siguiente
                    End If
                End If
            End If

            ' Vocal sola
            nNuc = nNuc + 1
            nucIni(nNuc) = i
            nucFin(nNuc) = i
            i = i + 1
            GoTo Siguiente

        End If

        i = i + 1

Siguiente:
    Loop

End Sub

' ============================================================
'   CÁLCULO DE SÍLABAS IB
' ============================================================
Sub CalcularSilabas_IB(ByVal t As String, _
                       ByRef nucIni() As Integer, _
                       ByRef nucFin() As Integer, _
                       ByVal nNuc As Integer, _
                       ByRef silIni() As Integer, _
                       ByRef silFin() As Integer)

    Dim i As Integer, L As Integer
    Dim a As Integer, b As Integer
    Dim k As Integer
    Dim c1 As String, c2 As String, c3 As String, c4 As String
    Dim grup As String

    L = Len(t)
    silIni(1) = 1

    For i = 1 To nNuc - 1

        a = nucFin(i)
        b = nucIni(i + 1)

        k = IIf(b > a + 1, b - a - 1, 0)

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
                grup = c1 & c2

                ' Dígrafos indivisibles IB
                If grup = "rr" Or grup = "ll" Or grup = "ch" Then
                    silFin(i) = a
                    silIni(i + 1) = a + 1
                    GoTo Siguiente
                End If

                ' Grupos de ataque IB
                If EsGrupoAtaque_IB(grup) Then
                    silFin(i) = a
                    silIni(i + 1) = a + 1
                Else
                    silFin(i) = a + 1
                    silIni(i + 1) = a + 2
                End If

            Case 3
                c1 = Mid$(t, a + 1, 1)
                c2 = Mid$(t, a + 2, 1)
                c3 = Mid$(t, a + 3, 1)

                ' REGLA IB: L·L + vocal es indivisible
                If c1 = "l" And c2 = "·" And c3 = "l" Then
                    If b <= L Then
                        c4 = Mid$(t, a + 4, 1)
                        If EsVocal_IB(c4) Then
                            silFin(i) = a + 4
                            silIni(i + 1) = a + 5
                            GoTo Siguiente
                        End If
                    End If
                End If

                ' Regla general
                If PuedeCerrarSilaba_IB(c2) Then
                    silFin(i) = a + 2
                    silIni(i + 1) = a + 3
                Else
                    silFin(i) = a + 1
                    silIni(i + 1) = a + 2
                End If

            Case Else
                silFin(i) = a + 2
                silIni(i + 1) = a + 3

        End Select

Siguiente:
    Next i

    silFin(nNuc) = L

End Sub

' ============================================================
'   PREFIJOS IB (MODO MORFOSILÁBICO)
' ============================================================
Private Function DetectarPrefijo_IB(ByVal t As String) As String
    Dim p As Variant

    If Not prefijosCargados_IB Then CargarPrefijos_IB

    For Each p In prefijosEstrictos_IB
        If Len(t) = Len(p) Then Exit For
        If Left$(t, Len(p)) = p Then
            DetectarPrefijo_IB = p
            Exit Function
        End If
    Next p

    DetectarPrefijo_IB = ""
End Function

Private Sub CargarPrefijos_IB()
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim i As Long

    If prefijosCargados_IB Then Exit Sub

'ToDo
    sql = "SELECT Prefijo FROM qryPrefijos " & _
          "WHERE Activo = 1 " & _
            "AND Tipo Like 'auténtico' " & _
            "AND [ca-ib] = true " & _
          "ORDER BY Len(Prefijo) DESC, Prefijo ASC"

    Set rs = CurrentDb.OpenRecordset(sql)

    If Not rs.EOF Then
        rs.MoveLast
        ReDim prefijosEstrictos_IB(1 To rs.RecordCount)
        rs.MoveFirst

        i = 1
        Do Until rs.EOF
            prefijosEstrictos_IB(i) = LCase$(rs!Prefijo)
            i = i + 1
            rs.MoveNext
        Loop
    End If

    rs.Close
    prefijosCargados_IB = True
End Sub

Public Function SilabearMorfologico_IB(ByVal t As String) As String
    Dim pref As String
    Dim resto As String

    If Not respetarPrefijos Then
        SilabearMorfologico_IB = SilabearOrtog_IB(t)
        Exit Function
    End If

    pref = DetectarPrefijo_IB(t)

    If pref = "" Then
        SilabearMorfologico_IB = SilabearOrtog_IB(t)
        Exit Function
    End If

    resto = Mid$(t, Len(pref) + 1)

    SilabearMorfologico_IB = pref & " | " & SilabearOrtog_IB(resto)
End Function

' ============================================================
'   NORMALIZACIÓN ORTOGRÁFICA IB
' ============================================================
Function Normalizar_IB(t As String) As String
    Dim s As String

    s = t

    ' 1. Normalizar apóstrofos
    s = Replace(s, "’", "'")
    s = Replace(s, "´", "'")
    s = Replace(s, "`", "'")

    ' 2. No tocar vocales balears (á, à, è, é, í, ï, ò, ó, ú, ü, ç)

    ' 3. Normalizar guiones tipográficos
    s = Replace(s, "–", "-")
    s = Replace(s, "—", "-")

    ' 4. Espacios múltiples
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop

    s = Trim$(s)

    Normalizar_IB = s
End Function

' ============================================================
'   CIERRE DE SÍLABA IB
' ============================================================
Private Function PuedeCerrarSilaba_IB(ByVal c As String) As Boolean
    ' Consonantes que NO pueden cerrar sílaba en IB (ajustable)
    PuedeCerrarSilaba_IB = Not (c = "r" Or c = "l" Or c = "h")
End Function

' ============================================================
'   GRUPOS DE ATAQUE IB
' ============================================================
Private Function EsGrupoAtaque_IB(ByVal g As String) As Boolean
    Dim AC As Variant
    AC = Array("pr", "br", "tr", "dr", "cr", "gr", "fr", _
               "pl", "bl", "cl", "gl", "fl")
    EsGrupoAtaque_IB = (UBound(Filter(AC, g)) >= 0)
End Function

' ============================================================
'   FUNCIONES DE VOCAL — IB
' ============================================================
Function EsVocal_IB(c As String) As Boolean
    Select Case c
        Case "a", "à", "á", _
             "e", "è", "é", _
             "i", "í", "ï", _
             "o", "ò", "ó", _
             "u", "ú", "ü"
            EsVocal_IB = True
        Case Else
            EsVocal_IB = False
    End Select
End Function

Function EsVocalForta_IB(c As String) As Boolean
    Select Case c
        Case "a", "à", "á", _
             "e", "è", "é", _
             "o", "ò", "ó"
            EsVocalForta_IB = True
        Case Else
            EsVocalForta_IB = False
    End Select
End Function

Function EsVocalFeble_IB(c As String) As Boolean
    Select Case c
        Case "i", "í", "ï", _
             "u", "ú", "ü"
            EsVocalFeble_IB = True
        Case Else
            EsVocalFeble_IB = False
    End Select
End Function

Function EsSemivocal_IB(c As String) As Boolean
    Select Case c
        Case "i", "í", "ï", _
             "u", "ú", "ü"
            EsSemivocal_IB = True
        Case Else
            EsSemivocal_IB = False
    End Select
End Function

Function EsDiptong_IB(c1 As String, c2 As String) As Boolean
    Dim par As String
    par = c1 & c2

    ' Secuencias explícitamente NO diptongo
    Select Case par
        Case "aï", "eï", "oï", "uï", _
             "aü", "eü", "oü", _
             "qü", "qüe", "qüi", "qüo"
            EsDiptong_IB = False
            Exit Function
    End Select

    ' Diptongos decrecientes IB
    Select Case par
        Case "ai", "ei", "oi", "ui", _
             "au", "eu", "ou"
            EsDiptong_IB = True
            Exit Function
    End Select

    ' Diptongos crecientes IB
    Select Case par
        Case "ia", "ie", "io", "iu", _
             "ua", "ue", "uo", "ui"
            EsDiptong_IB = True
            Exit Function
    End Select

    ' Diptongos con dièresi
    Select Case par
        Case "üa", "üe", "üi", "üo"
            EsDiptong_IB = True
            Exit Function
    End Select

    EsDiptong_IB = False
End Function

Function EsHiat_IB(c1 As String, c2 As String) As Boolean
    If EsVocal_IB(c1) And EsVocal_IB(c2) Then
        EsHiat_IB = Not EsDiptong_IB(c1, c2)
    Else
        EsHiat_IB = False
    End If
End Function

'=================================================================
'=================================================================
'                 SECCIÓN MÓDULO ACENTOS
'=================================================================
'=================================================================

' ============================================================
'   DETECTAR SÍLABAS TÓNICAS IB
' ============================================================
Private Sub CalcularTonicas_IB()

    Dim tGlobal As New Collection
    Dim elementos As Collection
    Set elementos = ObtenerPalabrasDesdeSilabasAuto_IB()

    Dim globalIndex As Long
    Dim i As Long

    globalIndex = 0

    For i = 1 To elementos.count

        If TypeName(elementos(i)) = "Collection" Then
            Dim w As Collection
            Set w = elementos(i)

            Dim tLocal As Long
            tLocal = DetectarTonica_IB(w)

            If tLocal > 0 Then
                tGlobal.Add globalIndex + tLocal
            End If

            globalIndex = globalIndex + w.count

        Else
            ' HUECO
            globalIndex = globalIndex + 1
        End If

    Next i

    ObjDTO.SilabasTonicas = JoinCollection_IB(tGlobal)

End Sub

' ============================================================
'   DETECTAR TÓNICA LOCAL EN UNA PALABRA IB
' ============================================================
Private Function DetectarTonica_IB(w As Collection) As Long

    Dim i As Long
    Dim palabra As String
    Dim ultima As String
    Dim terminaLlana As Boolean

    ' 1) Si alguna sílaba tiene tilde gráfica ? tónica directa
    For i = 1 To w.count
        If TieneTilde_IB(w(i)) Then
            DetectarTonica_IB = i
            Exit Function
        End If
    Next i

    ' 2) Sin tilde: aplicar regla general (similar CA, ajustable IB)
    palabra = ""
    For i = 1 To w.count
        palabra = palabra & w(i)
    Next i

    ultima = Right$(palabra, 1)
    terminaLlana = False

    ' Vocal o terminaciones típicas llanas
    If InStr("aeiouàèéíïòóúü", ultima) > 0 Then terminaLlana = True
    If LCase$(Right$(palabra, 2)) Like "*as" Then terminaLlana = True
    If LCase$(Right$(palabra, 2)) Like "*es" Then terminaLlana = True
    If LCase$(Right$(palabra, 2)) Like "*is" Then terminaLlana = True
    If LCase$(Right$(palabra, 2)) Like "*os" Then terminaLlana = True
    If LCase$(Right$(palabra, 2)) Like "*us" Then terminaLlana = True
    If LCase$(Right$(palabra, 2)) Like "*en" Then terminaLlana = True
    If LCase$(Right$(palabra, 2)) Like "*in" Then terminaLlana = True

    If terminaLlana And w.count >= 2 Then
        DetectarTonica_IB = w.count - 1   ' penúltima
    Else
        DetectarTonica_IB = w.count       ' última
    End If

End Function

' ============================================================
'   DETECTAR SÍLABAS SECUNDARIAS IB
' ============================================================
Private Sub CalcularSecundarias_IB()

    Dim sGlobal As New Collection
    Dim elementos As Collection
    Set elementos = ObtenerPalabrasDesdeSilabasAuto_IB()

    Dim globalIndex As Long
    Dim i As Long
    Dim tLocal As Long

    globalIndex = 0

    For i = 1 To elementos.count

        If TypeName(elementos(i)) = "Collection" Then

            Dim w As Collection
            Set w = elementos(i)

            ' Tónica local
            tLocal = DetectarTonica_IB(w)

            ' Secundarias locales
            Dim secs As Collection
            Set secs = DetectarSecundarias_IB(w, tLocal)

            ' Pasar a índices globales
            Dim x As Variant
            For Each x In secs
                sGlobal.Add globalIndex + CLng(x)
            Next x

            globalIndex = globalIndex + w.count

        Else
            globalIndex = globalIndex + 1
        End If

    Next i

    ObjDTO.SilabasSecundarias = JoinCollection_IB(sGlobal)

End Sub

Private Function DetectarSecundarias_IB(w As Collection, tPos As Long) As Collection

    Dim secs As New Collection
    Dim n As Long
    Dim pos2 As Long

    n = w.count

    ' 1–3 sílabas ? sin secundaria
    If n < 4 Then
        Set DetectarSecundarias_IB = secs
        Exit Function
    End If

    ' Primera secundaria siempre en la sílaba 1
    secs.Add 1

    ' Palabras de 6+ sílabas ? segunda secundaria
    If n >= 6 Then
        pos2 = tPos - 2
        If pos2 > 1 Then
            secs.Add pos2
        End If
    End If

    Set DetectarSecundarias_IB = secs

End Function

' ============================================================
'   MARCAR TÓNICAS Y SECUNDARIAS EN LA CADENA FINAL IB
' ============================================================
Private Sub MarcarTonicaYSecundariaEnCadena_IB()

    Dim sils As Variant
    Dim i As Long
    Dim out() As String

    Dim t As Variant
    Dim x As Variant

    sils = Split(ObjDTO.SilabasAuto, " | ")

    ReDim out(LBound(sils) To UBound(sils))

    For i = LBound(sils) To UBound(sils)
        out(i) = sils(i)
    Next i

    ' 1) TÓNICAS
    If ObjDTO.SilabasTonicas <> "" Then

        t = Split(ObjDTO.SilabasTonicas, ",")

        For Each x In t
            Dim idx As Long
            idx = CLng(x) - 1

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
        Dim idx2 As Long

        s = Split(ObjDTO.SilabasSecundarias, ",")

        For Each y In s
            idx2 = CLng(y) - 1

            If idx2 >= LBound(out) And idx2 <= UBound(out) Then
                If Trim$(out(idx2)) <> "" Then
                    out(idx2) = "[ " & out(idx2) & " ]"
                End If
            End If
        Next y
    End If

    ObjDTO.SilabasAcentuadas = Join(out, " | ")

End Sub

' ============================================================
'   DETECTAR TILDE EN UNA SÍLABA IB
' ============================================================
Private Function TieneTilde_IB(ByVal silaba As String) As Boolean
    TieneTilde_IB = (InStr(silaba, "à") > 0 Or _
                     InStr(silaba, "á") > 0 Or _
                     InStr(silaba, "è") > 0 Or _
                     InStr(silaba, "é") > 0 Or _
                     InStr(silaba, "í") > 0 Or _
                     InStr(silaba, "ï") > 0 Or _
                     InStr(silaba, "ò") > 0 Or _
                     InStr(silaba, "ó") > 0 Or _
                     InStr(silaba, "ú") > 0 Or _
                     InStr(silaba, "ü") > 0)
End Function

' ============================================================
'   OBTENER PALABRAS DESDE SILABAS AUTO IB
' ============================================================
Private Function ObtenerPalabrasDesdeSilabasAuto_IB() As Collection

    Dim resultado As New Collection
    Dim palabraActual As New Collection

    Dim sils As Variant
    sils = Split(ObjDTO.SilabasAuto, " | ")

    Dim i As Long
    For i = LBound(sils) To UBound(sils)

        If Trim$(sils(i)) = "" Then
            ' HUECO
            If palabraActual.count > 0 Then
                resultado.Add palabraActual
                Set palabraActual = New Collection
            End If
            resultado.Add "HUECO"
        Else
            palabraActual.Add sils(i)
        End If

    Next i

    If palabraActual.count > 0 Then resultado.Add palabraActual

    Set ObtenerPalabrasDesdeSilabasAuto_IB = resultado

End Function

' ============================================================
'   JOIN COLLECTION IB
' ============================================================
Private Function JoinCollection_IB(col As Collection) As String

    Dim arr() As String
    Dim i As Long

    If col Is Nothing Then
        JoinCollection_IB = ""
        Exit Function
    End If

    If col.count = 0 Then
        JoinCollection_IB = ""
        Exit Function
    End If

    ReDim arr(1 To col.count)

    For i = 1 To col.count
        arr(i) = CStr(col(i))
    Next i

    JoinCollection_IB = Join(arr, ",")

End Function

