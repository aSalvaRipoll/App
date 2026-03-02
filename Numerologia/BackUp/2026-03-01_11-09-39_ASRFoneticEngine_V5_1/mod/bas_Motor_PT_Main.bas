Attribute VB_Name = "bas_Motor_PT_Main"

Option Compare Database
Option Explicit

' ============================================================
'   MOTOR DE SILABEO — PORTUGUÉS EUROPEO (PT-EU)
'   Basado en el motor GL, adaptado a PT
'   (solo silabeo + acentuación; fonética va aparte)
' ============================================================

Private usarSilabeoMorfologico As Boolean
Private modoPrefijosEstrictos As Boolean
Private respetarPrefijos As Boolean

Private prefijosPersonalizados As Variant
Private prefijosEstrictos As Variant
Private prefijosCargados As Boolean

Private Const strSQL As String = "SELECT Prefijo FROM qryPrefijos " & _
          "WHERE Activo = 1 " & _
            "AND Tipo Like 'auténtico' " & _
            "AND [pt-eu] = true " & _
          "ORDER BY Len(Prefijo) DESC, Prefijo ASC"

' ============================================================
'   ENTRADA PRINCIPAL PT-EU
' ============================================================

Public Function Entrada_Motor_PT(texto As String) As String

    Set ObjDTO = New clsDTO_Motor

    ObjDTO.TextoOriginal = texto
    ObjDTO.NormalizaEntrada

    Call SilabearFrase
    Call CalcularTonicas
    Call CalcularSecundarias
    Call MarcarTonicaYSecundariaEnCadena

    ObjDTO.SilabasFinal = ObjDTO.SilabasAcentuadas

    ' Fonética PT (otro módulo)
    Call ConstruirCadenaFonemas_PT

    Entrada_Motor_PT = ObjDTO.SilabasAuto

End Function

' ============================================================
'   SILABEO DE FRASE
' ============================================================

Private Sub SilabearFrase()

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
            sil = SilabearPalabra(limpia)

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

Private Function SilabearPalabra(ByVal texto As String) As String
    Dim t As String

    t = LCase$(Trim$(texto))

    If usarSilabeoMorfologico Then
        SilabearPalabra = SilabearMorfologico(t)
    Else
        SilabearPalabra = SilabearOrtog(t)
    End If

End Function

' ============================================================
'   SILABEO MORFOLÓGICO (PREFIJOS)
' ============================================================

Private Function SilabearMorfologico(ByVal t As String) As String
    Dim pref As String
    Dim resto As String

    If Not respetarPrefijos Then
        SilabearMorfologico = SilabearOrtog(t)
        Exit Function
    End If

    pref = DetectarPrefijo(t)

    If pref = "" Then
        SilabearMorfologico = SilabearOrtog(t)
        Exit Function
    End If

    resto = Mid$(t, Len(pref) + 1)

    SilabearMorfologico = pref & " | " & SilabearOrtog(resto)

End Function

' ============================================================
'   DETECTAR PREFIJO PT-EU
' ============================================================

Private Function DetectarPrefijo(ByVal t As String) As String
    Dim p As Variant
    Dim resto As String
    Dim primera As String, segunda As String

    If Not prefijosCargados Then CargarPrefijos

    For Each p In prefijosEstrictos

        If Left$(t, Len(p)) <> p Then GoTo SiguienteP

        resto = Mid$(t, Len(p) + 1)
        If Len(resto) = 0 Then GoTo SiguienteP

        primera = Left$(resto, 1)
        segunda = Mid$(resto, 2, 1)

        If Not AtaqueSilabicoValido(primera, segunda) Then GoTo SiguienteP

        DetectarPrefijo = p
        Exit Function

SiguienteP:
    Next p

    DetectarPrefijo = ""
End Function

' ============================================================
'   ATAQUES SILÁBICOS PT-EU
' ============================================================

Private Function AtaqueSilabicoValido(ByVal a As String, ByVal b As String) As Boolean
    Dim grupo As String
    grupo = LCase$(a & b)

    Select Case grupo
        Case "pr", "br", "tr", "dr", "cr", "gr", "fr", _
             "pl", "bl", "cl", "gl", "fl", _
             "vr", "vl", "ps", "pn", "mn", "gn"
            AtaqueSilabicoValido = True
        Case Else
            AtaqueSilabicoValido = False
    End Select
End Function

' ============================================================
'   CARGAR PREFIJOS PT-EU
' ============================================================

Private Sub CargarPrefijos()
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim i As Long

    If prefijosCargados Then Exit Sub

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

' ============================================================
'   SILABEO ORTOGRÁFICO
' ============================================================

Private Function SilabearOrtog(ByVal t As String) As String
    Dim nucIni() As Byte, nucFin() As Byte
    Dim silIni() As Byte, silFin() As Byte
    Dim nNuc As Byte, i As Byte
    Dim silabas() As String

    LocalizarNucleosOrtog t, nucIni, nucFin, nNuc

    ReDim silIni(1 To nNuc)
    ReDim silFin(1 To nNuc)

    CalcularSilabas t, nucIni, nucFin, nNuc, silIni, silFin

    ReDim silabas(1 To nNuc)
    For i = 1 To nNuc
        silabas(i) = Mid$(t, silIni(i), silFin(i) - silIni(i) + 1)
    Next i

    SilabearOrtog = Join(silabas, " | ")
End Function

' ============================================================
'   LOCALIZAR NÚCLEOS (VOCÁLICOS)
' ============================================================

Private Sub LocalizarNucleosOrtog(ByVal t As String, _
                                  ByRef nucIni() As Byte, _
                                  ByRef nucFin() As Byte, _
                                  ByRef nNuc As Byte)

    Dim i As Byte, L As Byte
    Dim c1 As String, c2 As String, c3 As String

    L = Len(t)
    ReDim nucIni(1 To L)
    ReDim nucFin(1 To L)
    nNuc = 0

    i = 1
    Do While i <= L

        c1 = Mid$(t, i, 1)

        If EsVocal(c1) Then

            If i + 2 <= L Then
                c2 = Mid$(t, i + 1, 1)
                c3 = Mid$(t, i + 2, 1)
                If EsTriptongo(c1, c2, c3) Then
                    nNuc = nNuc + 1
                    nucIni(nNuc) = i
                    nucFin(nNuc) = i + 2
                    i = i + 3
                    GoTo siguiente
                End If
            End If

            If i + 1 <= L Then
                c2 = Mid$(t, i + 1, 1)
                If EsDiptongo(c1, c2) Then
                    nNuc = nNuc + 1
                    nucIni(nNuc) = i
                    nucFin(nNuc) = i + 1
                    i = i + 2
                    GoTo siguiente
                End If
            End If

            nNuc = nNuc + 1
            nucIni(nNuc) = i
            nucFin(nNuc) = i
            i = i + 1

        Else
            i = i + 1
        End If

siguiente:
    Loop

End Sub

' ============================================================
'   CÁLCULO DE FRONTERAS SILÁBICAS
' ============================================================

Private Sub CalcularSilabas(ByVal t As String, _
                            ByRef nucIni() As Byte, _
                            ByRef nucFin() As Byte, _
                            ByVal nNuc As Byte, _
                            ByRef silIni() As Byte, _
                            ByRef silFin() As Byte)

    Dim i As Byte, L As Byte
    Dim a As Byte, b As Byte
    Dim k As Byte
    Dim c1 As String, c2 As String, grupo As String

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
                grupo = c1 & c2

                If grupo = "rr" Or grupo = "ll" Or grupo = "ch" _
                   Or grupo = "nh" Or grupo = "lh" Then
                    silFin(i) = a
                    silIni(i + 1) = a + 1
                    GoTo siguiente
                End If

                If EsGrupoAtaque(grupo) Then
                    silFin(i) = a
                    silIni(i + 1) = a + 1
                Else
                    silFin(i) = a + 1
                    silIni(i + 1) = a + 2
                End If

            Case 3
                silFin(i) = a + 1
                silIni(i + 1) = a + 2

            Case Else
                silFin(i) = a + 2
                silIni(i + 1) = a + 3

        End Select

siguiente:
    Next i

    silFin(nNuc) = L

End Sub

' ============================================================
'   INVENTARIO VOCÁLICO PT-EU
' ============================================================

Private Function EsVocal(ByVal c As String) As Boolean
    EsVocal = InStr("aeiouáéíóúâêôãõ", LCase$(c)) > 0
End Function

Private Function EsVocalFuerte(ByVal c As String) As Boolean
    EsVocalFuerte = InStr("aáâãeéêoóôõ", LCase$(c)) > 0
End Function

Private Function EsVocalDebil(ByVal c As String) As Boolean
    EsVocalDebil = InStr("iíuú", LCase$(c)) > 0
End Function

Private Function EsVocalDebilTonica(ByVal c As String) As Boolean
    EsVocalDebilTonica = InStr("íú", LCase$(c)) > 0
End Function

' ============================================================
'   DIPTONGOS PT-EU
' ============================================================

Private Function EsDiptongo(ByVal v1 As String, ByVal v2 As String) As Boolean

    If Not EsVocal(v1) Or Not EsVocal(v2) Then Exit Function

    If v1 Like "[ãõ]" Or v2 Like "[ãõ]" Then
        EsDiptongo = True
        Exit Function
    End If

    If EsVocalFuerte(v1) And EsVocalFuerte(v2) Then Exit Function

    If EsVocalDebil(v1) And EsVocalDebil(v2) Then
        EsDiptongo = True
        Exit Function
    End If

    If (EsVocalDebilTonica(v1) And EsVocalFuerte(v2)) _
    Or (EsVocalFuerte(v1) And EsVocalDebilTonica(v2)) Then
        Exit Function
    End If

    EsDiptongo = True

End Function

' ============================================================
'   TRIPTONGOS PT-EU
' ============================================================

Private Function EsTriptongo(ByVal v1 As String, ByVal v2 As String, ByVal v3 As String) As Boolean

    If EsVocalDebil(v1) And Not EsVocalDebilTonica(v1) _
       And EsVocalFuerte(v2) _
       And EsVocalDebil(v3) And Not EsVocalDebilTonica(v3) Then
        EsTriptongo = True
        Exit Function
    End If

    If v2 Like "[ãõ]" And EsVocalDebil(v1) And EsVocalDebil(v3) Then
        EsTriptongo = True
    End If

End Function

' ============================================================
'   GRUPOS DE ATAQUE PT-EU
' ============================================================

Private Function EsGrupoAtaque(ByVal g As String) As Boolean
    Dim AC As Variant
    AC = Array("pr", "br", "tr", "dr", "cr", "gr", "fr", _
               "pl", "bl", "cl", "gl", "fl", _
               "vr", "vl", "ps", "pn", "mn", "gn")
    EsGrupoAtaque = (UBound(Filter(AC, LCase$(g))) >= 0)
End Function

' ============================================================
'   ACENTUACIÓN: TILDE EN SÍLABA
' ============================================================

Private Function TieneTilde(ByVal silaba As String) As Boolean
    TieneTilde = (InStr(silaba, "á") > 0 Or _
                  InStr(silaba, "é") > 0 Or _
                  InStr(silaba, "í") > 0 Or _
                  InStr(silaba, "ó") > 0 Or _
                  InStr(silaba, "ú") > 0 Or _
                  InStr(silaba, "â") > 0 Or _
                  InStr(silaba, "ê") > 0 Or _
                  InStr(silaba, "ô") > 0)
End Function

' ============================================================
'   TERMINA EN VOCAL / N / S / M (PT-EU)
' ============================================================

Private Function TerminaEnVocalNSM(ByVal s As String) As Boolean
    Dim c As String
    c = Right$(s, 1)

    TerminaEnVocalNSM = _
        EsVocal(c) Or _
        c = "n" Or _
        c = "m" Or _
        c = "s"
End Function

' ============================================================
'   DETECTAR TÓNICA LOCAL (colección de sílabas)
' ============================================================

Private Function DetectarTonica(w As Collection) As Byte
    Dim i As Byte
    Dim ultimaLetra As String

    For i = 1 To w.count
        If TieneTilde(w(i)) Then
            DetectarTonica = i
            Exit Function
        End If
    Next i

    ultimaLetra = Right$(w(w.count), 1)

    If ultimaLetra Like "[aeiouãõâêônmNsS]" Then
        DetectarTonica = w.count - 1
    Else
        DetectarTonica = w.count
    End If
End Function

' ============================================================
'   DETECTAR TÓNICAS (GLOBAL)
' ============================================================

Private Sub CalcularTonicas()

    Dim tGlobal As New Collection
    Dim elementos As Collection
    Dim globalIndex As Long
    Dim i As Long

    Set elementos = ObtenerPalabrasDesdeSilabasAuto()
    globalIndex = 0

    For i = 1 To elementos.count

        If TypeName(elementos(i)) = "Collection" Then
            Dim w As Collection
            Dim tLocal As Long

            Set w = elementos(i)
            tLocal = DetectarTonica(w)

            If tLocal > 0 Then
                tGlobal.Add globalIndex + tLocal
            End If

            globalIndex = globalIndex + w.count

        Else
            globalIndex = globalIndex + 1
        End If

    Next i

    ObjDTO.SilabasTonicas = JoinCollection(tGlobal)

End Sub

' ============================================================
'   DETECTAR SECUNDARIAS (GLOBAL)
' ============================================================

Private Sub CalcularSecundarias()

    Dim sGlobal As New Collection
    Dim elementos As Collection
    Dim globalIndex As Long
    Dim i As Long
    Dim tLocal As Byte

    Set elementos = ObtenerPalabrasDesdeSilabasAuto()
    globalIndex = 0

    For i = 1 To elementos.count

        If TypeName(elementos(i)) = "Collection" Then

            Dim w As Collection
            Dim secs As Collection
            Dim x As Variant

            Set w = elementos(i)
            tLocal = DetectarTonica(w)
            Set secs = DetectarSecundarias(w, tLocal)

            For Each x In secs
                sGlobal.Add globalIndex + CByte(x)
            Next x

            globalIndex = globalIndex + w.count

        Else
            globalIndex = globalIndex + 1
        End If

    Next i

    ObjDTO.SilabasSecundarias = JoinCollection(sGlobal)

End Sub

Private Function DetectarSecundarias(w As Collection, tPos As Byte) As Collection
    Dim secs As New Collection
    Dim n As Byte
    Dim pos2 As Byte

    n = w.count

    If n < 4 Then
        Set DetectarSecundarias = secs
        Exit Function
    End If

    secs.Add 1

    If n >= 6 Then
        pos2 = tPos - 2
        If pos2 > 1 Then secs.Add pos2
    End If

    Set DetectarSecundarias = secs
End Function

' ============================================================
'   MARCAR TÓNICAS Y SECUNDARIAS EN LA CADENA FINAL
' ============================================================

Private Sub MarcarTonicaYSecundariaEnCadena()

    Dim sils As Variant
    Dim out() As String
    Dim i As Long
    Dim t As Variant, x As Variant
    Dim s As Variant, y As Variant
    Dim idx As Long, idx2 As Long

    sils = Split(ObjDTO.SilabasAuto, " | ")
    ReDim out(LBound(sils) To UBound(sils))

    For i = LBound(sils) To UBound(sils)
        out(i) = sils(i)
    Next i

    ' Tónicas
    If ObjDTO.SilabasTonicas <> "" Then
        t = Split(ObjDTO.SilabasTonicas, ",")
        For Each x In t
            idx = CLng(x) - 1
            If idx >= LBound(out) And idx <= UBound(out) Then
                If Trim$(out(idx)) <> "" Then
                    out(idx) = "( " & out(idx) & " )"
                End If
            End If
        Next x
    End If

    ' Secundarias
    If ObjDTO.SilabasSecundarias <> "" Then
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
'   OBTENER PALABRAS DESDE ObjDTO.SilabasAuto
' ============================================================

Private Function ObtenerPalabrasDesdeSilabasAuto() As Collection

    Dim resultado As New Collection
    Dim palabraActual As New Collection
    Dim sils As Variant
    Dim i As Long

    sils = Split(ObjDTO.SilabasAuto, " | ")

    For i = LBound(sils) To UBound(sils)

        If Trim$(sils(i)) = "" Then
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

    Set ObtenerPalabrasDesdeSilabasAuto = resultado

End Function

' ============================================================
'   JOIN COLLECTION
' ============================================================

Private Function JoinCollection(col As Collection) As String
    Dim arr() As String
    Dim i As Long

    If col Is Nothing Then
        JoinCollection = ""
        Exit Function
    End If

    If col.count = 0 Then
        JoinCollection = ""
        Exit Function
    End If

    ReDim arr(1 To col.count)

    For i = 1 To col.count
        arr(i) = CStr(col(i))
    Next i

    JoinCollection = Join(arr, ",")

End Function


