Attribute VB_Name = "bas_Motor_EU_Main"

Option Compare Database
Option Explicit

' ============================================================
'   MOTOR DE SILABEO 4.x — ORTOGRÁFICO + MORFOSILÁBICO
'   VERSIÓN EUSKERA (EU)
' ============================================================

Private usarSilabeoMorfologico As Boolean
Private modoPrefijosEstrictos As Boolean
Private respetarPrefijos As Boolean

Private prefijosEstrictos_EU As Variant
Private prefijosCargados_EU As Boolean

Private Const strSQL_EU As String = "SELECT Prefijo FROM qryPrefijos " & _
          "WHERE Activo = 1 " & _
            "AND Tipo Like 'auténtico' " & _
            "AND [eu] = true " & _
          "ORDER BY Len(Prefijo) DESC, Prefijo ASC"


Public Function Entrada_Motor_EU(texto As String) As String

    Set ObjDTO = New clsDTO_Motor

    ObjDTO.TextoOriginal = texto
    ObjDTO.NormalizaEntrada

    ' 2) Silabeo automático
    Call SilabearFrase_EU

    ' 3) Detectar tónicas (en euskera es trivial)
    Call CalcularTonicas_EU

    ' 4) Detectar secundarias (opcional)
'    Call CalcularSecundarias_EU

    ' 5) Marcar tónicas y secundarias
    Call MarcarTonicaYSecundariaEnCadena_EU

    ObjDTO.SilabasFinal = ObjDTO.SilabasAcentuadas

    ' 6) Fonética (cuando la tengas)
    Call ConstruirCadenaFonemas_EU

    Entrada_Motor_EU = ObjDTO.SilabasAuto

End Function

Private Sub SilabearFrase_EU()

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
            sil = SilabearPalabra_EU(limpia)
            If resultado = "" Then
                resultado = sil
            Else
                resultado = resultado & " |   | " & sil
            End If
        End If
    Next i

    ObjDTO.SilabasAuto = resultado

End Sub

Private Function SilabearPalabra_EU(ByVal texto As String) As String
    Dim t As String

    t = LCase$(Trim$(texto))

    If usarSilabeoMorfologico Then
        SilabearPalabra_EU = SilabearMorfologico_EU(t)
    Else
        SilabearPalabra_EU = SilabearOrtog_EU(t)
    End If

End Function

Private Function SilabearOrtog_EU(ByVal t As String) As String
    Dim nucIni() As Byte, nucFin() As Byte
    Dim silIni() As Byte, silFin() As Byte
    Dim nNuc As Byte, i As Byte
    Dim silabas() As String

    If Len(t) < 2 Then
        SilabearOrtog_EU = t
        Exit Function
    End If

    LocalizarNucleosOrtog_EU t, nucIni, nucFin, nNuc

    ReDim silIni(1 To nNuc)
    ReDim silFin(1 To nNuc)

    CalcularSilabas_EU t, nucIni, nucFin, nNuc, silIni, silFin

    ReDim silabas(1 To nNuc)
    For i = 1 To nNuc
        silabas(i) = Mid$(t, silIni(i), silFin(i) - silIni(i) + 1)
    Next i

    SilabearOrtog_EU = Join(silabas, " | ")

End Function

Private Function SilabearMorfologico_EU(ByVal t As String) As String
    Dim pref As String
    Dim resto As String

    If Not respetarPrefijos Then
        SilabearMorfologico_EU = SilabearOrtog_EU(t)
        Exit Function
    End If

    pref = DetectarPrefijo_EU(t)

    If pref = "" Then
        SilabearMorfologico_EU = SilabearOrtog_EU(t)
        Exit Function
    End If

    resto = Mid$(t, Len(pref) + 1)

    SilabearMorfologico_EU = pref & " | " & SilabearOrtog_EU(resto)
End Function

Private Function DetectarPrefijo_EU(ByVal t As String) As String
    Dim p As Variant
    Dim resto As String
    Dim primera As String, segunda As String

    If Not prefijosCargados_EU Then CargarPrefijos_EU

    For Each p In prefijosEstrictos_EU

        ' 1) El prefijo debe coincidir con el inicio
        If Left$(t, Len(p)) <> p Then GoTo SiguienteP

        ' 2) Obtener el resto de la palabra
        resto = Mid$(t, Len(p) + 1)
        If Len(resto) = 0 Then GoTo SiguienteP

        primera = Left$(resto, 1)
        segunda = Mid$(resto, 2, 1)

        ' ============================================================
        '   REGLA EU: NO SE PERMITEN ATAQUES CC
        ' ============================================================
        If Not EsVocal_EU(primera) And Not EsVocal_EU(segunda) Then
            GoTo SiguienteP
        End If

        DetectarPrefijo_EU = p
        Exit Function

SiguienteP:
    Next p

    DetectarPrefijo_EU = ""
End Function

'Private Function DetectarPrefijo_EU(ByVal t As String) As String
'    Dim p As Variant
'    Dim resto As String
'    Dim primera As String, segunda As String
'
'    If Not prefijosCargados_EU Then CargarPrefijos_EU
'
'    For Each p In prefijosEstrictos_EU
'
'        ' 1) El prefijo debe coincidir con el inicio
'        If Left$(t, Len(p)) <> p Then GoTo SiguienteP
'
'        ' 2) Obtener el resto de la palabra
'        resto = Mid$(t, Len(p) + 1)
'        If Len(resto) = 0 Then GoTo SiguienteP
'
'        primera = Left$(resto, 1)
'        segunda = Mid$(resto, 2, 1)
'
'        ' ============================================================
'        '   CASO ESPECIAL: PREFIJO "a-" (privativo)
'        ' ============================================================
'        If p = "a" Then
'
'            ' 2.1) La base debe empezar por vocal
'            If Not EsVocal_EU(primera) Then
'                GoTo SiguienteP
'            End If
'
'            ' 2.2) No puede romper un diptongo
'            If EsDiptongo_EU(Left$(t, 1), Mid$(t, 2, 1)) Then
'                GoTo SiguienteP
'            End If
'
'            ' 2.3) El ataque resultante debe ser válido
'            If Not EsGrupoAtaque_EU(primera & segunda) Then
'                GoTo SiguienteP
'            End If
'
'            DetectarPrefijo_EU = p
'            Exit Function
'        End If
'
'        ' ============================================================
'        '   PREFIJOS NORMALES (anti-, inter-, sub-, re-, etc.)
'        ' ============================================================
'
'        ' 3) El ataque resultante debe ser válido
'        If Not AtaqueSilabicoValido(primera, segunda) Then
'            GoTo SiguienteP
'        End If
'
'        DetectarPrefijo_EU = p
'        Exit Function
'
'SiguienteP:
'    Next p
'
'    DetectarPrefijo_EU = ""
'End Function

'Private Function DetectarPrefijo_EU(ByVal t As String) As String
'    Dim p As Variant
'    Dim resto As String
'
'    If Not prefijosCargados_EU Then CargarPrefijos_EU
'
'    For Each p In prefijosEstrictos_EU
'
'        If Left$(t, Len(p)) <> p Then GoTo SiguienteP
'
'        resto = Mid$(t, Len(p) + 1)
'        If Len(resto) = 0 Then GoTo SiguienteP
'
'        DetectarPrefijo_EU = p
'        Exit Function
'
'SiguienteP:
'    Next p
'
'    DetectarPrefijo_EU = ""
'End Function


Private Sub CargarPrefijos_EU()
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim i As Long

    If prefijosCargados_EU Then Exit Sub

    sql = strSQL_EU

    Set rs = CurrentDb.OpenRecordset(sql)

    If Not rs.EOF Then
        rs.MoveLast
        ReDim prefijosEstrictos_EU(1 To rs.RecordCount)
        rs.MoveFirst

        i = 1
        Do Until rs.EOF
            prefijosEstrictos_EU(i) = LCase$(rs!Prefijo)
            i = i + 1
            rs.MoveNext
        Loop
    End If

    rs.Close
    prefijosCargados_EU = True
End Sub

Private Sub LocalizarNucleosOrtog_EU(ByVal t As String, _
                                  ByRef nucIni() As Byte, _
                                  ByRef nucFin() As Byte, _
                                  ByRef nNuc As Byte)

    Dim i As Byte, L As Byte
    Dim c1 As String

    L = Len(t)
    ReDim nucIni(1 To L)
    ReDim nucFin(1 To L)
    nNuc = 0

    i = 1
    Do While i <= L
        c1 = Mid$(t, i, 1)

        If EsVocal_EU(c1) Then

            nNuc = nNuc + 1
            nucIni(nNuc) = i
            nucFin(nNuc) = i

            i = i + 1
        Else
            i = i + 1
        End If

    Loop

End Sub

Private Sub CalcularSilabas_EU(ByVal t As String, _
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
            
                ' Dígrafos indivisibles en euskera
                If grupo = "rr" Or grupo = "ll" Or grupo = "tx" Or grupo = "ts" Or grupo = "tz" Then
                    silFin(i) = a
                    silIni(i + 1) = a + 1
                    GoTo Siguiente
                End If
            
                ' Grupos de ataque válidos
                If EsGrupoAtaque_EU(grupo) Then
                    silFin(i) = a
                    silIni(i + 1) = a + 1
                Else
                    silFin(i) = a + 1
                    silIni(i + 1) = a + 2
                End If


'            Case 2
'                c1 = Mid$(t, a + 1, 1)
'                c2 = Mid$(t, a + 2, 1)
'                grupo = c1 & c2
'
'                If EsGrupoAtaque_EU(grupo) Then
'                    silFin(i) = a
'                    silIni(i + 1) = a + 1
'                Else
'                    silFin(i) = a + 1
'                    silIni(i + 1) = a + 2
'                End If

            Case Else
                silFin(i) = a + 1
                silIni(i + 1) = a + 2

        End Select
Siguiente:

    Next i

    silFin(nNuc) = L

End Sub

Private Function EsVocal_EU(ByVal c As String) As Boolean
    EsVocal_EU = InStr("aeiou", LCase$(c)) > 0
End Function

Private Function EsDiptongo_EU(v1 As String, v2 As String) As Boolean
    EsDiptongo_EU = False
End Function

Private Function EsTriptongo_EU(v1 As String, v2 As String, v3 As String) As Boolean
    EsTriptongo_EU = False
End Function

Private Function EsGrupoAtaque_EU(ByVal grupo As String) As Boolean
    grupo = LCase$(grupo)

    Select Case grupo
        Case "pl", "bl", "kl", "gl", _
             "pr", "br", "tr", "dr", "kr", "gr", _
             "fl", "fr", _
             "tx", "ts", "tz"
            EsGrupoAtaque_EU = True
        Case Else
            EsGrupoAtaque_EU = False
    End Select
End Function


'=======================================================================
Private Sub CalcularTonicas_EU()

    Dim tGlobal As New Collection
    Dim elementos As Collection
    Set elementos = ObtenerPalabrasDesdeSilabasAuto_EU()

    Dim globalIndex As Integer
    Dim i As Byte

    globalIndex = 0

    For i = 1 To elementos.count

        If TypeName(elementos(i)) = "Collection" Then
            Dim w As Collection
            Set w = elementos(i)

            Dim tLocal As Long
            tLocal = DetectarTonica_EU(w)

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

Private Function DetectarTonica_EU(w As Collection) As Byte

    If w.count = 1 Then
        DetectarTonica_EU = 1
    Else
        DetectarTonica_EU = w.count - 1
    End If

End Function

Private Function DetectarSecundarias_EU(w As Collection, tLocal As Byte) As Collection
    Dim c As New Collection
    ' En euskera no hay secundarias normativas ? devolver vacío
    Set DetectarSecundarias_EU = c
End Function

Private Sub MarcarTonicaYSecundariaEnCadena_EU()

    Dim sils As Variant
    Dim i As Integer
    Dim out() As String

    Dim t As Variant
    Dim x As Variant

    sils = Split(ObjDTO.SilabasAuto, " | ")

    ReDim out(LBound(sils) To UBound(sils))

    For i = LBound(sils) To UBound(sils)
        out(i) = sils(i)
    Next i

    ' TÓNICAS
    If ObjDTO.SilabasTonicas <> "" Then

        t = Split(ObjDTO.SilabasTonicas, ",")

        For Each x In t
            Dim idx As Long
            idx = CInt(x) - 1

            If idx >= LBound(out) And idx <= UBound(out) Then
                If Trim$(out(idx)) <> "" Then
                    out(idx) = "( " & out(idx) & " )"
                End If
            End If
        Next x
    End If

    ObjDTO.SilabasAcentuadas = Join(out, " | ")

End Sub

Private Function ObtenerPalabrasDesdeSilabasAuto_EU() As Collection

    Dim resultado As New Collection
    Dim partes As Variant
    Dim i As Long
    Dim sils As Variant
    Dim w As Collection
    Dim s As Variant

    ' Dividir por separador de palabras
    partes = Split(ObjDTO.SilabasAuto, " |   | ")

    For i = LBound(partes) To UBound(partes)

        ' Dividir cada palabra en sílabas
        sils = Split(partes(i), " | ")

        Set w = New Collection

        For Each s In sils
            If Trim$(s) <> "" Then
                w.Add s
            End If
        Next s

        resultado.Add w

    Next i

    Set ObtenerPalabrasDesdeSilabasAuto_EU = resultado

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

