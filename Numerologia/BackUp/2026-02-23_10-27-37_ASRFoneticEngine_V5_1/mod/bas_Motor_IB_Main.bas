Attribute VB_Name = "bas_Motor_IB_Main"

' ============================================================
'   MOTOR FONÉTICO — ILLAS BALEARS (MÓDULO ORTOGRÁFICO)
'   Arquitectura idéntica al motor catalán
' ============================================================

Option Compare Database
Option Explicit

Private usarSilabeoMorfologico As Boolean
Private modoPrefijosEstrictos As Boolean
Private respetarPrefijos As Boolean

Private prefijosEstrictos As Variant
Private prefijosCargados As Boolean

Private Const strSQL As String = _
        "SELECT Prefijo FROM qryPrefijos " & _
        "WHERE Activo = 1 " & _
            "AND Tipo Like 'auténtico' " & _
            "AND [ca-ib] = 1 " & _
        "ORDER BY Len(Prefijo) DESC, Prefijo ASC"

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

    ' 1) Normalización general (DTO)
    ObjDTO.TextoOriginal = texto
    ObjDTO.NormalizaEntrada

    ' 2) Silabeo automático
    Call SilabearFrase

    ' 3) Detectar tónica
    Call CalcularTonicas

    ' 4) Detectar secundárias
    Call CalcularSecundarias

    ' 5) Marcar tónica i secundárias
    Call MarcarTonicaYSecundariaEnCadena

    ObjDTO.SilabasFinal = ObjDTO.SilabasAcentuadas

    ' 6) Fonètica
    Call ConstruirCadenaFonemas_IB

''    << Eliminar en producción >>
'    Call MF_DebugDTO("Silabear")
    
    ' 7) Retorn (igual que català)
    Entrada_Motor_IB = ObjDTO.SilabasAuto

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

    If DebugMotor Then
        addLog
        addLog "SilabearFrase ? " & ObjDTO.SilabasAuto
    End If

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
'   SILABEO ORTOGRÁFICO IB
'   (estructura idéntica a CA, reglas IB dentro)
' ============================================================
Private Function SilabearOrtog(ByVal t As String) As String
    Dim nucIni() As Byte, nucFin() As Byte
    Dim silIni() As Byte, silFin() As Byte
    Dim nNuc As Byte, i As Byte
    Dim silabas() As String
    
    If Len(Trim$(t)) < 2 Then
        If t = "y" Then
            SilabearOrtog = "y"
            Exit Function
        ElseIf t = "i" Then
            SilabearOrtog = "i"
            Exit Function
        End If
    End If

    LocalizarNucleosOrtog t, nucIni, nucFin, nNuc
                                                  
    ReDim silIni(1 To nNuc)
    ReDim silFin(1 To nNuc)
                                                        
    CalcularSilabas t, nucIni, nucFin, nNuc, silIni, silFin
                                                  
    ReDim silabas(1 To nNuc)
    For i = 1 To nNuc
        silabas(i) = Mid$(t, silIni(i), silFin(i) - silIni(i) + 1)
        If DebugMotor Then
            addLog "IB — Sílaba " & i & ": " & silabas(i)
        End If
    Next i

    SilabearOrtog = Join(silabas, " | ")

End Function

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
'   PREFIJOS IB (MODO MORFOSILÁBICO)
' ============================================================
Private Function DetectarPrefijo(ByVal t As String) As String
    Dim p As Variant
    Dim resto As String
    Dim primera As String, segunda As String

    If Not prefijosCargados Then CargarPrefijos

    For Each p In prefijosEstrictos

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
            If Not EsVocal(primera) Then
                GoTo SiguienteP
            End If

            ' 2.2) No puede romper un diptongo
            If EsDiptongo(Left$(t, 1), Mid$(t, 2, 1)) Then
                GoTo SiguienteP
            End If

            ' 2.3) El ataque resultante debe ser válido
            If Not AtaqueSilabicoValido(primera, segunda) Then
                GoTo SiguienteP
            End If
            
            DetectarPrefijo = p
            Exit Function
        End If

        ' ============================================================
        '   PREFIJOS NORMALES (anti-, inter-, sub-, re-, etc.)
        ' ============================================================

        ' 3) El ataque resultante debe ser válido
        If Not AtaqueSilabicoValido(primera, segunda) Then
            GoTo SiguienteP
        End If

        DetectarPrefijo = p
        Exit Function

SiguienteP:
    Next p

    DetectarPrefijo = ""
End Function

Private Sub CargarPrefijos()
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim i As Long

    If prefijosCargados Then Exit Sub

'    sql = "SELECT Prefijo FROM qryPrefijos " & _
          "WHERE Activo = 1 " & _
            "AND Tipo Like 'auténtico' " & _
            "AND [ca-ib] = true " & _
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

Private Function AtaqueSilabicoValido(ByVal a As String, ByVal b As String) As Boolean
    Dim grupo As String
    grupo = LCase$(a & b)

    Select Case grupo
        Case "pr", "pl", "br", "bl", "tr", "dr", "cr", "cl", "gr", "gl", "fr", "fl"
            AtaqueSilabicoValido = True
        Case Else
            AtaqueSilabicoValido = False
    End Select
End Function


' ============================================================
'   LOCALIZAR NÚCLEOS VOCÁLICOS IB
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

    If DebugMotor Then
        addLog
        addLog "---------------------------------------"
        addLog " Procedimiento LocalizarNucleosOrtog_IB"
    End If

    i = 1

    Do While i <= L
        c1 = Mid$(t, i, 1)

        If EsVocal(c1) Then
            ' Triptongo IB
            If i + 2 <= L Then
                c2 = Mid$(t, i + 1, 1)
                c3 = Mid$(t, i + 2, 1)
                If EsTriptongo(c1, c2, c3) Then
                    nNuc = nNuc + 1
                    nucIni(nNuc) = i
                    nucFin(nNuc) = i + 2
                    If DebugMotor Then
                        addLog "Triptongo IB: " & c1 & c2 & c3
                    End If
                    i = i + 3
                    GoTo Siguiente
                End If
            End If

            ' Diptongo IB
            If i + 1 <= L Then
                c2 = Mid$(t, i + 1, 1)
                If EsDiptongo(c1, c2) Then
                    nNuc = nNuc + 1
                    nucIni(nNuc) = i
                    nucFin(nNuc) = i + 1
                    If DebugMotor Then
                        addLog "Diptongo IB: " & c1 & c2
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
                addLog "Vocal sola IB: " & c1
            End If
            i = i + 1
                                                  

        Else
            i = i + 1
        End If

Siguiente:
    Loop

    If DebugMotor Then
        addLog "Total núcleos IB: " & nNuc
        addLog " Fin LocalizarNucleosOrtog"
        addLog "---------------------------------------"
    End If
End Sub

' ============================================================
'   CÁLCULO DE SÍLABAS — IB
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
    Dim c1 As String, c2 As String, c3 As String, grupo As String

    L = Len(t)
    silIni(1) = 1

    For i = 1 To nNuc - 1
        a = nucFin(i)
        b = nucIni(i + 1)

        k = IIf(b > a + 1, b - a - 1, 0)

        If DebugMotor Then
            addLog
            addLog "---- Frontera IB entre núcleo " & i & " y " & (i + 1)
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

                ' Dígrafos indivisibles IB
                If grupo = "rr" Or grupo = "ll" Or grupo = "ch" Then
                    silFin(i) = a
                    silIni(i + 1) = a + 1
                    GoTo Siguiente
                End If

                ' Grupos de ataque IB
                If EsGrupoAtaque(grupo) Then
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

                If PuedeCerrarSilaba(c2) Then
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
'   FUNCIONES DE VOCAL — IB
' ============================================================
Private Function EsVocal(c As String) As Boolean
    Select Case c
        Case "a", "à", "á", _
             "e", "è", "é", _
             "i", "í", "ï", _
             "o", "ò", "ó", _
             "u", "ú", "ü"
            EsVocal = True
        Case Else
            EsVocal = False
    End Select
End Function

Private Function EsVocalFuerte(c As String) As Boolean
    Select Case c
        Case "a", "à", "á", _
             "e", "è", "é", _
             "o", "ò", "ó"
            EsVocalFuerte = True
        Case Else
            EsVocalFuerte = False
    End Select
End Function

Private Function EsVocalDebilTonica(ByVal c As String) As Boolean
    EsVocalDebilTonica = InStr("íú", LCase$(c)) > 0
End Function

Private Function EsVocalDebil(ByVal c As String) As Boolean
    EsVocalDebil = InStr("iíïuúü", LCase$(c)) > 0
End Function

Private Function EsSemivocal(c As String) As Boolean
    Select Case c
        Case "i", "í", "ï", _
             "u", "ú", "ü"
            EsSemivocal = True
        Case Else
            EsSemivocal = False
    End Select
End Function

Private Function EsTriptongo(ByVal v1 As String, ByVal v2 As String, ByVal v3 As String) As Boolean
    If EsVocalDebil(v1) And Not EsVocalDebilTonica(v1) _
       And EsVocalFuerte(v2) _
       And EsVocalDebil(v3) And Not EsVocalDebilTonica(v3) Then
        EsTriptongo = True
    End If
End Function

Private Function EsDiptongo(c1 As String, c2 As String) As Boolean
    Dim par As String
    par = c1 & c2

    ' Secuencias explícitamente NO diptongo
    Select Case par
        Case "aï", "eï", "oï", "uï", _
             "aü", "eü", "oü", _
             "qü", "qüe", "qüi", "qüo"
            EsDiptongo = False
            Exit Function
    End Select

    ' Diptongos decrecientes IB
    Select Case par
        Case "ai", "ei", "oi", "ui", _
             "au", "eu", "ou"
            EsDiptongo = True
            Exit Function
    End Select

    ' Diptongos crecientes IB
    Select Case par
        Case "ia", "ie", "io", "iu", _
             "ua", "ue", "uo", "ui"
            EsDiptongo = True
            Exit Function
    End Select

    ' Diptongos con dièresi
    Select Case par
        Case "üa", "üe", "üi", "üo"
            EsDiptongo = True
            Exit Function
    End Select

    EsDiptongo = False
End Function

' ============================================================
'   CIERRE DE SÍLABA IB
' ============================================================
Private Function PuedeCerrarSilaba(ByVal c As String) As Boolean
    ' Consonantes que NO pueden cerrar sílaba en IB (ajustable)
    PuedeCerrarSilaba = Not (c = "r" Or c = "l" Or c = "h")
End Function

' ============================================================
'   GRUPOS DE ATAQUE IB
' ============================================================
Private Function EsGrupoAtaque(ByVal g As String) As Boolean
    Dim AC As Variant
    AC = Array("pr", "br", "tr", "dr", "cr", "gr", "fr", _
               "pl", "bl", "cl", "gl", "fl")
    EsGrupoAtaque = (UBound(Filter(AC, g)) >= 0)
End Function

' ============================================================
'   DETECTAR TILDE EN UNA SÍLABA IB
' ============================================================
Private Function TieneTilde(ByVal silaba As String) As Boolean
    TieneTilde = (InStr(silaba, "à") > 0 Or _
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

'=================================================================
'=================================================================
'                 SECCIÓN MÓDULO ACENTOS
'=================================================================
'=================================================================

' ============================================================
'   DETECTAR SÍLABES TÒNIQUES — IB
'   (estructura idèntica al CA)
' ============================================================
Private Sub CalcularTonicas()

    Dim tGlobal As New Collection
    Dim elementos As Collection
    Set elementos = ObtenerPalabrasDesdeSilabasAuto()

    Dim globalIndex As Byte
    Dim i As Byte

    globalIndex = 0

    For i = 1 To elementos.count

        If TypeName(elementos(i)) = "Collection" Then
            Dim w As Collection
            Set w = elementos(i)

            Dim tLocal As Byte
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

    If DebugMotor Then
        addLog "CalcularTonicas ? " & ObjDTO.SilabasTonicas
    End If

End Sub

' ============================================================
'   DETECTAR SÍLABAS SECUNDARIAS IB
' ============================================================
Private Sub CalcularSecundarias() 'Malo

    Dim sGlobal As New Collection
    Dim elementos As Collection
    Set elementos = ObtenerPalabrasDesdeSilabasAuto()

    Dim globalIndex As Byte
    Dim i As Byte
    Dim tLocal As Byte

    globalIndex = 0

    For i = 1 To elementos.count

        If TypeName(elementos(i)) = "Collection" Then

            Dim w As Collection
            Set w = elementos(i)

            tLocal = DetectarTonica(w)

            Dim secs As Collection
            Set secs = DetectarSecundarias(w, tLocal)


            Dim x As Variant
            For Each x In secs
                sGlobal.Add globalIndex + CByte(x)
            Next x

            globalIndex = globalIndex + w.count

        Else
            globalIndex = globalIndex + 1
        End If

    Next i

    ObjDTO.SilabasSecundarias = JoinCollection(sGlobal)

    If DebugMotor Then
        addLog "CalcularSecundarias ? " & ObjDTO.SilabasSecundarias
    End If
End Sub
' ============================================================
'   MARCAR TÓNICAS Y SECUNDARIAS EN LA CADENA FINAL IB
' ============================================================
Private Sub MarcarTonicaYSecundariaEnCadena()

    Dim sils As Variant
    Dim i As Byte
    Dim out() As String

    sils = Split(ObjDTO.SilabasAuto, " | ")
    ReDim out(LBound(sils) To UBound(sils))

    For i = LBound(sils) To UBound(sils)
        out(i) = sils(i)
    Next i

    ' TÒNICA
    If ObjDTO.SilabasTonicas <> "" Then
        Dim t As Variant, x As Variant
        t = Split(ObjDTO.SilabasTonicas, ",")
        For Each x In t
            Dim idx As Byte
            idx = CByte(x) - 1
            If idx >= LBound(out) And idx <= UBound(out) Then
                out(idx) = "( " & out(idx) & " )"
            End If
        Next x
    End If

    ' SECUNDÀRIES
    If ObjDTO.SilabasSecundarias <> "" Then
        Dim s As Variant, y As Variant
        s = Split(ObjDTO.SilabasSecundarias, ",")
        For Each y In s
            Dim idx2 As Byte
            idx2 = CByte(y) - 1
            If idx2 >= LBound(out) And idx2 <= UBound(out) Then
                out(idx2) = "[ " & out(idx2) & " ]"
            End If
        Next y
    End If

    ObjDTO.SilabasAcentuadas = Join(out, " | ")

    If DebugMotor Then
        addLog "MarcarTonicaYSecundariaEnCadena ? " & ObjDTO.SilabasAcentuadas
    End If

End Sub

' ============================================================
'   DETECTAR TÓNICA LOCAL EN UNA PALABRA IB
' ============================================================
Private Function DetectarTonica(w As Collection) As Long

    Dim i As Long
    Dim palabra As String
    Dim ultima As String
    Dim terminaLlana As Boolean

    ' 1) Si alguna sílaba tiene tilde gráfica ? tónica directa
    For i = 1 To w.count
        If TieneTilde(w(i)) Then
            DetectarTonica = i
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
        DetectarTonica = w.count - 1   ' penúltima
    Else
        DetectarTonica = w.count       ' última
    End If

End Function

Private Function DetectarSecundarias(w As Collection, tPos As Byte) As Collection

    Dim secs As New Collection
    Dim n As Long
    Dim pos2 As Long

    n = w.count

    ' 1–3 sílabas ? sin secundaria
    If n < 4 Then
        Set DetectarSecundarias = secs
        Exit Function
    End If

    ' Primera secundaria siempre en la sílaba 1
    secs.Add 1

     ' Palabras de 6+ sílabas ? segunda secundaria
    If n >= 6 Then
        pos2 = tPos - 2   ' dos antes de la tónica

        If pos2 > 1 Then
            secs.Add pos2
        End If
    End If
    
'    ' Palabras de 6+ sílabas ? segunda secundaria
'    If n >= 6 Then
'        pos2 = tPos - 2
'        If pos2 > 1 Then
'            secs.Add pos2
'        End If
'    End If

    Set DetectarSecundarias = secs

End Function

' ============================================================
'   OBTENER PALABRAS DESDE SILABAS AUTO IB
' ============================================================
Private Function ObtenerPalabrasDesdeSilabasAuto() As Collection

    Dim resultado As New Collection
    Dim palabraActual As New Collection

    Dim sils As Variant
    sils = Split(ObjDTO.SilabasAuto, " | ")

    Dim i As Byte
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
'   JOIN COLLECTION IB
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
