Attribute VB_Name = "bas_Motor_VA_Main"

Option Compare Database
Option Explicit

' ============================================================
'   MOTOR DE SILABEO 4.x — ORTOGRÁFICO + MORFOSILÁBICO + ACENTUACIÓN
'   VERSIÓN VALENCIANO
' ============================================================

Private usarSilabeoMorfologico As Boolean
Private modoPrefijosEstrictos As Boolean
Private respetarPrefijos As Boolean

Private prefijosEstrictos As Variant
Private prefijosCargados As Boolean

Private Const strSQL As String = "SELECT Prefijo FROM qryPrefijos " & _
          "WHERE Activo = 1 " & _
            "AND Tipo Like 'auténtico' " & _
            "AND [ca-va] = true " & _
          "ORDER BY Len(Prefijo) DESC, Prefijo ASC"

' ============================================================
'   ENTRADA PRINCIPAL DEL MOTOR (VALENCIANO)
' ============================================================
Public Function Entrada_Motor_VA(texto As String) As String

    Set ObjDTO = New clsDTO_Motor

    ObjDTO.TextoOriginal = texto
    ObjDTO.NormalizaEntrada

    ' 2) Silabeo automático
    Call SilabearFrase

    ' 3) Detectar tónicas
    Call CalcularTonicas

    ' 4) Detectar secundarias
    Call CalcularSecundarias

    ' 5) Marcar Tónicas y Secundarias
    Call MarcarTonicaYSecundariaEnCadena

    ObjDTO.SilabasFinal = ObjDTO.SilabasAcentuadas

    ' 6) Generar fonética Valenciano
    Call ConstruirCadenaFonemas_VA

'    ' Debug opcional
'    Call MF_DebugDTO("Silabear")

    Entrada_Motor_VA = ObjDTO.SilabasAuto

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

    If DebugMotor Then
        addLog
        addLog "---------------------------------------"
        addLog " Procedimiento SilabearFrase"
        addLog
        addLog "frase: " & frase
        
    End If

    For i = LBound(palabras) To UBound(palabras)
    
        If DebugMotor Then
            addLog
            addLog "Bucle"
            addLog
            addLog "i: " & i
            
        End If
    
        limpia = Trim$(palabras(i))
        
        If DebugMotor Then
            addLog
            addLog "palabras(" & i & "): " & palabras(i)
            addLog
            addLog "limpia: " & limpia
            
        End If
        
        
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

    If DebugMotor Then
        addLog
        addLog "---------------------------------------"
        addLog " Procedimiento SilabearPalabra"
        addLog
        addLog "texto: " & texto
        
    End If

    t = LCase$(Trim$(texto))

    If usarSilabeoMorfologico Then
        SilabearPalabra = SilabearMorfologico(t)
    Else
        SilabearPalabra = SilabearOrtog(t)
    End If

End Function

' ============================================================
'   SILABEO ORTOGRÁFICO VA
' ============================================================
Private Function SilabearOrtog(ByVal t As String) As String
    Dim nucIni() As Byte, nucFin() As Byte
    Dim silIni() As Byte, silFin() As Byte
    Dim nNuc As Byte, i As Byte
    Dim silabas() As String

    If DebugMotor Then
        addLog
        addLog "---------------------------------------"
        addLog " Procedimiento SilabearOrtog"
        addLog
        addLog "t: " & t
        
    End If


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
            addLog "VA — Sílaba " & i & ": " & silabas(i)
        End If
    Next i

    SilabearOrtog = Join(silabas, " | ")

End Function

' ============================================================
'   SILABEO MORFOSILÁBICO VA (PREFIJOS)
' ============================================================
Private Function SilabearMorfologico(ByVal t As String) As String
    Dim pref As String
    Dim resto As String

    If DebugMotor Then
        addLog
        addLog "---------------------------------------"
        addLog " Procedimiento SilabearMorfologico"
        addLog
        addLog "t: " & t
        
    End If
    
    If Not respetarPrefijos Then
        SilabearMorfologico = SilabearOrtog(t)
        Exit Function
    End If

    pref = DetectarPrefijo(t)

    addLog
    addLog "pref: " & pref
    
    If pref = "" Then
        SilabearMorfologico = SilabearOrtog(t)
        Exit Function
    End If

    resto = Mid$(t, Len(pref) + 1)

    SilabearMorfologico = pref & " | " & SilabearOrtog(resto)
End Function

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
            If Not EsGrupoAtaque(primera & segunda) Then
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

Private Function AtaqueSilabicoValido(ByVal a As String, ByVal b As String) As Boolean
    Dim grupo As String
    grupo = a & b

    Select Case grupo
        Case "pr", "pl", "br", "bl", "tr", "dr", "cr", "cl", "gr", "gl", "fr", "fl"
            AtaqueSilabicoValido = True
        Case Else
            AtaqueSilabicoValido = False
    End Select
End Function

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
'   LOCALIZAR NÚCLEOS VOCÁLICOS VA
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
        addLog " Procedimiento LocalizarNucleosOrtog"
        addLog
        addLog "t: " & t
        addLog "nucIni: " & CStr(nucIni)
        addLog "nucFin: " & CStr(nucFin)
        addLog "nNuc: " & CStr(nNuc)
        
    End If

    i = 1
    Do While i <= L
        c1 = Mid$(t, i, 1)

        If EsVocal(c1) Then

            ' Triptongo
            If i + 2 <= L Then
                c2 = Mid$(t, i + 1, 1)
                c3 = Mid$(t, i + 2, 1)
                If EsTriptongo(c1, c2, c3) Then
                    nNuc = nNuc + 1
                    nucIni(nNuc) = i
                    nucFin(nNuc) = i + 2
                    If DebugMotor Then
                        addLog "Triptongo VA: " & c1 & c2 & c3
                    End If
                    i = i + 3
                    GoTo Siguiente
                End If
            End If

            ' Diptongo
            If i + 1 <= L Then
                c2 = Mid$(t, i + 1, 1)
                If EsDiptongo(c1, c2) Then
                    nNuc = nNuc + 1
                    nucIni(nNuc) = i
                    nucFin(nNuc) = i + 1
                    If DebugMotor Then
                        addLog "Diptongo VA: " & c1 & c2
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
                addLog "Vocal sola VA: " & c1
            End If
            i = i + 1

        Else
            i = i + 1
        End If

Siguiente:
    Loop

    If DebugMotor Then
        addLog "Total núcleos VA: " & nNuc
        addLog " Fin LocalizarNucleosOrtog"
        addLog "---------------------------------------"
    End If
End Sub

' ============================================================
'   CÁLCULO DE SÍLABAS — VA
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
            addLog "---- Frontera VA entre núcleo " & i & " y " & (i + 1)
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
            
                ' Dígrafos indivisibles
                If grupo = "rr" Or grupo = "ll" Or grupo = "ch" Or grupo = "l·l" Then
                    silFin(i) = a
                    silIni(i + 1) = a + 1
                    GoTo Siguiente
                End If
            
                ' Ataque fonotácticamente válido
                If EsGrupoAtaque(grupo) Then
                    silFin(i) = a
                    silIni(i + 1) = a + 1
                Else
                    ' C1 se queda en la coda
                    silFin(i) = a + 1
                    silIni(i + 1) = a + 2
                End If

            Case 3
                c1 = Mid$(t, a + 1, 1)
                c2 = Mid$(t, a + 2, 1)
                c3 = Mid$(t, a + 3, 1)
            
                ' Si C2C3 es ataque válido ? VC1 | C2C3V
                If EsGrupoAtaque(c2 & c3) Or EsConsonanteSimple(c2) Then
                    silFin(i) = a + 1
                    silIni(i + 1) = a + 2
                Else
                    ' Si no ? VC1C2 | C3V
                    silFin(i) = a + 2
                    silIni(i + 1) = a + 3
                End If

            Case Else
                silFin(i) = a + 2
                silIni(i + 1) = a + 3

        End Select

Siguiente:
    Next i

    silFin(nNuc) = L

End Sub

Private Function EsConsonanteSimple(ch As String) As Boolean
    ch = LCase$(ch)
    EsConsonanteSimple = (ch Like "[bcdfghjklmnpqrstvwxyz]")
End Function

' ============================================================
'   FUNCIONES DE VOCAL — VALENCIANO
' ============================================================
Private Function EsVocal(ByVal c As String) As Boolean

    addLog
    addLog "---------------------------------------"
    addLog " Procedimiento EsVocal"
    addLog "---------------------------------------"
    addLog "c: " & c
    
    EsVocal = InStr("aeiouàèéíòóú", LCase$(c)) > 0
    addLog "EsVocal: " & EsVocal
    
End Function

Private Function EsVocalFuerte(ByVal c As String) As Boolean
    
    addLog
    addLog "---------------------------------------"
    addLog " Procedimiento EsVocalFuerte"
    addLog "---------------------------------------"
    addLog "c: " & c
    
    EsVocalFuerte = InStr("aàeèéoò", LCase$(c)) > 0
    addLog "EsVocalFuerte: " & EsVocalFuerte
    
End Function

Private Function EsVocalDebil(ByVal c As String) As Boolean
    
    addLog
    addLog "---------------------------------------"
    addLog " Procedimiento EsVocalDebil"
    addLog "---------------------------------------"
    addLog "c: " & c
    
    EsVocalDebil = InStr("iíuú", LCase$(c)) > 0
    addLog "EsVocalDebil: " & EsVocalDebil
    
End Function

Private Function EsVocalDebilTonica(ByVal c As String) As Boolean
    
    addLog
    addLog "---------------------------------------"
    addLog " Procedimiento EsVocalDebilTonica"
    addLog "---------------------------------------"
    addLog "c: " & c
    
    EsVocalDebilTonica = InStr("íú", LCase$(c)) > 0
    addLog "EsVocalDebilTonica: " & EsVocalDebilTonica
End Function

Private Function EsDiptongo(v1 As String, v2 As String) As Boolean
    v1 = LCase$(v1)
    v2 = LCase$(v2)

    addLog
    addLog "-----------------------------------------------------------"
    addLog "Función EsDiptongo:  v1 => " & v1 & "  |  v2: => " & v2
    addLog "-----------------------------------------------------------"
    
    EsDiptongo = False
    
    Select Case v1 & v2
        ' Diptongos crecientes
        Case "ia", "ie", "io", "ua", "ue", "uo"
            EsDiptongo = True
            'addLog "EsDiptongo: " & EsDiptongo
            'Exit Function

        ' Diptongos decrecientes
        Case "ai", "ei", "oi", "au", "eu", "ou"
            EsDiptongo = True
            'addLog "EsDiptongo: " & EsDiptongo
            'Exit Function

        ' Semivocales
        Case "iu", "ui"
            EsDiptongo = True
            'addLog "EsDiptongo: " & EsDiptongo
            'Exit Function
    End Select
    
    addLog "EsDiptongo: " & EsDiptongo
    
End Function

Private Function EsTriptongo(ByVal v1 As String, ByVal v2 As String, ByVal v3 As String) As Boolean

    addLog
    addLog "-----------------------------------------------------------------------------"
    addLog "Función EsTriptongo:  v1 => " & v1 & " | v2: => " & v2 & " | v2: => " & v3
    addLog "-----------------------------------------------------------------------------"
    
    If EsVocalDebil(v1) And Not EsVocalDebilTonica(v1) _
       And EsVocalFuerte(v2) _
       And EsVocalDebil(v3) And Not EsVocalDebilTonica(v3) Then
        EsTriptongo = True
    End If
    
    addLog "EsTriptongo: " & EsTriptongo
    
End Function

Private Function EsGrupoAtaque(ByVal grupo As String) As Boolean
    grupo = LCase$(grupo)

    Select Case grupo
        ' Oclusiva + líquida
        Case "pl", "bl", "cl", "gl", _
             "pr", "br", "tr", "dr", "kr", "gr"
            EsGrupoAtaque = True
            Exit Function

        ' Fricativa + líquida
        Case "fl", "fr"
            EsGrupoAtaque = True
            Exit Function

        ' S + consonante (limitado)
        Case "sp", "st", "sk"
            EsGrupoAtaque = True
            Exit Function

        ' Todo lo demás NO es ataque válido
        Case Else
            EsGrupoAtaque = False
    End Select
End Function

Private Function PuedeCerrarSilaba(ByVal c As String) As Boolean
    PuedeCerrarSilaba = Not (c = "r" Or c = "l" Or c = "h")
End Function

Private Function TieneTilde(ByVal silaba As String) As Boolean
    TieneTilde = (InStr(silaba, "à") > 0 Or _
                     InStr(silaba, "è") > 0 Or _
                     InStr(silaba, "é") > 0 Or _
                     InStr(silaba, "í") > 0 Or _
                     InStr(silaba, "ò") > 0 Or _
                     InStr(silaba, "ó") > 0 Or _
                     InStr(silaba, "ú") > 0)
End Function

'=================================================================
'                 SECCIÓN MÓDULO ACENTOS — VA
'=================================================================

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

            Dim tLocal As Long
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

Private Sub CalcularSecundarias()

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

End Sub

Private Sub MarcarTonicaYSecundariaEnCadena()

    Dim sils As Variant
    Dim i As Byte
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
            idx = CByte(x) - 1

            If idx >= LBound(out) And idx <= UBound(out) Then
                If Trim$(out(idx)) <> "" Then
                    out(idx) = "( " & out(idx) & " )"
                End If
            End If
        Next x
    End If

    ' SECUNDARIAS
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

    ObjDTO.SilabasAcentuadas = Join(out, " | ")

End Sub

Private Function DetectarTonica(w As Collection) As Byte

    Dim palabra As String
    Dim ultimaSilaba As String
    Dim grafFinal As String
    Dim i As Byte
    Dim esConsonanteFonetica As Boolean
    'Dim TieneTilde As Boolean

    ' 1) Reconstruir palabra
    palabra = ""
    For i = 1 To w.count
        palabra = palabra & w(i)
    Next i

    ultimaSilaba = w(w.count)

    ' 2) Detectar tilde explícita
    For i = 1 To w.count
        If TieneTilde(w(i)) Then
            DetectarTonica = i
            Exit Function
        End If
    Next i

    ' 3) Detectar si la última sílaba termina en consonante fonética
    esConsonanteFonetica = False

    ' --- Detectar dígrafos consonánticos finales ---
    If Right$(palabra, 2) = "tx" _
    Or Right$(palabra, 2) = "tj" _
    Or Right$(palabra, 2) = "tg" _
    Or Right$(palabra, 2) = "ix" _
    Or Right$(palabra, 2) = "ig" _
    Or Right$(palabra, 2) = "ny" _
    Or Right$(palabra, 2) = "ll" Then

        esConsonanteFonetica = True

    Else
        ' --- Detectar consonante simple final ---
        grafFinal = Right$(palabra, 1)

        If InStr("bcdfghjklmnpqrstvwxyzç", grafFinal) > 0 Then
            esConsonanteFonetica = True
        End If
    End If

    ' 4) Aplicar reglas valencianas reales

    If esConsonanteFonetica Then
        ' Palabra acabada en consonante fonética ? llana por defecto
        If w.count >= 2 Then
            DetectarTonica = w.count - 1
        Else
            DetectarTonica = 1
        End If

    Else
        ' Palabra acabada en vocal fonética ? llana por defecto
        If w.count >= 2 Then
            DetectarTonica = w.count - 1
        Else
            DetectarTonica = 1
        End If
    End If

End Function

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
        If pos2 > 1 Then
            secs.Add pos2
        End If
    End If

    Set DetectarSecundarias = secs

End Function

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

