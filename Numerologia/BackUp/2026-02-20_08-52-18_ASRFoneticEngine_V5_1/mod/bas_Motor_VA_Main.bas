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

Private prefijosEstrictos_VA As Variant
Private prefijosCargados_VA As Boolean

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
    Call SilabearFrase_VA

    ' 3) Detectar tónicas
    Call CalcularTonicas_VA

    ' 4) Detectar secundarias
    Call CalcularSecundarias_VA

    ' 5) Marcar Tónicas y Secundarias
    Call MarcarTonicaYSecundariaEnCadena_VA

    ObjDTO.SilabasFinal = ObjDTO.SilabasAcentuadas

    ' 6) (cuando tengas fonética VA)
    Call ConstruirCadenaFonemas_VA

    ' Debug opcional
    Call MF_DebugDTO("Silabear_VA")

    Entrada_Motor_VA = ObjDTO.SilabasAuto

End Function

' ============================================================
'   SILABEO DE FRASE
' ============================================================
Private Sub SilabearFrase_VA()

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
        addLog " Procedimiento SilabearFrase_VA"
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
            sil = SilabearPalabra_VA(limpia)
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
        addLog "SilabearFrase_VA ? " & ObjDTO.SilabasAuto
    End If

End Sub

' ============================================================
'   SILABEO DE PALABRA
' ============================================================
Private Function SilabearPalabra_VA(ByVal texto As String) As String
    Dim t As String

    If DebugMotor Then
        addLog
        addLog "---------------------------------------"
        addLog " Procedimiento SilabearPalabra_VA"
        addLog
        addLog "texto: " & texto
        
    End If

    t = LCase$(Trim$(texto))

    If usarSilabeoMorfologico Then
        SilabearPalabra_VA = SilabearMorfologico_VA(t)
    Else
        SilabearPalabra_VA = SilabearOrtog_VA(t)
    End If

End Function

' ============================================================
'   SILABEO ORTOGRÁFICO VA
' ============================================================
Private Function SilabearOrtog_VA(ByVal t As String) As String
    Dim nucIni() As Byte, nucFin() As Byte
    Dim silIni() As Byte, silFin() As Byte
    Dim nNuc As Byte, i As Byte
    Dim silabas() As String

    If DebugMotor Then
        addLog
        addLog "---------------------------------------"
        addLog " Procedimiento SilabearOrtog_VA"
        addLog
        addLog "t: " & t
        
    End If


    If Len(Trim$(t)) < 2 Then
        If t = "y" Then
            SilabearOrtog_VA = "y"
            Exit Function
        ElseIf t = "i" Then
            SilabearOrtog_VA = "i"
            Exit Function
        End If
    End If

    LocalizarNucleosOrtog_VA t, nucIni, nucFin, nNuc

    ReDim silIni(1 To nNuc)
    ReDim silFin(1 To nNuc)

    CalcularSilabas_VA t, nucIni, nucFin, nNuc, silIni, silFin

    ReDim silabas(1 To nNuc)
    For i = 1 To nNuc
        silabas(i) = Mid$(t, silIni(i), silFin(i) - silIni(i) + 1)
        If DebugMotor Then
            addLog "VA — Sílaba " & i & ": " & silabas(i)
        End If
    Next i

    SilabearOrtog_VA = Join(silabas, " | ")

End Function

' ============================================================
'   SILABEO MORFOSILÁBICO VA (PREFIJOS)
' ============================================================
Private Function SilabearMorfologico_VA(ByVal t As String) As String
    Dim pref As String
    Dim resto As String

    If DebugMotor Then
        addLog
        addLog "---------------------------------------"
        addLog " Procedimiento SilabearMorfologico_VA"
        addLog
        addLog "t: " & t
        
    End If
    
    If Not respetarPrefijos Then
        SilabearMorfologico_VA = SilabearOrtog_VA(t)
        Exit Function
    End If

    pref = DetectarPrefijo_VA(t)

    addLog
    addLog "pref: " & pref
    
    If pref = "" Then
        SilabearMorfologico_VA = SilabearOrtog_VA(t)
        Exit Function
    End If

    resto = Mid$(t, Len(pref) + 1)

    SilabearMorfologico_VA = pref & " | " & SilabearOrtog_VA(resto)
End Function

Private Function DetectarPrefijo_VA(ByVal t As String) As String
    Dim p As Variant
    Dim resto As String
    Dim primera As String, segunda As String

    If Not prefijosCargados_VA Then CargarPrefijos_VA

    For Each p In prefijosEstrictos_VA

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
            If Not EsVocal_VA(primera) Then
                GoTo SiguienteP
            End If

            ' 2.2) No puede romper un diptongo
            If EsDiptongo_VA(Left$(t, 1), Mid$(t, 2, 1)) Then
                GoTo SiguienteP
            End If

            ' 2.3) El ataque resultante debe ser válido
            If Not EsGrupoAtaque_VA(primera & segunda) Then
                GoTo SiguienteP
            End If
            
            DetectarPrefijo_VA = p
            Exit Function
        End If

        ' ============================================================
        '   PREFIJOS NORMALES (anti-, inter-, sub-, re-, etc.)
        ' ============================================================

        ' 3) El ataque resultante debe ser válido
        If Not AtaqueSilabicoValido(primera, segunda) Then
            GoTo SiguienteP
        End If

        DetectarPrefijo_VA = p
        Exit Function

SiguienteP:
    Next p

    DetectarPrefijo_VA = ""
End Function

'Private Function DetectarPrefijo_VA(ByVal t As String) As String
'    Dim p As Variant
'
'    If DebugMotor Then
'        addLog
'        addLog "---------------------------------------"
'        addLog " Procedimiento DetectarPrefijo_VA"
'        addLog
'        addLog "t: " & t
'
'    End If
'
'    If Not prefijosCargados_VA Then CargarPrefijos_VA
'
'    For Each p In prefijosEstrictos_VA
'        If Len(t) = Len(p) Then Exit For
'        If Left$(t, Len(p)) = p Then
'            DetectarPrefijo_VA = p
'            Exit Function
'        End If
'    Next p
'
'    DetectarPrefijo_VA = ""
'End Function

Private Sub CargarPrefijos_VA()
    Dim rs As dao.Recordset
    Dim sql As String
    Dim i As Long

    If prefijosCargados_VA Then Exit Sub

    sql = strSQL

    Set rs = CurrentDb.OpenRecordset(sql)

    If Not rs.EOF Then
        rs.MoveLast
        ReDim prefijosEstrictos_VA(1 To rs.RecordCount)
        rs.MoveFirst

        i = 1
        Do Until rs.EOF
            prefijosEstrictos_VA(i) = LCase$(rs!prefijo)
            i = i + 1
            rs.MoveNext
        Loop
    End If

    rs.Close
    prefijosCargados_VA = True
End Sub

' ============================================================
'   LOCALIZAR NÚCLEOS VOCÁLICOS VA
' ============================================================
Private Sub LocalizarNucleosOrtog_VA(ByVal t As String, _
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
        addLog " Procedimiento LocalizarNucleosOrtog_VA"
        addLog
        addLog "t: " & t
        addLog "nucIni: " & CStr(nucIni)
        addLog "nucFin: " & CStr(nucFin)
        addLog "nNuc: " & CStr(nNuc)
        
    End If

    i = 1
    Do While i <= L
        c1 = Mid$(t, i, 1)

        If EsVocal_VA(c1) Then

            ' Triptongo
            If i + 2 <= L Then
                c2 = Mid$(t, i + 1, 1)
                c3 = Mid$(t, i + 2, 1)
                If EsTriptongo_VA(c1, c2, c3) Then
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
                If EsDiptongo_VA(c1, c2) Then
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
        addLog " Fin LocalizarNucleosOrtog_VA"
        addLog "---------------------------------------"
    End If
End Sub

' ============================================================
'   CÁLCULO DE SÍLABAS — VA
' ============================================================
Private Sub CalcularSilabas_VA(ByVal t As String, _
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
                If EsGrupoAtaque_VA(grupo) Then
                    silFin(i) = a
                    silIni(i + 1) = a + 1
                Else
                    ' C1 se queda en la coda
                    silFin(i) = a + 1
                    silIni(i + 1) = a + 2
                End If

'            Case 2
'                c1 = Mid$(t, a + 1, 1)
'                c2 = Mid$(t, a + 2, 1)
'                grupo = c1 & c2
'
'                ' Dígrafos indivisibles
'                If grupo = "rr" Or grupo = "ll" Or grupo = "ch" Or grupo = "l·l" Then
'                    silFin(i) = a
'                    silIni(i + 1) = a + 1
'                    GoTo Siguiente
'                End If
'
'                ' Grupos de ataque válidos
'                If EsGrupoAtaque_VA(grupo) Then
'                    silFin(i) = a
'                    silIni(i + 1) = a + 1
'                Else
'                    silFin(i) = a + 1
'                    silIni(i + 1) = a + 2
'                End If

'            Case 2
'                c1 = Mid$(t, a + 1, 1)
'                c2 = Mid$(t, a + 2, 1)
'                grupo = c1 & c2
'
'                ' Dígrafos indivisibles (puedes ampliar si quieres)
'                If grupo = "rr" Or grupo = "ll" Or grupo = "ch" Or grupo = "l·l" Then
'                    silFin(i) = a
'                    silIni(i + 1) = a + 1
'                    GoTo Siguiente
'                End If
'
'                ' Grupos de ataque válidos
'                If EsGrupoAtaque_VA(grupo) Then
'                    silFin(i) = a
'                    silIni(i + 1) = a + 1
'                Else
'                    silFin(i) = a + 1
'                    silIni(i + 1) = a + 2
'                End If

            Case 3
                c1 = Mid$(t, a + 1, 1)
                c2 = Mid$(t, a + 2, 1)
                c3 = Mid$(t, a + 3, 1)
            
                ' Si C2C3 es ataque válido ? VC1 | C2C3V
                If EsGrupoAtaque_VA(c2 & c3) Or EsConsonanteSimple(c2) Then
                    silFin(i) = a + 1
                    silIni(i + 1) = a + 2
                Else
                    ' Si no ? VC1C2 | C3V
                    silFin(i) = a + 2
                    silIni(i + 1) = a + 3
                End If

'            Case 3
'                c1 = Mid$(t, a + 1, 1)
'                c2 = Mid$(t, a + 2, 1)
'                c3 = Mid$(t, a + 3, 1)
'
'                ' Si C2C3 es ataque válido ? VC1 | C2C3V
'                If EsGrupoAtaque_VA(c2 & c3) Then
'                    silFin(i) = a + 1
'                    silIni(i + 1) = a + 2
'                Else
'                    ' Si no ? VC1C2 | C3V
'                    silFin(i) = a + 2
'                    silIni(i + 1) = a + 3
'                End If

'            Case 3
'                c1 = Mid$(t, a + 1, 1)
'                c2 = Mid$(t, a + 2, 1)
'                c3 = Mid$(t, a + 3, 1)
'
'                If PuedeCerrarSilaba_VA(c2) Then
'                    silFin(i) = a + 2
'                    silIni(i + 1) = a + 3
'                Else
'                    silFin(i) = a + 1
'                    silIni(i + 1) = a + 2
'                End If

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

'Private Function ValidarAtaqueValenciano(ByVal ataque As String) As Boolean
'    ataque = LCase$(ataque)
'
'    ' Ataques simples siempre válidos
'    If Len(ataque) = 1 Then
'        ValidarAtaqueValenciano = True
'        Exit Function
'    End If
'
'    ' Ataques complejos válidos en valenciano
'    Select Case ataque
'        Case "pl", "bl", "cl", "gl", _
'             "pr", "br", "tr", "dr", "kr", "gr", _
'             "fl", "fr", _
'             "sp", "st", "sk"
'            ValidarAtaqueValenciano = True
'        Case Else
'            ValidarAtaqueValenciano = False
'    End Select
'End Function


' ============================================================
'   FUNCIONES DE VOCAL — VALENCIANO
' ============================================================
Private Function EsVocal_VA(ByVal c As String) As Boolean

    addLog
    addLog "---------------------------------------"
    addLog " Procedimiento EsVocal_VA"
    addLog "---------------------------------------"
    addLog "c: " & c
    
    EsVocal_VA = InStr("aeiouàèéíòóú", LCase$(c)) > 0
    addLog "EsVocal_VA: " & EsVocal_VA
    
End Function

Private Function EsVocalForta_VA(ByVal c As String) As Boolean
    
    addLog
    addLog "---------------------------------------"
    addLog " Procedimiento EsVocalForta_VA"
    addLog "---------------------------------------"
    addLog "c: " & c
    
    EsVocalForta_VA = InStr("aàeèéoò", LCase$(c)) > 0
    addLog "EsVocalForta_VA: " & EsVocalForta_VA
    
End Function

Private Function EsVocalDebil_VA(ByVal c As String) As Boolean
    
    addLog
    addLog "---------------------------------------"
    addLog " Procedimiento EsVocalDebil_VA"
    addLog "---------------------------------------"
    addLog "c: " & c
    
    EsVocalDebil_VA = InStr("iíuú", LCase$(c)) > 0
    addLog "EsVocalDebil_VA: " & EsVocalDebil_VA
    
End Function

Private Function EsVocalDebilTonica_VA(ByVal c As String) As Boolean
    
    addLog
    addLog "---------------------------------------"
    addLog " Procedimiento EsVocalDebilTonica_VA"
    addLog "---------------------------------------"
    addLog "c: " & c
    
    EsVocalDebilTonica_VA = InStr("íú", LCase$(c)) > 0
    addLog "EsVocalDebilTonica_VA: " & EsVocalDebilTonica_VA
End Function

Private Function EsDiptongo_VA(v1 As String, v2 As String) As Boolean
    v1 = LCase$(v1)
    v2 = LCase$(v2)

    addLog
    addLog "-----------------------------------------------------------"
    addLog "Función EsDiptongo_VA:  v1 => " & v1 & "  |  v2: => " & v2
    addLog "-----------------------------------------------------------"
    
    EsDiptongo_VA = False
    
    Select Case v1 & v2
        ' Diptongos crecientes
        Case "ia", "ie", "io", "ua", "ue", "uo"
            EsDiptongo_VA = True
            'addLog "EsDiptongo_VA: " & EsDiptongo_VA
            'Exit Function

        ' Diptongos decrecientes
        Case "ai", "ei", "oi", "au", "eu", "ou"
            EsDiptongo_VA = True
            'addLog "EsDiptongo_VA: " & EsDiptongo_VA
            'Exit Function

        ' Semivocales
        Case "iu", "ui"
            EsDiptongo_VA = True
            'addLog "EsDiptongo_VA: " & EsDiptongo_VA
            'Exit Function
    End Select
    
    addLog "EsDiptongo_VA: " & EsDiptongo_VA
    
End Function

'Private Function EsDiptongo_VA(v1 As String, v2 As String) As Boolean
'    v1 = LCase$(v1)
'    v2 = LCase$(v2)
'
'    Select Case v1 & v2
'        ' Diptongos crecientes
'        Case "ia", "ie", "io", "ua", "ue", "uo"
'            EsDiptongo_VA = True
'            Exit Function
'
'        ' Diptongos decrecientes
'        Case "ai", "ei", "oi", "au", "eu", "ou"
'            EsDiptongo_VA = True
'            Exit Function
'
'        ' Semivocales
'        Case "iu", "ui"
'            EsDiptongo_VA = True
'            Exit Function
'    End Select
'
'    EsDiptongo_VA = False
'End Function

'Private Function EsDiptongo_VA(ByVal v1 As String, ByVal v2 As String) As Boolean
'
'    If Not EsVocal_VA(v1) Or Not EsVocal_VA(v2) Then Exit Function
'
'    ' Dos fuertes ? hiato
'    If EsVocalForta_VA(v1) And EsVocalForta_VA(v2) Then Exit Function
'
'    ' Dos débiles ? diptongo
'    If EsVocalDebil_VA(v1) And EsVocalDebil_VA(v2) Then
'        EsDiptongo_VA = True
'        Exit Function
'    End If
'
'    ' Débil tónica + fuerte ? hiato
'    If (EsVocalDebilTonica_VA(v1) And EsVocalForta_VA(v2)) _
'    Or (EsVocalForta_VA(v1) And EsVocalDebilTonica_VA(v2)) Then
'        Exit Function
'    End If
'
'    ' Resto ? diptongo
'    EsDiptongo_VA = True
'
'End Function

Private Function EsTriptongo_VA(ByVal v1 As String, ByVal v2 As String, ByVal v3 As String) As Boolean

    addLog
    addLog "-----------------------------------------------------------------------------"
    addLog "Función EsTriptongo_VA:  v1 => " & v1 & " | v2: => " & v2 & " | v2: => " & v3
    addLog "-----------------------------------------------------------------------------"
    
    If EsVocalDebil_VA(v1) And Not EsVocalDebilTonica_VA(v1) _
       And EsVocalForta_VA(v2) _
       And EsVocalDebil_VA(v3) And Not EsVocalDebilTonica_VA(v3) Then
        EsTriptongo_VA = True
    End If
    
    addLog "EsTriptongo_VA: " & EsTriptongo_VA
    
End Function

Private Function EsGrupoAtaque_VA(ByVal grupo As String) As Boolean
    grupo = LCase$(grupo)

    Select Case grupo
        ' Oclusiva + líquida
        Case "pl", "bl", "cl", "gl", _
             "pr", "br", "tr", "dr", "kr", "gr"
            EsGrupoAtaque_VA = True
            Exit Function

        ' Fricativa + líquida
        Case "fl", "fr"
            EsGrupoAtaque_VA = True
            Exit Function

        ' S + consonante (limitado)
        Case "sp", "st", "sk"
            EsGrupoAtaque_VA = True
            Exit Function

        ' Todo lo demás NO es ataque válido
        Case Else
            EsGrupoAtaque_VA = False
    End Select
End Function

'Private Function EsGrupoAtaque_VA(ByVal grupo As String) As Boolean
'    grupo = LCase$(grupo)
'
'    Select Case grupo
'        ' Oclusiva + líquida
'        Case "pl", "bl", "cl", "gl", _
'             "pr", "br", "tr", "dr", "kr", "gr"
'
'            EsGrupoAtaque_VA = True
'            Exit Function
'
'        ' Fricativa + líquida
'        Case "fl", "fr"
'            EsGrupoAtaque_VA = True
'            Exit Function
'
'        ' S + consonante (limitado)
'        Case "sp", "st", "sk"
'            EsGrupoAtaque_VA = True
'            Exit Function
'
'        ' Todo lo demás NO es ataque válido
'        Case Else
'            EsGrupoAtaque_VA = False
'    End Select
'End Function

'Private Function EsGrupoAtaque_VA(ByVal g As String) As Boolean
'    Dim AC As Variant
'    AC = Array("pr", "br", "tr", "dr", "cr", "gr", "fr", _
'               "pl", "bl", "cl", "gl", "fl")
'    EsGrupoAtaque_VA = (UBound(Filter(AC, g)) >= 0)
'End Function

Private Function PuedeCerrarSilaba_VA(ByVal c As String) As Boolean
    PuedeCerrarSilaba_VA = Not (c = "r" Or c = "l" Or c = "h")
End Function

Private Function TieneTilde_VA(ByVal silaba As String) As Boolean
    TieneTilde_VA = (InStr(silaba, "à") > 0 Or _
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

Private Sub CalcularTonicas_VA()

    Dim tGlobal As New Collection
    Dim elementos As Collection
    Set elementos = ObtenerPalabrasDesdeSilabasAuto_VA()

    Dim globalIndex As Byte
    Dim i As Byte

    globalIndex = 0

    For i = 1 To elementos.count

        If TypeName(elementos(i)) = "Collection" Then
            Dim w As Collection
            Set w = elementos(i)

            Dim tLocal As Long
            tLocal = DetectarTonica_VA(w)

            If tLocal > 0 Then
                tGlobal.Add globalIndex + tLocal
            End If

            globalIndex = globalIndex + w.count

        Else
            globalIndex = globalIndex + 1
        End If

    Next i

    ObjDTO.SilabasTonicas = JoinCollection_VA(tGlobal)

End Sub

Private Sub CalcularSecundarias_VA()

    Dim sGlobal As New Collection
    Dim elementos As Collection
    Set elementos = ObtenerPalabrasDesdeSilabasAuto_VA()

    Dim globalIndex As Byte
    Dim i As Byte
    Dim tLocal As Byte

    globalIndex = 0

    For i = 1 To elementos.count

        If TypeName(elementos(i)) = "Collection" Then

            Dim w As Collection
            Set w = elementos(i)

            tLocal = DetectarTonica_VA(w)

            Dim secs As Collection
            Set secs = DetectarSecundarias_VA(w, tLocal)

            Dim x As Variant
            For Each x In secs
                sGlobal.Add globalIndex + CByte(x)
            Next x

            globalIndex = globalIndex + w.count

        Else
            globalIndex = globalIndex + 1
        End If

    Next i

    ObjDTO.SilabasSecundarias = JoinCollection_VA(sGlobal)

End Sub

Private Sub MarcarTonicaYSecundariaEnCadena_VA()

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

Private Function DetectarTonica_VA(w As Collection) As Byte

    Dim palabra As String
    Dim ultimaSilaba As String
    Dim grafFinal As String
    Dim i As Byte
    Dim esConsonanteFonetica As Boolean
    Dim tieneTilde As Boolean

    ' 1) Reconstruir palabra
    palabra = ""
    For i = 1 To w.count
        palabra = palabra & w(i)
    Next i

    ultimaSilaba = w(w.count)

    ' 2) Detectar tilde explícita
    For i = 1 To w.count
        If TieneTilde_VA(w(i)) Then
            DetectarTonica_VA = i
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
            DetectarTonica_VA = w.count - 1
        Else
            DetectarTonica_VA = 1
        End If

    Else
        ' Palabra acabada en vocal fonética ? llana por defecto
        If w.count >= 2 Then
            DetectarTonica_VA = w.count - 1
        Else
            DetectarTonica_VA = 1
        End If
    End If

End Function

'Private Function DetectarTonica_VA(w As Collection) As Byte
'
'    Dim ultima As String
'    Dim palabra As String
'    Dim esAguda As Boolean
'    Dim i As Byte
'
'    palabra = ""
'
'    ' 1) Tilde gráfica
'    For i = 1 To w.count
'        If TieneTilde_VA(w(i)) Then
'            DetectarTonica_VA = i
'            Exit Function
'        End If
'    Next i
'
'    ' 2) Sin tilde: regla general valenciana
'    For i = 1 To w.count
'        palabra = palabra & w(i)
'    Next i
'
'    ultima = Right$(palabra, 1)
'    esAguda = False
'
'    If InStr("aeiouàèéíòóú", ultima) > 0 Then esAguda = True
'    If LCase$(Right$(palabra, 2)) Like "*as" Then esAguda = True
'    If LCase$(Right$(palabra, 2)) Like "*es" Then esAguda = True
'    If LCase$(Right$(palabra, 2)) Like "*is" Then esAguda = True
'    If LCase$(Right$(palabra, 2)) Like "*os" Then esAguda = True
'    If LCase$(Right$(palabra, 2)) Like "*us" Then esAguda = True
'    If LCase$(Right$(palabra, 2)) = "en" Then esAguda = True
'    If LCase$(Right$(palabra, 2)) = "in" Then esAguda = True
'
'    If esAguda Then
'        DetectarTonica_VA = w.count
'    ElseIf w.count >= 2 Then
'        DetectarTonica_VA = w.count - 1
'    Else
'        DetectarTonica_VA = w.count
'    End If
'
'End Function

Private Function DetectarSecundarias_VA(w As Collection, tPos As Byte) As Collection

    Dim secs As New Collection
    Dim n As Byte
    Dim pos2 As Byte

    n = w.count

    If n < 4 Then
        Set DetectarSecundarias_VA = secs
        Exit Function
    End If

    secs.Add 1

    If n >= 6 Then
        pos2 = tPos - 2
        If pos2 > 1 Then
            secs.Add pos2
        End If
    End If

    Set DetectarSecundarias_VA = secs

End Function

Private Function ObtenerPalabrasDesdeSilabasAuto_VA() As Collection

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

    Set ObtenerPalabrasDesdeSilabasAuto_VA = resultado

End Function

Private Function JoinCollection_VA(col As Collection) As String

    Dim arr() As String
    Dim i As Byte

    If col Is Nothing Then
        JoinCollection_VA = ""
        Exit Function
    End If

    If col.count = 0 Then
        JoinCollection_VA = ""
        Exit Function
    End If

    ReDim arr(1 To col.count)

    For i = 1 To col.count
        arr(i) = CStr(col(i))
    Next i

    JoinCollection_VA = Join(arr, ",")

End Function


