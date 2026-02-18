Attribute VB_Name = "bas_Motor_ES_Main"

Option Compare Database
Option Explicit

' ============================================================
'   MOTOR DE SILABEO 4.1 ó ORTOGR¡FICO + MORFOSIL¡BICO + ACENTUACI”N
'   Autor: Alba Salv· Ripoll + Copilot
'
'   Este mÛdulo implementa:
'   ? Silabeo ortogr·fico completo (RAE 2010)
'   ? Silabeo morfosil·bico (prefijos)
'   ? Diptongos, triptongos, hiatos con tilde
'   ? DÌgrafos indivisibles: ch, ll, rr
'   ? Casos especiales: qu, gu, g¸e/g¸i, tl
'   ? SÌlaba tÛnica
'   ? Tracking de depuraciÛn (modoDebug)
'
'   Variables de control:
'       usarSilabeoMorfologico   ? True = usar prefijos
'       modoPrefijosEstrictos    ? True = usar lista interna completa
'       respetarPrefijos         ? True = separar prefijo como sÌlaba
'       modoDebug                ? True = activar logs
'
' ============================================================

Private usarSilabeoMorfologico As Boolean
Private modoPrefijosEstrictos As Boolean
Private respetarPrefijos As Boolean

' Prefijos personalizados (modo flexible)
Private prefijosPersonalizados As Variant

' Prefijos canÛnicos (modo estricto)
Private prefijosEstrictos As Variant

Private prefijosCargados As Boolean

' ============================================================
'   ENTRADA PRINCIPAL DEL MOTOR (ESPA—OL)
' ============================================================
' ----------------------------------------------------------------
' Procedimiento: Entrada_Motor_ES
' PropÛsito:     Punto de entrada al motor fonÈtico del espaÒol general
' Tipo proc.:    Function
' Acceso proc.:  Public

' Parameter Texto (String): Texto que se recibe (nombre o apellido espaÒol general)

' Tipo retorno: String -> Texto que contiene la lista de fonemas
'   resultado de la conversiÛn

' Autor:        Alba Salv·
' Fecha:        11/02/2026
' ----------------------------------------------------------------
Public Function Entrada_Motor_ES(Texto As String) As String

    Set ObjDTO = New clsDTO_Motor

    ' 1) Asignamos el texto recibido y
    '    NormalizaciÛn (dentro del DTO)
    ObjDTO.TextoOriginal = Texto
    ObjDTO.NormalizaEntrada

    ' 2) Silabeo autom·tico
    Call SilabearFrase
    
    ' 3) Detectar tÛnicas
    Call CalcularTonicas
    
    ' 4) Detectar secundarias
    Call CalcularSecundarias
    
    ' 5) Marcar TÛnicas y Secundarias
    Call MarcarTonicaYSecundariaEnCadena
    
    ObjDTO.SilabasFinal = ObjDTO.SilabasAcentuadas
    
    Call ConstruirCadenaFonemas_ES
    
    Call MF_DebugDTO("Silabear")
    
    ' 3) Devolver resultado
    Entrada_Motor_ES = ObjDTO.SilabasAuto

End Function

'Private Function SilabearFrase(ByVal frase As String) As String
Private Sub SilabearFrase()
    
    Dim frase As String
    Dim palabras() As String
    Dim resultado As String
    Dim i As Long
    Dim limpia As String
    Dim sil As String

    ' Tratamiento prefijos
'       usarSilabeoMorfologico   => True = usar prefijos
    usarSilabeoMorfologico = True

'       modoPrefijosEstrictos    => True = usar lista interna completa
    modoPrefijosEstrictos = True
'       respetarPrefijos         => True = separar prefijo como sÌlaba
    respetarPrefijos = True

    frase = ObjDTO.TextoNormalizado

    If modoDebug Then
'        InitLog
        addLog
        addLog "============================================================="
        addLog "        LOG DE DEPURACI”N SILABEO 5.0"
        addLog "============================================================="
        addLog "Entrada: '" & frase & "'"
'        addLog "Normalizado: '" & T & "'"
        addLog "Longitud: " & Len(frase) & " letras y espacios."
    End If
    
    palabras = Split(frase, " ")

    For i = LBound(palabras) To UBound(palabras)
        'limpia = LimpiarPalabra(palabras(i))
        limpia = Trim$(palabras(i))
        
        If limpia <> "" Then
            sil = SilabearPalabra(limpia)

            ' AÒadir separador de palabra:  |   |
            If resultado = "" Then
                resultado = sil
            Else
                resultado = resultado & " |   | " & sil
            End If
        End If
    Next i

    'SilabearFrase = resultado
    ObjDTO.SilabasAuto = resultado
    
    If modoDebug Then
        addLog
        addLog "Resultado final: " & ObjDTO.SilabasAuto
        addLog "============================================================="
        addLog "                   FIN LOG DE DEPURACI”N"
        addLog "============================================================="
'        PrintLog
    End If
End Sub

'Private Function LimpiarPalabra(ByVal p As String) As String
'    Dim c As String, i As Long
'    Dim r As String
'
'    For i = 1 To Len(p)
'        c = Mid$(p, i, 1)
'        If c Like "[A-Za-z¡…Õ”⁄‹·ÈÌÛ˙¸Ò—]" Then
'            r = r & c
'        End If
'    Next i
'
'    LimpiarPalabra = r
'End Function


Public Function SilabearPalabra(ByVal Texto As String) As String
    Dim t As String
    
    t = LCase$(Trim$(Texto))

    If modoDebug Then
'        InitLog
        addLog
        addLog "============================================================="
'        addLog "        LOG DE DEPURACI”N SILABEO 4.1"
        addLog "  SÕLABAS DE '" & Texto & "'"
        addLog "============================================================="
        addLog "Entrada: '" & Texto & "'"
        addLog "Normalizado: '" & t & "'"
        addLog "Longitud: " & Len(t) & " letras."
    End If

    If usarSilabeoMorfologico Then
        SilabearPalabra = SilabearMorfologico(t)
    Else
        SilabearPalabra = SilabearOrtog(t)
    End If

    If modoDebug Then
        addLog
        addLog "Resultado palabra: " & SilabearPalabra
        addLog "============================================================="
        addLog " FIN SILABEO '" & Texto & "'"
        addLog "============================================================="
        'PrintLog
    End If
End Function

Public Function SilabearOrtog(ByVal t As String) As String
    Dim nucIni() As Byte, nucFin() As Byte
    Dim silIni() As Byte, silFin() As Byte
    Dim nNuc As Byte, i As Byte
    Dim silabas() As String

    If t = "y" Then
        SilabearOrtog = "y"
        Exit Function
    End If

    LocalizarNucleosOrtog t, nucIni, nucFin, nNuc

    ReDim silIni(1 To nNuc)
    ReDim silFin(1 To nNuc)

    CalcularSilabas t, nucIni, nucFin, nNuc, silIni, silFin

    ReDim silabas(1 To nNuc)
    For i = 1 To nNuc
        silabas(i) = Mid$(t, silIni(i), silFin(i) - silIni(i) + 1)
        If modoDebug Then addLog "SÌlaba " & i & ": " & silabas(i)
    Next i

    SilabearOrtog = Join(silabas, " | ")
End Function

Public Function SilabearMorfologico(ByVal t As String) As String
' Prefijos
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

    If modoDebug Then
        addLog "Prefijo detectado: " & pref
        addLog "Resto: " & resto
    End If

    SilabearMorfologico = pref & " | " & SilabearOrtog(resto)
End Function

Private Function DetectarPrefijo(ByVal t As String) As String
    Dim p As Variant

    If Not prefijosCargados Then CargarPrefijos

    For Each p In prefijosEstrictos
    
        ' Si la palabra es exactamente igual al prefijo ? NO separar
        If Len(t) = Len(p) Then
'            GoTo SiguientePrefijo
            Exit For
        End If
    
        If Left$(t, Len(p)) = p Then
            DetectarPrefijo = p
            Exit Function
        End If
    
SiguientePrefijo:
    Next p


'    For Each p In prefijosEstrictos
'        If Left$(T, Len(p)) = p Then
'            DetectarPrefijo = p
'            Exit Function
'        End If
'    Next p

    DetectarPrefijo = ""
End Function

Private Sub CargarPrefijos()
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim i As Long

    If prefijosCargados Then Exit Sub

    sql = "SELECT Prefijo FROM tbmPrefijos " & _
          "WHERE Activo = 1 " & _
            "AND Tipo Like 'autÈntico' " & _
            "AND [es-es] = true " & _
          "ORDER BY Len(Prefijo) DESC, Prefijo ASC"

    Set rs = CurrentDb.OpenRecordset(sql)

    If Not rs.EOF Then
        rs.MoveLast
        ReDim prefijosEstrictos(1 To rs.RecordCount)
        rs.MoveFirst

        i = 1
        Do Until rs.EOF
            prefijosEstrictos(i) = LCase$(rs!prefijo)
            i = i + 1
            rs.MoveNext
        Loop
    End If

    rs.Close
    prefijosCargados = True
End Sub

'Private Function DetectarPrefijo(ByVal T As String) As String
'    Dim p As Variant
'
'    If IsEmpty(prefijosEstrictos) Then
''        prefijosEstrictos = Array("a", "ante", "anti", "auto", "bi", "contra", "de", "des", "dis", _
'                                  "extra", "hiper", "hipo", "in", "im", "inter", "intra", "macro", _
'                                  "micro", "multi", "post", "pre", "pro", "re", "semi", "sub", _
'                                  "super", "trans", "ultra")
' ' Ordenados de mayor a menor longitud para evitar falsos positivos
'        prefijosEstrictos = Array("super", "contra", "extra", "inter", "intra", "macro", "micro", _
'                                  "multi", "trans", "ultra", "ante", "anti", "auto", "post", "pre", _
'                                  "pro", "semi", "sub", "des", "dis", "hiper", "hipo", "in", "im", _
'                                  "re", "de", "a", "bi")
'
'
'    End If
'
'    If modoPrefijosEstrictos Then
'        For Each p In prefijosEstrictos
'            If Left$(T, Len(p)) = p Then
'                DetectarPrefijo = p
'                Exit Function
'            End If
'        Next p
'    End If
'
'    If Not IsEmpty(prefijosPersonalizados) Then
'        For Each p In prefijosPersonalizados
'            If Left$(T, Len(p)) = p Then
'                DetectarPrefijo = p
'                Exit Function
'            End If
'        Next p
'    End If
'
'    DetectarPrefijo = ""
'End Function

Private Sub LocalizarNucleosOrtog(ByVal t As String, _
                                  ByRef nucIni() As Byte, _
                                  ByRef nucFin() As Byte, _
                                  ByRef nNuc As Byte)

    Dim i As Byte, L As Byte
    Dim C1 As String, C2 As String, c3 As String

    L = Len(t)
    ReDim nucIni(1 To L)
    ReDim nucFin(1 To L)
    nNuc = 0

    If modoDebug Then
        addLog
        addLog "---------------------------------------"
        addLog " Procedimiento LocalizarNucleosOrtog"
    End If

    i = 1
    Do While i <= L
        C1 = Mid$(t, i, 1)

        If EsVocal(C1) Then

            ' Intentar triptongo
            If i + 2 <= L Then
                C2 = Mid$(t, i + 1, 1)
                c3 = Mid$(t, i + 2, 1)
                If EsTriptongo(C1, C2, c3) Then
                    nNuc = nNuc + 1
                    nucIni(nNuc) = i
                    nucFin(nNuc) = i + 2
                    If modoDebug Then addLog "Triptongo: " & C1 & C2 & c3
                    i = i + 3
                    GoTo Siguiente
                End If
            End If

            ' Intentar diptongo
            If i + 1 <= L Then
                C2 = Mid$(t, i + 1, 1)
                If EsDiptongo(C1, C2) Then
                    nNuc = nNuc + 1
                    nucIni(nNuc) = i
                    nucFin(nNuc) = i + 1
                    If modoDebug Then addLog "Diptongo: " & C1 & C2
                    i = i + 2
                    GoTo Siguiente
                End If
            End If

            ' Vocal sola
            nNuc = nNuc + 1
            nucIni(nNuc) = i
            nucFin(nNuc) = i
            If modoDebug Then addLog "Vocal sola: " & C1
            i = i + 1

        Else
            i = i + 1
        End If

Siguiente:
    Loop

    If modoDebug Then
        addLog "Total n˙cleos: " & nNuc
        addLog " Fin LocalizarNucleosOrtog"
        addLog "---------------------------------------"
    End If
End Sub

Private Sub CalcularSilabas(ByVal t As String, _
                            ByRef nucIni() As Byte, _
                            ByRef nucFin() As Byte, _
                            ByVal nNuc As Byte, _
                            ByRef silIni() As Byte, _
                            ByRef silFin() As Byte)

' Con control de dÌgrafos

    Dim i As Byte, L As Byte
    Dim a As Byte, b As Byte
    Dim k As Byte
    Dim C1 As String, C2 As String, grupo As String

    L = Len(t)
    silIni(1) = 1

    For i = 1 To nNuc - 1
        a = nucFin(i)
        b = nucIni(i + 1)

        k = IIf(b > a + 1, b - a - 1, 0)

        If modoDebug Then
            addLog
            addLog "---- Frontera entre n˙cleo " & i & " y " & (i + 1)
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
                C1 = Mid$(t, a + 1, 1)
                C2 = Mid$(t, a + 2, 1)
                grupo = C1 & C2

                ' DÌgrafos indivisibles
                If grupo = "rr" Or grupo = "ll" Or grupo = "ch" Then
                    silFin(i) = a
                    silIni(i + 1) = a + 1
                    GoTo Siguiente
                End If

                ' tl ? siempre se separa
                If grupo = "tl" Then
                    silFin(i) = a + 1
                    silIni(i + 1) = a + 2
                    GoTo Siguiente
                End If

                ' Grupos de ataque v·lidos
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

Siguiente:
    Next i

    silFin(nNuc) = L
End Sub

Private Function EsVocal(ByVal c As String) As Boolean
    EsVocal = InStr("aeiou·ÈÌÛ˙¸", c) > 0
End Function

Private Function EsVocalFuerte(ByVal c As String) As Boolean
    EsVocalFuerte = InStr("a·eÈoÛ", c) > 0
End Function

Private Function EsVocalDebil(ByVal c As String) As Boolean
    EsVocalDebil = InStr("iÌu˙¸", c) > 0
End Function

Private Function EsVocalDebilTonica(ByVal c As String) As Boolean
    EsVocalDebilTonica = InStr("Ì˙", c) > 0
End Function

Private Function EsDiptongo(ByVal v1 As String, ByVal v2 As String) As Boolean
    ' Si alguna no es vocal ? no hay diptongo
    If Not EsVocal(v1) Or Not EsVocal(v2) Then Exit Function

    ' Dos vocales fuertes ? nunca diptongo
    If EsVocalFuerte(v1) And EsVocalFuerte(v2) Then Exit Function

    ' *** REGLA ESPECIAL ***
    ' Dos vocales dÈbiles (i, u, ¸) ? SIEMPRE diptongo,
    ' incluso si una lleva tilde (caso g¸Ì de "ling¸Ìstica")
    If EsVocalDebil(v1) And EsVocalDebil(v2) Then
        EsDiptongo = True
        Exit Function
    End If

    ' Si una es dÈbil tÛnica y la otra es fuerte ? hiato
    If (EsVocalDebilTonica(v1) And EsVocalFuerte(v2)) _
    Or (EsVocalFuerte(v1) And EsVocalDebilTonica(v2)) Then
        Exit Function
    End If

    ' En cualquier otro caso ? diptongo
    EsDiptongo = True
End Function

'Private Function EsDiptongo(ByVal v1 As String, ByVal v2 As String) As Boolean
'    If Not EsVocal(v1) Or Not EsVocal(v2) Then Exit Function
'
'    ' Dos fuertes ? nunca diptongo
'    If EsVocalFuerte(v1) And EsVocalFuerte(v2) Then Exit Function
'
'    ' Caso especial: dos dÈbiles (iu / ui), aunque una lleve tilde, SIGUE siendo diptongo
'    If EsVocalDebil(v1) And EsVocalDebil(v2) Then
'        EsDiptongo = True
'        Exit Function
'    End If
'
'    ' Para combinaciones fuerte + dÈbil:
'    ' si la dÈbil es tÛnica ? hiato (no diptongo)
'    If EsVocalDebilTonica(v1) Or EsVocalDebilTonica(v2) Then Exit Function
'
'    ' Resto de casos: diptongo
'    EsDiptongo = True
'End Function

'Private Function EsDiptongo(ByVal v1 As String, ByVal v2 As String) As Boolean
'    If Not EsVocal(v1) Or Not EsVocal(v2) Then Exit Function
'    If EsVocalFuerte(v1) And EsVocalFuerte(v2) Then Exit Function
'    If EsVocalDebilTonica(v1) Or EsVocalDebilTonica(v2) Then Exit Function
'    EsDiptongo = True
'End Function

Private Function EsTriptongo(ByVal v1 As String, ByVal v2 As String, ByVal v3 As String) As Boolean
    If EsVocalDebil(v1) And Not EsVocalDebilTonica(v1) _
       And EsVocalFuerte(v2) _
       And EsVocalDebil(v3) And Not EsVocalDebilTonica(v3) Then
        EsTriptongo = True
    End If
End Function

Private Function EsGrupoAtaque(ByVal g As String) As Boolean
    Dim AC As Variant
    AC = Array("pr", "br", "tr", "dr", "cr", "gr", "fr", _
               "pl", "bl", "cl", "gl", "fl")
    EsGrupoAtaque = (UBound(Filter(AC, g)) >= 0)
End Function

Public Function SilabaTonica(ByVal palabra As String) As Byte
    Dim sil As String
    Dim silabas() As String
    Dim i As Byte

    sil = SilabearPalabra(palabra)
    silabas = Split(sil, " | ")

    For i = 0 To UBound(silabas)
        If TieneTilde(silabas(i)) Then
            SilabaTonica = i + 1
            Exit Function
        End If
    Next i

    If TerminaEnVocalNS(palabra) Then
        SilabaTonica = UBound(silabas)
    Else
        SilabaTonica = UBound(silabas) + 1
    End If
End Function

'Private Function TieneTilde(ByVal s As String) As Boolean
'    TieneTilde = (InStr(s, "·") Or InStr(s, "È") Or InStr(s, "Ì") Or InStr(s, "Û") Or InStr(s, "˙")) > 0
'End Function

Private Function TieneTilde(ByVal silaba As String) As Boolean
    TieneTilde = (InStr(silaba, "·") > 0 Or _
                  InStr(silaba, "È") > 0 Or _
                  InStr(silaba, "Ì") > 0 Or _
                  InStr(silaba, "Û") > 0 Or _
                  InStr(silaba, "˙") > 0)
End Function

'Private Function TieneTilde(sil As String) As Boolean
'
'    Dim acentos As String
'    Dim i As Byte
'
'    acentos = "·ÈÌÛ˙"
'
'    For i = 1 To Len(sil)
'        If InStr(acentos, Mid$(sil, i, 1)) > 0 Then
'            TieneTilde = True
'            Exit Function
'        End If
'    Next i
'
'End Function

Private Function TerminaEnVocalNS(ByVal s As String) As Boolean
    Dim c As String
    c = Right$(s, 1)
    TerminaEnVocalNS = (EsVocal(c) Or c = "n" Or c = "s")
End Function


'=================================================================
'=================================================================
'                 SECCI”N M”DULO FON…TICO (ACENTOS)
'=================================================================
'=================================================================

' ============================================================
'   DETECTAR SÕLABAS T”NICAS
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
            ' palabra real
            Dim w As Collection
            Set w = elementos(i)

            Dim tLocal As Long
            tLocal = DetectarTonica(w)

            If tLocal > 0 Then
                tGlobal.Add globalIndex + tLocal
            End If

            globalIndex = globalIndex + w.count

        Else
            ' HUECO
            globalIndex = globalIndex + 1
        End If

    Next i

    ObjDTO.SilabasTonicas = JoinCollection(tGlobal)

End Sub


' ============================================================
'   DETECTAR SÕLABAS SECUNDARIAS (pueden ser varias)
' ============================================================
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

            ' Detectar tÛnica local
            tLocal = DetectarTonica(w)

            ' Detectar secundarias locales
            Dim secs As Collection
            Set secs = DetectarSecundarias(w, tLocal)

            ' Convertir secundarias locales a globales
            Dim x As Variant
            For Each x In secs
                sGlobal.Add globalIndex + CByte(x)
            Next x

            globalIndex = globalIndex + w.count

        Else
            ' HUECO
            globalIndex = globalIndex + 1
        End If

    Next i

    ObjDTO.SilabasSecundarias = JoinCollection(sGlobal)

End Sub


' ============================================================
'   MARCAR T”NICAS Y SECUNDARIAS EN LA CADENA FINAL
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

    ' 1) T”NICAS
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

    ' 2) SECUNDARIAS
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

    ' 3) UNIR RESULTADO
    ObjDTO.SilabasAcentuadas = Join(out, " | ")

End Sub


'-------------------------------------------------------------
'             AUXILIARES T”NICAS Y SECUNDARIAS
'-------------------------------------------------------------
Private Function DetectarTonica(w As Collection) As Byte

    Dim i As Byte
    Dim ultima As String

    For i = 1 To w.count
        If TieneTilde(w(i)) Then
            DetectarTonica = i
            Exit Function
        End If
    Next i

    ultima = Right$(w(w.count), 1)

    If ultima Like "[aeiouns]" Then
        DetectarTonica = w.count - 1
    Else
        DetectarTonica = w.count
    End If

End Function

Private Function DetectarSecundarias(w As Collection, tPos As Byte) As Collection

    Dim secs As New Collection
    Dim n As Byte
    Dim pos2 As Byte

    n = w.count

    ' Palabras de 1ñ3 sÌlabas ? sin secundaria
    If n < 4 Then
        Set DetectarSecundarias = secs
        Exit Function
    End If

    ' Primera secundaria SIEMPRE en la sÌlaba 1
    secs.Add 1

    ' Palabras de 6+ sÌlabas ? segunda secundaria
    If n >= 6 Then
        pos2 = tPos - 2   ' dos antes de la tÛnica

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
            If palabraActual.count > 0 Then
                resultado.Add palabraActual
                Set palabraActual = New Collection
            End If
            resultado.Add "HUECO"
        Else
            ' Es una sÌlaba real
            palabraActual.Add sils(i)
        End If

    Next i

    ' ⁄ltima palabra
    If palabraActual.count > 0 Then resultado.Add palabraActual

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


