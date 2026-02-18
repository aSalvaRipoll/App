Attribute VB_Name = "MÛdulo10"

Option Compare Database
Option Explicit

' ============================================================
'   MOTOR DE SILABEO 4.1 ó ORTOGR¡FICO + MORFOSIL¡BICO
'   Autor: Alba + Copilot
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
Private modoDebug As Boolean

' Prefijos personalizados (modo flexible)
Private prefijosPersonalizados As Variant

' Prefijos canÛnicos (modo estricto)
Private prefijosEstrictos As Variant

Private prefijosCargados As Boolean

Public Function SilabearFrase(ByVal frase As String) As String
    Dim palabras() As String
    Dim resultado As String
    Dim i As Long
    Dim limpia As String
    Dim sil As String

    ' Activar Logs
    modoDebug = True
    
    ' Tratamiento prefijos
'       usarSilabeoMorfologico   => True = usar prefijos
    usarSilabeoMorfologico = True

'       modoPrefijosEstrictos    => True = usar lista interna completa
    modoPrefijosEstrictos = True
'       respetarPrefijos         => True = separar prefijo como sÌlaba
    respetarPrefijos = True


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

    SilabearFrase = resultado
    
    If modoDebug Then
        addLog
        addLog "Resultado final: " & SilabearFrase
        addLog "============================================================="
        addLog "                   FIN LOG DE DEPURACI”N"
        addLog "============================================================="
'        PrintLog
    End If
End Function

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
    Dim T As String
    
    T = LCase$(Trim$(Texto))

    If modoDebug Then
'        InitLog
        addLog
        addLog "============================================================="
'        addLog "        LOG DE DEPURACI”N SILABEO 4.1"
        addLog "  SÕLABAS DE '" & Texto & "'"
        addLog "============================================================="
        addLog "Entrada: '" & Texto & "'"
        addLog "Normalizado: '" & T & "'"
        addLog "Longitud: " & Len(T) & " letras."
    End If

    If usarSilabeoMorfologico Then
        SilabearPalabra = SilabearMorfologico(T)
    Else
        SilabearPalabra = SilabearOrtog(T)
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

Public Function SilabearOrtog(ByVal T As String) As String
    Dim nucIni() As Byte, nucFin() As Byte
    Dim silIni() As Byte, silFin() As Byte
    Dim nNuc As Byte, i As Byte
    Dim silabas() As String

    If T = "y" Then
        SilabearOrtog = "y"
        Exit Function
    End If

    LocalizarNucleosOrtog T, nucIni, nucFin, nNuc

    ReDim silIni(1 To nNuc)
    ReDim silFin(1 To nNuc)

    CalcularSilabas T, nucIni, nucFin, nNuc, silIni, silFin

    ReDim silabas(1 To nNuc)
    For i = 1 To nNuc
        silabas(i) = Mid$(T, silIni(i), silFin(i) - silIni(i) + 1)
        If modoDebug Then addLog "SÌlaba " & i & ": " & silabas(i)
    Next i

    SilabearOrtog = Join(silabas, " | ")
End Function

Public Function SilabearMorfologico(ByVal T As String) As String
' Prefijos
    Dim pref As String
    Dim resto As String

    If Not respetarPrefijos Then
        SilabearMorfologico = SilabearOrtog(T)
        Exit Function
    End If

    pref = DetectarPrefijo(T)

    If pref = "" Then
        SilabearMorfologico = SilabearOrtog(T)
        Exit Function
    End If

    resto = Mid$(T, Len(pref) + 1)

    If modoDebug Then
        addLog "Prefijo detectado: " & pref
        addLog "Resto: " & resto
    End If

    SilabearMorfologico = pref & " | " & SilabearOrtog(resto)
End Function

Private Function DetectarPrefijo(ByVal T As String) As String
    Dim p As Variant

    If Not prefijosCargados Then CargarPrefijos

    For Each p In prefijosEstrictos
    
        ' Si la palabra es exactamente igual al prefijo ? NO separar
        If Len(T) = Len(p) Then
'            GoTo SiguientePrefijo
            Exit For
        End If
    
        If Left$(T, Len(p)) = p Then
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

Private Sub LocalizarNucleosOrtog(ByVal T As String, _
                                  ByRef nucIni() As Byte, _
                                  ByRef nucFin() As Byte, _
                                  ByRef nNuc As Byte)

    Dim i As Byte, L As Byte
    Dim C1 As String, C2 As String, c3 As String

    L = Len(T)
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
        C1 = Mid$(T, i, 1)

        If EsVocal(C1) Then

            ' Intentar triptongo
            If i + 2 <= L Then
                C2 = Mid$(T, i + 1, 1)
                c3 = Mid$(T, i + 2, 1)
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
                C2 = Mid$(T, i + 1, 1)
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

Private Sub CalcularSilabas(ByVal T As String, _
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

    L = Len(T)
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
                C1 = Mid$(T, a + 1, 1)
                C2 = Mid$(T, a + 2, 1)
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

Private Function TieneTilde(ByVal s As String) As Boolean
    TieneTilde = (InStr(s, "·") Or InStr(s, "È") Or InStr(s, "Ì") Or InStr(s, "Û") Or InStr(s, "˙")) > 0
End Function

Private Function TerminaEnVocalNS(ByVal s As String) As Boolean
    Dim c As String
    c = Right$(s, 1)
    TerminaEnVocalNS = (EsVocal(c) Or c = "n" Or c = "s")
End Function


