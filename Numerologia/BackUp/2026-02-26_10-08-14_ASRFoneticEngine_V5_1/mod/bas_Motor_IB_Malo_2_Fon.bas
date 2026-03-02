Attribute VB_Name = "bas_Motor_IB_Malo_2_Fon"

'Option Compare Database
'Option Explicit
'
'' ============================================================
''   MOTOR FONÈTIC — BALEAR (IB) V 3.1
''   Construcció de IdsFonemes + IPA final
''   Arquitectura paral·lela al motor CA
'' ============================================================
'
'Public Sub ConstruirCadenaFonemas_IB()
'
'    Dim arrSilabas As Variant
'    Dim i As Long
'    Dim silabaCruda As String
'    Dim ultimaSilaba As String
'    Dim siguienteSilaba As String
'    Dim frase As String
'    Dim arrFon As Variant
'    Dim Fon As Variant
'    Dim res As String
'
'    ObjDTO.IdsFonemas = ""
'    ObjDTO.FonemasFinal = ""
'
'    ' ---------------------------------------------------------
'    ' 0. Normalització (IB no toca grafies alienes)
'    ' ---------------------------------------------------------
'    frase = NormalizarVocales(ObjDTO.SilabasFinal)
'
'    ' ---------------------------------------------------------
'    ' 1. Separar síl·labes per "|"
'    ' ---------------------------------------------------------
'    arrSilabas = Split(frase, "|")
'
'    For i = 0 To UBound(arrSilabas)
'
'        silabaCruda = arrSilabas(i)
'
'        ' ---------------------------------------------------------
'        ' 2. Separador de paraula (síl·laba buida)
'        ' ---------------------------------------------------------
'        If Trim$(silabaCruda) = "" Then
'
'            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "#"
'
'            ' Ligadura automàtica (Mode 2)
'            ultimaSilaba = BuscarSilabaRealAnterior(arrSilabas, i)
'            siguienteSilaba = BuscarSilabaRealPosterior(arrSilabas, i)
'
'            If HayLigaduraAutomatica(ultimaSilaba, siguienteSilaba) Then
'                ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "84,"
'            End If
'
'            GoTo SiguienteIteracion
'        End If
'
'        ' ---------------------------------------------------------
'        ' 3. Modificadors prosòdics (acento)
'        ' ---------------------------------------------------------
'        If InStr(silabaCruda, "(") > 0 Then
'            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "80,"   ' tònica
'        ElseIf InStr(silabaCruda, "[") > 0 Then
'            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "81,"   ' secundària
'        Else
'            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "82,"   ' àtona
'        End If
'
'        ' ---------------------------------------------------------
'        ' 4. Processar grafemes (IB)
'        ' ---------------------------------------------------------
'        ProcesarSilaba silabaCruda
'
'        ' ---------------------------------------------------------
'        ' 5. Separador sil·làbic
'        ' ---------------------------------------------------------
'        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "|"
'
'SiguienteIteracion:
'    Next i
'
'    ' ---------------------------------------------------------
'    ' 6. Neteja final (igual que CA)
'    ' ---------------------------------------------------------
'    ObjDTO.IdsFonemas = Replace(ObjDTO.IdsFonemas, "|#", "#")
'    ObjDTO.IdsFonemas = Replace(ObjDTO.IdsFonemas, "#|", "#")
'    ObjDTO.IdsFonemas = Replace(ObjDTO.IdsFonemas, ",|", "|")
'    ObjDTO.IdsFonemas = Replace(ObjDTO.IdsFonemas, "84,84,", "84,")
'
'    While Right$(ObjDTO.IdsFonemas, 1) = "|" Or _
'          Right$(ObjDTO.IdsFonemas, 1) = "#" Or _
'          Right$(ObjDTO.IdsFonemas, 1) = ","
'
'        ObjDTO.IdsFonemas = Left$(ObjDTO.IdsFonemas, Len(ObjDTO.IdsFonemas) - 1)
'    Wend
'
'    ' ---------------------------------------------------------
'    ' 7. Construcció IPA final (mateix esquema que CA)
'    ' ---------------------------------------------------------
'    arrSilabas = Split(ObjDTO.IdsFonemas, "|")
'    res = ""
'
'    For i = 0 To UBound(arrSilabas)
'
'        If i = 0 Then res = res & "/"
'
'        arrFon = Split(arrSilabas(i), ",")
'
'        For Each Fon In arrFon
'
'            If Fon = "#84" Then
'                Fon = 84
'            ElseIf Fon = "#82" Then
'                res = res & "/ /"
'                Fon = 82
'            End If
'
'            If Left$(Fon, 1) = "#" Then
'                res = res & "/ /"
'                Fon = Replace(Fon, "#", "")
'            End If
'
'            If Trim$(Fon) <> "" And Trim$(Fon) <> "82" Then
'                res = res & Replace(ObtenerIPA(Fon), "/", "")
'            End If
'
'        Next Fon
'
'        If i = UBound(arrSilabas) Then
'            res = res & "/ "
'        End If
'
'    Next i
'
'    ObjDTO.FonemasFinal = Trim$(res)
'
'End Sub
'
'' ============================================================
''   PROCESSAR SÍL·LABA — BALEAR (IB)  (VERSIÓ RECONSTRUÏDA)
'' ============================================================
'Private Sub ProcesarSilaba(ByVal silabaCruda As String)
'
'    Dim silaba As String
'    Dim ligaduraID As Byte
'    Dim i As Long
'    Dim grafema As String
'    Dim id As Byte
'    Dim antCh As String
'    Dim sigCh As String
'    Dim esAtona As Boolean
'
'    ' ---------------------------------------------------------
'    ' 0. Detectar si la síl·laba és àtona
'    ' ---------------------------------------------------------
'    esAtona = True
'    If InStr(silabaCruda, "(") > 0 Then esAtona = False
'    If InStr(silabaCruda, "[") > 0 Then esAtona = False
'
'    ' ---------------------------------------------------------
'    ' 1. Netejar marcadors
'    ' ---------------------------------------------------------
'    silaba = silabaCruda
'    silaba = Replace(silaba, "(", "")
'    silaba = Replace(silaba, ")", "")
'    silaba = Replace(silaba, "[", "")
'    silaba = Replace(silaba, "]", "")
'    silaba = Trim$(silaba)
'
'    ' ---------------------------------------------------------
'    ' 2. Ligadura manual (Mode 1)
'    ' ---------------------------------------------------------
'    ligaduraID = DetectarLigaduraManual(silaba)
'    If ligaduraID <> 0 Then
'        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & CStr(ligaduraID) & ","
'    End If
'
'    ' ---------------------------------------------------------
'    ' 3. Processar grafemes IB
'    ' ---------------------------------------------------------
'    i = 1
'    Do While i <= Len(silaba)
'
'        grafema = DetectarDigrafo(silaba, i)
'
'        ' Context anterior
'        If i > 1 Then
'            antCh = Mid$(silaba, i - 1, 1)
'        Else
'            antCh = ""
'        End If
'
'        ' Context següent
'        If i + Len(grafema) <= Len(silaba) Then
'            sigCh = Mid$(silaba, i + Len(grafema), 1)
'        Else
'            sigCh = ""
'        End If
'
'        ' ---------------------------------------------------------
'        ' 4. SCHWA IB (només proclítics)
'        ' ---------------------------------------------------------
'        If esAtona Then
'            If EsProclitic(silabaCruda) Then
'                If grafema = "a" Or grafema = "e" Then
'                    ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "8,"
'                    i = i + Len(grafema)
'                    GoTo SegGraf
'                End If
'            End If
'        End If
'
'        ' ---------------------------------------------------------
'        ' 5. CASOS ESPECIALS MALLORQUINS
'        ' ---------------------------------------------------------
'
'        ' --- GU + a/o/u ? /g/ + /w/ ---
'        If grafema = "gu" Then
'            If sigCh Like "[aouàòóú]" Then
'                ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "35,22,"
'                i = i + 2
'                GoTo SegGraf
'            End If
'        End If
'
'        ' --- GÜ + e/i ? /g/ + /w/ ---
'        If grafema = "gü" Then
'            If sigCh Like "[eéiíè]" Then
'                ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "35,22,"
'                i = i + 2
'                GoTo SegGraf
'            End If
'        End If
'
'        ' --- QU + e/i ? /k/ ---
'        If grafema = "qu" Then
'            If sigCh Like "[eéiíè]" Then
'                ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "34,"
'                i = i + 2
'                GoTo SegGraf
'            End If
'        End If
'
'        ' --- QÜ + e/i ? /k/ + /w/ ---
'        If grafema = "qü" Then
'            If sigCh Like "[eéiíè]" Then
'                ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "34,22,"
'                i = i + 3
'                GoTo SegGraf
'            End If
'        End If
'
'        ' --- X intervocàlica ? /ks/ ---
'        If grafema = "x" Then
'            If antCh Like "[aeiouàèéíïòóúü]" And sigCh Like "[aeiouàèéíïòóúü]" Then
'                ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "34,42,"
'                i = i + 1
'                GoTo SegGraf
'            End If
'        End If
'
'        ' --- EX + vocal ? /eks/ ---
'        If grafema = "x" Then
'            If antCh = "e" And sigCh Like "[aeiouàèéíïòóúü]" Then
'                ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "34,42,"
'                i = i + 1
'                GoTo SegGraf
'            End If
'        End If
'
'        ' --- IX ? /?/ ---
'        If grafema = "ix" Then
'            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "44,"
'            i = i + 2
'            GoTo SegGraf
'        End If
'
'        ' --- TX ? /t??/ ---
'        If grafema = "tx" Then
'            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "57,"
'            i = i + 2
'            GoTo SegGraf
'        End If
'
'        ' --- TG / TJ ? /d??/ ---
'        If grafema = "tg" Or grafema = "tj" Then
'            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "58,"
'            i = i + 2
'            GoTo SegGraf
'        End If
'
'        ' --- TS / TZ ? /ts/ ---
'        If grafema = "ts" Or grafema = "tz" Then
'            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "46,"
'            i = i + 2
'            GoTo SegGraf
'        End If
'
'        ' ---------------------------------------------------------
'        ' 6. FONEMA BASE (CAS GENERAL)
'        ' ---------------------------------------------------------
'        id = AsignarFonemaBase(grafema, sigCh, antCh)
'        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & CStr(id) & ","
'        i = i + Len(grafema)
'
'SegGraf:
'    Loop
'
'End Sub
'
'' ============================================================
''   FONEMA BASE — BALEAR (IB)  (VERSIÓ RECONSTRUÏDA)
''   Només fonemes simples. Cap fonema doble aquí.
'' ============================================================
'Private Function AsignarFonemaBase(grafema As String, _
'                                   Optional sig As String = "", _
'                                   Optional ant As String = "") As Byte
'
'    grafema = LCase$(grafema)
'    sig = LCase$(sig)
'    ant = LCase$(ant)
'
'    ' ---------------------------------------------------------
'    ' VOCALS IB
'    ' ---------------------------------------------------------
'    Select Case grafema
'        Case "a", "à", "á": AsignarFonemaBase = 1: Exit Function
'        Case "e", "é": AsignarFonemaBase = 2: Exit Function
'        Case "è": AsignarFonemaBase = 3: Exit Function
'        Case "i", "í", "ï": AsignarFonemaBase = 4: Exit Function
'        Case "o", "ó": AsignarFonemaBase = 5: Exit Function
'        Case "ò": AsignarFonemaBase = 6: Exit Function
'        Case "u", "ú", "ü": AsignarFonemaBase = 7: Exit Function
'    End Select
'
'    ' ---------------------------------------------------------
'    ' SEMIVOCALS
'    ' ---------------------------------------------------------
'    If grafema = "j" Or grafema = "y" Then
'        AsignarFonemaBase = 21   ' /j/
'        Exit Function
'    End If
'
'    If grafema = "w" Then
'        AsignarFonemaBase = 22   ' /w/
'        Exit Function
'    End If
'
'    ' ---------------------------------------------------------
'    ' NASALS
'    ' ---------------------------------------------------------
'    Select Case grafema
'        Case "m": AsignarFonemaBase = 36: Exit Function
'        Case "n": AsignarFonemaBase = 37: Exit Function
'        Case "ny": AsignarFonemaBase = 38: Exit Function
'        Case "ng": AsignarFonemaBase = 39: Exit Function
'    End Select
'
'    ' ---------------------------------------------------------
'    ' LATERALS
'    ' ---------------------------------------------------------
'    If grafema = "l" Then AsignarFonemaBase = 62: Exit Function
'    If grafema = "ll" Then AsignarFonemaBase = 63: Exit Function
'    If grafema = "l·l" Then AsignarFonemaBase = 63: Exit Function
'
'    ' ---------------------------------------------------------
'    ' VIBRANTS
'    ' ---------------------------------------------------------
'    If grafema = "rr" Then
'        AsignarFonemaBase = 60   ' múltiple
'        Exit Function
'    End If
'
'    If grafema = "r" Then
'        If ant = "" Or Not ant Like "[aeiouàèéíïòóúü]" Then
'            AsignarFonemaBase = 60   ' inicial / múltiple
'        Else
'            AsignarFonemaBase = 59   ' intervocàlica
'        End If
'        Exit Function
'    End If
'
'    ' ---------------------------------------------------------
'    ' AFRICADES
'    ' ---------------------------------------------------------
'    If grafema = "tx" Then AsignarFonemaBase = 57: Exit Function
'    If grafema = "tg" Or grafema = "tj" Then AsignarFonemaBase = 58: Exit Function
'    If grafema = "ts" Or grafema = "tz" Then AsignarFonemaBase = 46: Exit Function
'
'    ' ---------------------------------------------------------
'    ' FRICATIVES
'    ' ---------------------------------------------------------
'    If grafema = "s" Then AsignarFonemaBase = 42: Exit Function
'    If grafema = "z" Then AsignarFonemaBase = 43: Exit Function
'
'    ' X simple ? /?/ (cas general mallorquí)
'    ' Els casos /ks/ ja es tracten a ProcesarSilaba
'    If grafema = "x" Then
'        AsignarFonemaBase = 44
'        Exit Function
'    End If
'
'    If grafema = "ix" Then
'        AsignarFonemaBase = 44
'        Exit Function
'    End If
'
'    If grafema = "j" Then AsignarFonemaBase = 45: Exit Function
'    If grafema = "ge" Or grafema = "gi" Then AsignarFonemaBase = 45: Exit Function
'
'    ' ---------------------------------------------------------
'    ' OCLUSIVES
'    ' ---------------------------------------------------------
'    Select Case grafema
'        Case "v": AsignarFonemaBase = 41: Exit Function
'        Case "p": AsignarFonemaBase = 30: Exit Function
'        Case "b": AsignarFonemaBase = 31: Exit Function
'        Case "t": AsignarFonemaBase = 32: Exit Function
'        Case "d": AsignarFonemaBase = 33: Exit Function
'        Case "c", "k", "qu": AsignarFonemaBase = 34: Exit Function
'        Case "g", "gu": AsignarFonemaBase = 35: Exit Function
'    End Select
'
'    ' ---------------------------------------------------------
'    ' NO RECONEGUT ? 255
'    ' ---------------------------------------------------------
'    AsignarFonemaBase = 255
'
'End Function
'
'' ============================================================
''   NEUTRALITZACIONS IB  (VERSIÓ RECONSTRUÏDA)
'' ============================================================
'Private Function AplicarNeutralizaciones(cadena As String) As String
'
'    ' ---------------------------------------------------------
'    ' CE / CI ? SE / SI  (34 + 2/4 ? 42 + 2/4)
'    ' ---------------------------------------------------------
'    cadena = Replace(cadena, "34 2", "42 2")
'    cadena = Replace(cadena, "34 4", "42 4")
'
'    ' ---------------------------------------------------------
'    ' GE / GI ? JE / JI  (35 + 2/4 ? 45 + 2/4)
'    ' ---------------------------------------------------------
'    cadena = Replace(cadena, "35 2", "45 2")
'    cadena = Replace(cadena, "35 4", "45 4")
'
'    ' ---------------------------------------------------------
'    ' R final simple ? eliminar (mallorquín)
'    ' ---------------------------------------------------------
'    If Right$(Trim$(cadena), 2) = "59" Then
'        cadena = Left$(Trim$(cadena), Len(Trim$(cadena)) - 2)
'    End If
'
'    AplicarNeutralizaciones = cadena
'End Function
'
'' ============================================================
''   ASSIMILACIONS IB  (VERSIÓ RECONSTRUÏDA)
'' ============================================================
'Private Function AplicarAssimilaciones(cadena As String) As String
'
'    ' ---------------------------------------------------------
'    ' S intervocàlica ? Z
'    ' ---------------------------------------------------------
'    cadena = Replace(cadena, "42 1", "43 1")
'    cadena = Replace(cadena, "42 2", "43 2")
'    cadena = Replace(cadena, "42 3", "43 3")
'    cadena = Replace(cadena, "42 4", "43 4")
'    cadena = Replace(cadena, "42 5", "43 5")
'    cadena = Replace(cadena, "42 6", "43 6")
'    cadena = Replace(cadena, "42 7", "43 7")
'    cadena = Replace(cadena, "42 8", "43 8")
'
'    ' ---------------------------------------------------------
'    ' N + K/G ? ? (velar)
'    ' ---------------------------------------------------------
'    cadena = Replace(cadena, "37 34", "39")
'    cadena = Replace(cadena, "37 35", "39")
'
'    ' ---------------------------------------------------------
'    ' R inicial ? RR
'    ' ---------------------------------------------------------
'    If Left$(Trim$(cadena), 2) = "59" Then
'        cadena = "60" & Mid$(Trim$(cadena), 3)
'    End If
'
'    ' ---------------------------------------------------------
'    ' U + N ? U + ?  (mallorquín)
'    ' ---------------------------------------------------------
'    cadena = Replace(cadena, "7 37", "7 39")
'
'    AplicarAssimilaciones = cadena
'End Function
'
'' ============================================================
''   REDUCCIONS IB  (VERSIÓ RECONSTRUÏDA)
'' ============================================================
'Private Function AplicarReducciones(cadena As String) As String
'
'    Do While InStr(cadena, "  ") > 0
'        cadena = Replace(cadena, "  ", " ")
'    Loop
'
'    AplicarReducciones = Trim$(cadena)
'End Function
'
'' ============================================================
''   SCHWA IB  (VERSIÓ RECONSTRUÏDA)
'' ============================================================
'Private Function AplicarSchwa(cadena As String) As String
'
'    ' ARTICLE SALAT
'    cadena = Replace(cadena, "es ", "8 42 ")
'    cadena = Replace(cadena, "sa ", "42 8 ")
'    cadena = Replace(cadena, "ses ", "42 8 42 ")
'
'    ' APÒCOPES
'    cadena = Replace(cadena, "can' ", "34 8 37 ")
'    cadena = Replace(cadena, "ca' ", "34 8 ")
'
'    ' PROCLÍTICS
'    cadena = Replace(cadena, "de ", "33 8 ")
'    cadena = Replace(cadena, "me ", "36 8 ")
'    cadena = Replace(cadena, "te ", "32 8 ")
'    cadena = Replace(cadena, "se ", "42 8 ")
'
'    AplicarSchwa = cadena
'End Function
'
