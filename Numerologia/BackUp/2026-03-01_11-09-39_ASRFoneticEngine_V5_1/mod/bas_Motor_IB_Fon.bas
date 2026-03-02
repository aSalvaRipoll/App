Attribute VB_Name = "bas_Motor_IB_Fon"

Option Compare Database
Option Explicit

' ============================================================
'   MOTOR FONÉTICA — BALEAR (IB) — Versión limpia 7.1
'   Arquitectura paralela a los motores CA i ES
' ============================================================
Public Sub ConstruirCadenaFonemas_IB()

    Dim arrSilabas As Variant
    Dim i As Long
    Dim silabaCruda As String
    Dim ultimaSilaba As String
    Dim siguienteSilaba As String
    Dim frase As String
    Dim arrFon As Variant
    Dim Fon As Variant
    
    Dim ligaduraID As String
    Dim res As String

    ObjDTO.IdsFonemas = ""
    ObjDTO.FonemasFinal = ""

    frase = NormalizarVocales(ObjDTO.SilabasFinal)
    arrSilabas = Split(frase, "|")

For i = 0 To UBound(arrSilabas)

        silabaCruda = arrSilabas(i)

        ' ---------------------------------------------------------
        ' 2. Separador de paraula (síl·laba buida)
        ' ---------------------------------------------------------
        If Trim$(silabaCruda) = "" Then

            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "#"

            ' Ligadura automàtica (Mode 2)
            'If CFG.ModoLigadura = 2 Then

                ultimaSilaba = BuscarSilabaRealAnterior(arrSilabas, i)
                siguienteSilaba = BuscarSilabaRealPosterior(arrSilabas, i)

                If HayLigaduraAutomatica(ultimaSilaba, siguienteSilaba) Then
                    ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "84,"
                End If

            'End If

            GoTo SiguienteIteracion
        End If

        ' 1. Ligadura manual (Mode 1)
        'If CFG.ModoLigadura = 1 Then
            ligaduraID = DetectarLigaduraManual(silabaCruda)
            If ligaduraID <> 0 Then
                ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & CStr(ligaduraID) & ","
            End If
        'End If


        ' ---------------------------------------------------------
        ' 3. Modificadors prosòdics (acento)
        ' ---------------------------------------------------------
        If InStr(silabaCruda, "(") > 0 Then
            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "80,"
        ElseIf InStr(silabaCruda, "[") > 0 Then
            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "81,"
        Else
            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "82,"
        End If

        ' ---------------------------------------------------------
        ' 4. Processar grafemes (amb schwa)
        '    Passem la síl·laba crua (amb parèntesis/brackets)
        ' ---------------------------------------------------------
        ProcesarSilaba silabaCruda

        ' ---------------------------------------------------------
        ' 5. Separador sil·làbic
        ' ---------------------------------------------------------
        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "|"

SiguienteIteracion:
    Next i

'    For i = 0 To UBound(arrSilabas)
'
'        silabaCruda = arrSilabas(i)
'
'        ' Separador de palabra
'        If Trim$(silabaCruda) = "" Then
'
'            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "#"
'
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
'        ' Prosodia
'        If InStr(silabaCruda, "(") > 0 Then
'            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "80,"
'        ElseIf InStr(silabaCruda, "[") > 0 Then
'            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "81,"
'        Else
'            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "82,"
'        End If
'
'        ' Procesar grafemas
'        ProcesarSilaba silabaCruda
'
'        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "|"
'
'SiguienteIteracion:
'    Next i

    ' Limpieza estructural
    ObjDTO.IdsFonemas = Replace(ObjDTO.IdsFonemas, "|#", "#")
    ObjDTO.IdsFonemas = Replace(ObjDTO.IdsFonemas, "#|", "#")
    ObjDTO.IdsFonemas = Replace(ObjDTO.IdsFonemas, ",|", "|")
    ObjDTO.IdsFonemas = Replace(ObjDTO.IdsFonemas, "84,84,", "84,")

    While Right$(ObjDTO.IdsFonemas, 1) = "|" Or _
          Right$(ObjDTO.IdsFonemas, 1) = "#" Or _
          Right$(ObjDTO.IdsFonemas, 1) = ","

        ObjDTO.IdsFonemas = Left$(ObjDTO.IdsFonemas, Len(ObjDTO.IdsFonemas) - 1)
    Wend

    ' Construcción IPA
    arrSilabas = Split(ObjDTO.IdsFonemas, "|")
    res = ""

    For i = 0 To UBound(arrSilabas)

        If i = 0 Then res = res & "/"

        arrFon = Split(arrSilabas(i), ",")

        For Each Fon In arrFon

            If Fon = "#84" Then
                Fon = 84
            ElseIf Fon = "#82" Then
                res = res & "/ /"
                Fon = 82
            End If

            If Left$(Fon, 1) = "#" Then
                res = res & "/ /"
                Fon = Replace(Fon, "#", "")
            End If

            If Trim$(Fon) <> "" And Trim$(Fon) <> "82" Then
                res = res & Replace(ObtenerIPA(Fon), "/", "")
            End If

        Next Fon

        If i = UBound(arrSilabas) Then
            res = res & "/ "
        End If

    Next i

    ObjDTO.FonemasFinal = Trim$(res)

End Sub

Private Sub ProcesarSilaba(ByVal silabaCruda As String)

    Dim silaba As String
    Dim ligaduraID As Byte
    Dim i As Long
    Dim grafema As String
    Dim id As Byte
    Dim antCh As String
    Dim sigCh As String
    Dim esAtona As Boolean

    esAtona = True
    If InStr(silabaCruda, "(") > 0 Then esAtona = False
    If InStr(silabaCruda, "[") > 0 Then esAtona = False

    silaba = silabaCruda
    silaba = Replace(silaba, "(", "")
    silaba = Replace(silaba, ")", "")
    silaba = Replace(silaba, "[", "")
    silaba = Replace(silaba, "]", "")
    silaba = Replace(silaba, "'", "")
    silaba = Replace(silaba, "’", "")
    silaba = Trim$(silaba)

    If silaba = "" Then Exit Sub

    ' Ligadura manual
    ligaduraID = DetectarLigaduraManual(silaba)
    If ligaduraID <> 0 Then ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "84,"

    silaba = Replace(silaba, "_", "")

    i = 1
    Do While i <= Len(silaba)

        grafema = DetectarDigrafo(silaba, i)
        If grafema = "" Then i = i + 1: GoTo SegGraf

        If i > 1 Then antCh = Mid$(silaba, i - 1, 1) Else antCh = ""
        If i + Len(grafema) <= Len(silaba) Then sigCh = Mid$(silaba, i + Len(grafema), 1) Else sigCh = ""

        ' SCHWA proclítica
        If esAtona And EsProclitic(silabaCruda) Then
            If grafema = "a" Or grafema = "e" Then
                ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "8,"
                i = i + Len(grafema)
                GoTo SegGraf
            End If
        End If

        ' ============================================================
        '   REGLA EX ? /eks/ o /e?/ segú TONICIDAD y excepción XT
        ' ============================================================
        If grafema = "x" And antCh = "e" Then
        
            Dim arr As Variant
            Dim idx As Long
            Dim sigSil As String
            Dim primera As String
            Dim esTonica As Boolean
        
            ' Obtener array de sílabas finales
            arr = Split(ObjDTO.SilabasFinal, "|")
        
            ' Buscar posició de la síl·laba actual
            For idx = LBound(arr) To UBound(arr)
                If Trim$(arr(idx)) = Trim$(silabaCruda) Then Exit For
            Next idx
        
            ' Obtenir síl·laba següent real
            sigSil = BuscarSilabaRealPosterior(arr, idx)
        
            If sigSil <> "" Then
        
                ' Netejar parèntesis i espais
                esTonica = (InStr(sigSil, "(") > 0)
                sigSil = Replace(sigSil, "(", "")
                sigSil = Replace(sigSil, ")", "")
                sigSil = Replace(sigSil, "[", "")
                sigSil = Replace(sigSil, "]", "")
                sigSil = Trim$(sigSil)
        
                If Len(sigSil) >= 1 Then primera = Left$(sigSil, 1)
        
                ' EXCEPCIÓ: XT ? /kst/
                If primera = "t" Then
                    ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "34,42,"   ' /ks/
                    i = i + 1
                    GoTo SegGraf
                End If
        
                ' TONICA ? /eks/
                If esTonica Then
                    ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "34,42,"   ' /ks/
                    i = i + 1
                    GoTo SegGraf
                End If
        
                ' ATONA ? /e?/
                ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "44,"          ' /?/
                i = i + 1
                GoTo SegGraf
        
            End If
        End If

        ' ============================================================
        '   REGLAS MALLORQUINES (QU / QÜ / GÜ / GU)
        ' ============================================================

        ' QU + A/O/U ? /kw/
        If grafema = "qu" Then
            If sigCh Like "[aouàòóú]" Then
                ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "34,22,"
                i = i + 2
                GoTo SegGraf
            End If
        End If

        ' QÜ + E/I ? /kwe/
        If grafema = "qü" Then
            If sigCh Like "[eéiíè]" Then
                ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "34,22," & AsignarFonemaBase(sigCh) & ","
                i = i + 3
                GoTo SegGraf
            End If
        End If

        ' GÜ + E/I ? /gwe/
        If grafema = "gü" Then
            If sigCh Like "[eéiíè]" Then
                ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "35,22," & AsignarFonemaBase(sigCh) & ","
                i = i + 3
                GoTo SegGraf
            End If
        End If

        ' GU + A/O/U ? /gw/
        If grafema = "gu" Then
            If sigCh Like "[aouàòóú]" Then
                ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "35,22,"
                i = i + 2
                GoTo SegGraf
            End If
        End If

        ' ============================================================
        '   RESTO DE CASOS ESPECIALES
        ' ============================================================

        ' IX ? /?/
        If grafema = "ix" Then
            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "44,"
            i = i + 2
            GoTo SegGraf
        End If

        ' TX ? /t?/
        If grafema = "tx" Then
            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "57,"
            i = i + 2
            GoTo SegGraf
        End If

        ' TG/TJ ? /d?/
        If grafema = "tg" Or grafema = "tj" Then
            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "58,"
            i = i + 2
            GoTo SegGraf
        End If

        ' TS/TZ ? /ts/
        If grafema = "ts" Or grafema = "tz" Then
            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "46,"
            i = i + 2
            GoTo SegGraf
        End If

        ' Fonema base
        id = AsignarFonemaBase(grafema, sigCh, antCh)
        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & id & ","
        i = i + Len(grafema)

SegGraf:
    Loop

End Sub

Private Function EsProclitic(sil As String) As Boolean
    Dim s As String
    s = LCase$(Trim$(sil))
    EsProclitic = (s = "de" Or s = "me" Or s = "te" Or s = "se" Or s = "es" Or s = "sa" Or s = "ses")
End Function

' ============================================================
'   DETECTAR DÍGRAFS — BALEAR (IB)
' ============================================================
Private Function DetectarDigrafo(t As String, pos As Long) As String

    Dim L As Long: L = Len(t)
    Dim par As String, tri As String, sig As String

    ' 1. Trigrama "l·l"
    If pos + 2 <= L Then
        tri = Mid$(t, pos, 3)
        If tri = "l·l" Then
            DetectarDigrafo = tri
            Exit Function
        End If
    End If

    ' 2. Dígrafs de 2 lletres
    If pos + 1 <= L Then
        par = Mid$(t, pos, 2)

        Select Case par
            Case "ny", "ll", "rr", "ss"
                DetectarDigrafo = par: Exit Function

            Case "tx", "tg", "tj", "ts", "tz", "ix"
                DetectarDigrafo = par: Exit Function
        End Select

        ' QU + vocal (todas) ? dígrafo
        If par = "qu" Then
            If pos + 2 <= L Then
                sig = Mid$(t, pos + 2, 1)
                If sig Like "[aeiouàèéíïòóúü]" Then
                    DetectarDigrafo = par
                    Exit Function
                End If
            End If
        End If
        
        ' QÜ + e/i
        If par = "qü" Then
            If pos + 2 <= L Then
                sig = Mid$(t, pos + 2, 1)
                If sig Like "[eéiíè]" Then
                    DetectarDigrafo = par
                    Exit Function
                End If
            End If
        End If
        
        ' GÜ + e/i
        If par = "gü" Then
            If pos + 2 <= L Then
                sig = Mid$(t, pos + 2, 1)
                If sig Like "[eéiíè]" Then
                    DetectarDigrafo = par
                    Exit Function
                End If
            End If
        End If

        ' GU + a/o/u
        If par = "gu" Then
            If pos + 2 <= L Then
                sig = Mid$(t, pos + 2, 1)
                If sig Like "[aouàòóú]" Then
                    DetectarDigrafo = par
                    Exit Function
                End If
            End If
        End If

    End If

    ' 3. "ig" final ? "tx"
    If pos + 1 = L Then
        If Mid$(t, pos, 2) = "ig" Then
            DetectarDigrafo = "tx"
            Exit Function
        End If
    End If

    ' 4. Si no és dígraf ? 1 lletra
    DetectarDigrafo = Mid$(t, pos, 1)

End Function

' ============================================================
'   FONEMA BASE — BALEAR (IB) net
' ============================================================
Private Function AsignarFonemaBase(grafema As String, _
                                   Optional sig As String = "", _
                                   Optional ant As String = "") As Byte

    grafema = LCase$(grafema)
    sig = LCase$(sig)
    ant = LCase$(ant)

    ' VOCALS
    Select Case grafema
        Case "a", "à", "á": AsignarFonemaBase = 1: Exit Function
        Case "e", "é": AsignarFonemaBase = 2: Exit Function
        Case "è": AsignarFonemaBase = 3: Exit Function
        Case "i", "í", "ï": AsignarFonemaBase = 4: Exit Function
        Case "o", "ó": AsignarFonemaBase = 5: Exit Function
        Case "ò": AsignarFonemaBase = 6: Exit Function
        Case "u", "ú", "ü": AsignarFonemaBase = 7: Exit Function
    End Select

    ' SEMIVOCALS
    If grafema = "j" Or grafema = "y" Then AsignarFonemaBase = 21: Exit Function
    If grafema = "w" Then AsignarFonemaBase = 22: Exit Function

    ' NASALS
    Select Case grafema
        Case "m": AsignarFonemaBase = 36: Exit Function
        Case "n": AsignarFonemaBase = 37: Exit Function
        Case "ny": AsignarFonemaBase = 38: Exit Function
        Case "ng": AsignarFonemaBase = 39: Exit Function
    End Select

    ' LATERALS
    If grafema = "l" Then AsignarFonemaBase = 62: Exit Function
    If grafema = "ll" Then AsignarFonemaBase = 63: Exit Function
    If grafema = "l·l" Then AsignarFonemaBase = 63: Exit Function

    ' VIBRANTS
    If grafema = "rr" Then AsignarFonemaBase = 60: Exit Function

    If grafema = "r" Then
        If ant = "" Or ant = " " Or Not ant Like "[aeiouàèéíïòóúü]" Then
            AsignarFonemaBase = 60   ' inicial / múltiple
        Else
            AsignarFonemaBase = 59   ' intervocàlica
        End If
        Exit Function
    End If

    ' AFRICADES
    If grafema = "tx" Then AsignarFonemaBase = 57: Exit Function
    If grafema = "tg" Or grafema = "tj" Then AsignarFonemaBase = 58: Exit Function
    If grafema = "ts" Or grafema = "tz" Then AsignarFonemaBase = 46: Exit Function

    ' FRICATIVES
    If grafema = "s" Then AsignarFonemaBase = 42: Exit Function
    If grafema = "ss" Then AsignarFonemaBase = 42: Exit Function
    If grafema = "z" Then AsignarFonemaBase = 43: Exit Function
    If grafema = "x" Then AsignarFonemaBase = 44: Exit Function
    If grafema = "ix" Then AsignarFonemaBase = 44: Exit Function
    If grafema = "j" Then AsignarFonemaBase = 45: Exit Function
    If grafema = "ge" Or grafema = "gi" Then AsignarFonemaBase = 45: Exit Function

    ' OCLUSIVES
    Select Case grafema
        Case "p": AsignarFonemaBase = 30: Exit Function
        Case "b": AsignarFonemaBase = 31: Exit Function
        Case "t": AsignarFonemaBase = 32: Exit Function
        Case "d": AsignarFonemaBase = 33: Exit Function

        Case "c"
            If sig Like "[eéiíè]" Then
                AsignarFonemaBase = 42   ' /s/
            Else
                AsignarFonemaBase = 34   ' /k/
            End If
            Exit Function

        Case "k", "qu": AsignarFonemaBase = 34: Exit Function

        Case "g"
            If sig Like "[eéiíè]" Then
                AsignarFonemaBase = 45   ' /?/
            Else
                AsignarFonemaBase = 35   ' /g/
            End If
            Exit Function

        Case "gu": AsignarFonemaBase = 35: Exit Function

        Case "v": AsignarFonemaBase = 41: Exit Function
        Case "f": AsignarFonemaBase = 40: Exit Function
    End Select

    AsignarFonemaBase = 255

End Function

' ============================================================
'   REGLAS POSTERIORES — NO-OP (no tocan IDs)
'   (Se dejan por compatibilidad, pero no modifican fonemas)
' ============================================================
Private Function AplicarNeutralizaciones(cadena As String) As String
    AplicarNeutralizaciones = Trim$(cadena)
End Function

Private Function AplicarAssimilaciones(cadena As String) As String
    AplicarAssimilaciones = Trim$(cadena)
End Function

Private Function AplicarSchwa(cadena As String) As String
    AplicarSchwa = Trim$(cadena)
End Function

Private Function AplicarReducciones(cadena As String) As String
    Do While InStr(cadena, "  ") > 0
        cadena = Replace(cadena, "  ", " ")
    Loop
    AplicarReducciones = Trim$(cadena)
End Function

' ============================================================
'   LIGADURA, NORMALITZACIÓ, CERCA SÍL·LABES, IPA
'   (igual que ja tenies, sense tocar-los)
' ============================================================
Private Function DetectarLigaduraManual(ByRef silaba As String) As Byte
    DetectarLigaduraManual = 0
    If InStr(silaba, "_") > 0 Then
        silaba = Replace(silaba, "_", "")
        DetectarLigaduraManual = 84
    End If
End Function

Private Function HayLigaduraAutomatica(ultimaSilaba As String, primeraSilaba As String) As Boolean

    Dim ult As String
    Dim pri As String

    HayLigaduraAutomatica = False

    If ultimaSilaba = "" Or primeraSilaba = "" Then Exit Function

    ult = Right$(Trim$(ultimaSilaba), 1)
    pri = Left$(Trim$(primeraSilaba), 1)

    If ult Like "[aeiouáéíóúü]" And pri Like "[aeiouáéíóúü]" Then
        HayLigaduraAutomatica = True
        Exit Function
    End If

    If ult Like "[aeiouáéíóúü]" And primeraSilaba Like "h[aeiouáéíóúü]*" Then
        HayLigaduraAutomatica = True
        Exit Function
    End If

End Function

Private Function NormalizarVocales(ByVal texto As String) As String
    NormalizarVocales = texto
End Function

Private Function BuscarSilabaRealAnterior(arr As Variant, pos As Long) As String
    Dim j As Long
    For j = pos - 1 To 0 Step -1
        If Trim$(arr(j)) <> "" Then
            BuscarSilabaRealAnterior = Trim$(arr(j))
            Exit Function
        End If
    Next j
    BuscarSilabaRealAnterior = ""
End Function

Private Function BuscarSilabaRealPosterior(arr As Variant, pos As Long) As String
    Dim j As Long
    For j = pos + 1 To UBound(arr)
        If Trim$(arr(j)) <> "" Then
            BuscarSilabaRealPosterior = Trim$(arr(j))
            Exit Function
        End If
    Next j
    BuscarSilabaRealPosterior = ""
End Function

Private Function ObtenerIPA(ByVal idFonema As Long) As String
    Dim rs As DAO.Recordset

    If IPA_Cache Is Nothing Then
        Set IPA_Cache = CreateObject("Scripting.Dictionary")
    End If

    If IPA_Cache.Exists(idFonema) Then
        ObtenerIPA = IPA_Cache(idFonema)
        Exit Function
    End If

    If idFonema = 255 Then
        IPA_Cache.Add idFonema, ""
        ObtenerIPA = ""
        Exit Function
    End If

    Set rs = CurrentDb.OpenRecordset( _
        "SELECT IPA FROM qryFonemasValor WHERE ID=" & idFonema & ";", _
        dbOpenSnapshot)

    If Not (rs.EOF And rs.BOF) Then
        IPA_Cache.Add idFonema, Nz(rs!ipa, "")
        ObtenerIPA = Nz(rs!ipa, "")
    Else
        IPA_Cache.Add idFonema, ""
        ObtenerIPA = ""
    End If

    rs.Close
    Set rs = Nothing
End Function
