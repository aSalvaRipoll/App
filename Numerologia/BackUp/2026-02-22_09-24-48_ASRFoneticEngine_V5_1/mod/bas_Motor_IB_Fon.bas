Attribute VB_Name = "bas_Motor_IB_Fon"

Option Compare Database
Option Explicit

' ============================================================
'   MOTOR FONÉTICO — BALEAR (IB)
'   Construcción de IdsFonemes + IPA final
'   Arquitectura paralela al motor CA
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
    Dim res As String

    ObjDTO.IdsFonemas = ""
    ObjDTO.FonemasFinal = ""

    ' 0. Normalització (IB no toca grafies alienes)
    frase = NormalizarVocales(ObjDTO.SilabasFinal)

    ' 1. Separar síl·labes per "|"
    arrSilabas = Split(frase, "|")

    For i = 0 To UBound(arrSilabas)

        silabaCruda = arrSilabas(i)

        ' ---------------------------------------------------------
        ' 2. Separador de palabra (sílaba vacía)
        ' ---------------------------------------------------------
        If Trim$(silabaCruda) = "" Then

            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "#"

            ' Ligadura automàtica (Mode 2)
            If CFG.ModoLigadura = 2 Then

                ultimaSilaba = BuscarSilabaRealAnterior(arrSilabas, i)
                siguienteSilaba = BuscarSilabaRealPosterior(arrSilabas, i)

                If HayLigaduraAutomatica(ultimaSilaba, siguienteSilaba) Then
                    ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "84,"
                End If

            End If

            GoTo SiguienteIteracion
        End If

        ' ---------------------------------------------------------
        ' 3. Modificadors prosòdics (acento)
        ' ---------------------------------------------------------
        If InStr(silabaCruda, "(") > 0 Then
            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "80,"   ' tònica
        ElseIf InStr(silabaCruda, "[") > 0 Then
            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "81,"   ' secundària
        Else
            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "82,"   ' àtona
        End If

        ' ---------------------------------------------------------
        ' 4. Processar grafemes (IB)
        ' ---------------------------------------------------------
        ProcesarSilaba silabaCruda

        ' ---------------------------------------------------------
        ' 5. Separador sil·làbic
        ' ---------------------------------------------------------
        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "|"

SiguienteIteracion:
    Next i

    ' ---------------------------------------------------------
    ' 6. Neteja final (igual que CA)
    ' ---------------------------------------------------------
    ObjDTO.IdsFonemas = Replace(ObjDTO.IdsFonemas, "|#", "#")
    ObjDTO.IdsFonemas = Replace(ObjDTO.IdsFonemas, "#|", "#")

    ObjDTO.IdsFonemas = Replace(ObjDTO.IdsFonemas, ",|", "|")
    
    While Right$(ObjDTO.IdsFonemas, 1) = "|" Or Right$(ObjDTO.IdsFonemas, 1) = "#" Or Right$(ObjDTO.IdsFonemas, 1) = ","
        ObjDTO.IdsFonemas = Left$(ObjDTO.IdsFonemas, Len(ObjDTO.IdsFonemas) - 1)
    Wend

    

    ' ---------------------------------------------------------
    ' 7. Construcció IPA final (mateix esquema que CA)
    ' ---------------------------------------------------------
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

' ============================================================
'   PROCESSAR SÍL·LABA — BALEAR (IB)
' ============================================================
Private Sub ProcesarSilaba(ByVal silabaCruda As String)

    Dim silaba As String
    Dim ligaduraID As Byte
    Dim i As Long
    Dim grafema As String
    Dim id As Byte
    Dim antCh As String
    Dim sigCh As String
    Dim esAtona As Boolean

    ' Detectar si la síl·laba és àtona
    esAtona = True
    If InStr(silabaCruda, "(") > 0 Then esAtona = False
    If InStr(silabaCruda, "[") > 0 Then esAtona = False

    ' Netejar marcadors
    silaba = silabaCruda
    silaba = Replace(silaba, "(", "")
    silaba = Replace(silaba, ")", "")
    silaba = Replace(silaba, "[", "")
    silaba = Replace(silaba, "]", "")
    silaba = Trim$(silaba)

    ' 1. Ligadura manual (Mode 1)
    If CFG.ModoLigadura = 1 Then
        ligaduraID = DetectarLigaduraManual(silaba)
        If ligaduraID <> 0 Then
            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & CStr(ligaduraID) & ","
        End If
    End If

    ' 2. Processar grafemes IB
    i = 1
    Do While i <= Len(silaba)

        grafema = DetectarDigrafo(silaba, i)

        ' Context anterior
        If i > 1 Then
            antCh = Mid$(silaba, i - 1, 1)
        Else
            antCh = ""
        End If

        ' Context següent
        If i + Len(grafema) <= Len(silaba) Then
            sigCh = Mid$(silaba, i + Len(grafema), 1)
        Else
            sigCh = ""
        End If

        ' --- Schwa IB (proclítics i apòcops) ---
        If esAtona Then
            If grafema = "a" Or grafema = "e" Then
                id = 8   ' schwa
                GoTo Agregar
            End If
        End If

        ' 3. Fonema base IB
        id = AsignarFonemaBase(grafema, sigCh, antCh)

Agregar:
        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & CStr(id) & ","

        i = i + Len(grafema)

    Loop

End Sub

' ============================================================
'   DETECTAR DÍGRAFS — BALEAR (IB)
' ============================================================
Private Function DetectarDigrafo(t As String, pos As Long) As String

    Dim L As Long: L = Len(t)
    Dim par As String, tri As String, sig As String

    ' 1. Trigrama "l·l" (igual que CA)
    If pos + 2 <= L Then
        tri = Mid$(t, pos, 3)
        If tri = "l·l" Then
            DetectarDigrafo = tri
            Exit Function
        End If
    End If

    ' 2. Dígrafs IB
    If pos + 1 <= L Then
        par = Mid$(t, pos, 2)
        sig = Mid$(t, pos + 2, 1)

        Select Case par

            Case "ny": DetectarDigrafo = par: Exit Function
            Case "ll": DetectarDigrafo = par: Exit Function
            Case "rr": DetectarDigrafo = par: Exit Function
            Case "ss": DetectarDigrafo = par: Exit Function

            Case "tx": DetectarDigrafo = par: Exit Function
            Case "tg", "tj": DetectarDigrafo = par: Exit Function

            Case "ts", "tz": DetectarDigrafo = par: Exit Function

            Case "ix": DetectarDigrafo = par: Exit Function

            ' Qu / qü
            Case "qu": DetectarDigrafo = par: Exit Function
            Case "qü": DetectarDigrafo = par: Exit Function

            ' Gu / güe / güi
            Case "gu"
                If sig = "e" Or sig = "i" Then
                    DetectarDigrafo = par
                    Exit Function
                End If

            Case "gü"
                If sig = "e" Or sig = "i" Then
                    DetectarDigrafo = par
                    Exit Function
                End If

        End Select
    End If

    ' 3. "ig" final ? /t?/ (igual que CA)
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
'   FONEMA BASE — BALEAR (IB)
' ============================================================
Private Function AsignarFonemaBase(grafema As String, _
                                      Optional sig As String = "", _
                                      Optional ant As String = "") As Byte

    grafema = LCase$(grafema)
    sig = LCase$(sig)
    ant = LCase$(ant)

    ' VOCALS IB
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
    If grafema = "j" Or grafema = "y" Then
        AsignarFonemaBase = 21: Exit Function   ' /j/
    End If

    If grafema = "w" Then
        AsignarFonemaBase = 22: Exit Function   ' /w/
    End If

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
        If ant = "" Or Not ant Like "[aeiouàèéíïòóúü]" Then
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
    If grafema = "z" Then AsignarFonemaBase = 43: Exit Function
    If grafema = "x" Then AsignarFonemaBase = 44: Exit Function
    If grafema = "ix" Then AsignarFonemaBase = 44: Exit Function
    If grafema = "j" Then AsignarFonemaBase = 45: Exit Function
    If grafema = "ge" Or grafema = "gi" Then AsignarFonemaBase = 45: Exit Function

    ' OCLUSIVES
    Select Case grafema
        Case "v": AsignarFonemaBase = 41: Exit Function
        Case "p": AsignarFonemaBase = 30: Exit Function
        Case "b": AsignarFonemaBase = 31: Exit Function
        Case "t": AsignarFonemaBase = 32: Exit Function
        Case "d": AsignarFonemaBase = 33: Exit Function
        Case "c", "k", "qu": AsignarFonemaBase = 34: Exit Function
        Case "g", "gu": AsignarFonemaBase = 35: Exit Function
    End Select

    ' No reconegut ? 255
    AsignarFonemaBase = 255

End Function

' ============================================================
'   NORMALITZAR VOCALS — BALEAR (neutre)
' ============================================================
Private Function NormalizarVocales(ByVal texto As String) As String
    NormalizarVocales = texto
End Function

' ============================================================
'   NEUTRALITZACIONS IB
' ============================================================
Private Function AplicarNeutralizaciones(cadena As String) As String
    AplicarNeutralizaciones = cadena
End Function


' ============================================================
'   ASSIMILACIONS IB
' ============================================================
Private Function AplicarAssimilaciones(cadena As String) As String

    ' S intervocàlica ? Z
    cadena = Replace(cadena, "42 1", "43 1")
    cadena = Replace(cadena, "42 2", "43 2")
    cadena = Replace(cadena, "42 3", "43 3")
    cadena = Replace(cadena, "42 4", "43 4")
    cadena = Replace(cadena, "42 5", "43 5")
    cadena = Replace(cadena, "42 6", "43 6")
    cadena = Replace(cadena, "42 7", "43 7")
    cadena = Replace(cadena, "42 8", "43 8")

    ' N + K/G ? ?
    cadena = Replace(cadena, "37 34", "39")
    cadena = Replace(cadena, "37 35", "39")

    ' R inicial ? RR
    If Left$(Trim$(cadena), 2) = "59" Then
        cadena = "60" & Mid$(Trim$(cadena), 3)
    End If

    AplicarAssimilaciones = cadena

End Function


' ============================================================
'   REDUCCIONS IB
' ============================================================
Private Function AplicarReducciones(cadena As String) As String

    Do While InStr(cadena, "  ") > 0
        cadena = Replace(cadena, "  ", " ")
    Loop

    AplicarReducciones = Trim$(cadena)

End Function


' ============================================================
'   SCHWA IB
' ============================================================
Private Function AplicarSchwa(cadena As String) As String

    ' ARTICLE SALAT
    cadena = Replace(cadena, "es ", "8 42 ")
    cadena = Replace(cadena, "sa ", "42 8 ")
    cadena = Replace(cadena, "ses ", "42 8 42 ")

    ' APÒCOPES
    cadena = Replace(cadena, "can' ", "34 8 37 ")
    cadena = Replace(cadena, "ca' ", "34 8 ")

    ' PROCLÍTICS
    cadena = Replace(cadena, "de ", "33 8 ")
    cadena = Replace(cadena, "me ", "36 8 ")
    cadena = Replace(cadena, "te ", "32 8 ")
    cadena = Replace(cadena, "se ", "42 8 ")

    AplicarSchwa = cadena

End Function

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

' ============================================================
'   BUSCAR SÍLABA REAL ANTERIOR
' ============================================================
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

' ============================================================
'   BUSCAR SÍLABA REAL POSTERIOR
' ============================================================
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


