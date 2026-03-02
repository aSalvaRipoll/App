Attribute VB_Name = "bas_Motor_EU_Fon"
Option Compare Database
Option Explicit

' ============================================================
'   MOTOR FONÈTIC — EUSKERA
'   Construcción de IdsFonemas + IPA final
'   MISMA ARQUITECTURA QUE EL MOTOR ES/CA
' ============================================================

Public Sub ConstruirCadenaFonemas_EU()

    Dim arrSilabas As Variant
    Dim i As Long
    Dim silabaCruda As String
    Dim ultimaSilaba As String
    Dim SiguienteSilaba As String
    
    Dim frase As String
    Dim arrFon As Variant
    Dim Fon As Variant
    
    Dim ligaduraID As String
    Dim res As String

    ObjDTO.IdsFonemas = ""
    ObjDTO.FonemasFinal = ""

    ' 0. Normalizar (neutro para EU)
    frase = NormalizarVocales(ObjDTO.SilabasFinal)

    ' 1. Separar sílabas por "|"
    arrSilabas = Split(frase, "|")

    For i = 0 To UBound(arrSilabas)

        silabaCruda = arrSilabas(i)

        ' ---------------------------------------------------------
        ' 2. Separador de palabra (sílaba vacía)
        ' ---------------------------------------------------------
        If Trim$(silabaCruda) = "" Then

            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "#"

            ' Ligadura automática (Modo 2)
            'If CFG.ModoLigadura = 2 Then

                ultimaSilaba = BuscarSilabaRealAnterior(arrSilabas, i)
                SiguienteSilaba = BuscarSilabaRealPosterior(arrSilabas, i)

                If HayLigaduraAutomatica(ultimaSilaba, SiguienteSilaba) Then
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
        ' 3. Modificadores prosódicos (acento)
        ' ---------------------------------------------------------
        If InStr(silabaCruda, "(") > 0 Then
            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "80,"
        ElseIf InStr(silabaCruda, "[") > 0 Then
            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "81,"
        Else
            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "82,"
        End If

        ' ---------------------------------------------------------
        ' 4. Procesar grafemas (EU)
        ' ---------------------------------------------------------
        ProcesarSilaba silabaCruda

        ' ---------------------------------------------------------
        ' 5. Separador silábico
        ' ---------------------------------------------------------
        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "|"

SiguienteIteracion:
    Next i

    ' ---------------------------------------------------------
    ' 6. Limpieza final (idéntica al motor ES/CA)
    ' ---------------------------------------------------------
    ObjDTO.IdsFonemas = Replace(ObjDTO.IdsFonemas, "|#", "#")
    ObjDTO.IdsFonemas = Replace(ObjDTO.IdsFonemas, "#|", "#")
    ObjDTO.IdsFonemas = Replace(ObjDTO.IdsFonemas, ",|", "|")
    ObjDTO.IdsFonemas = Replace(ObjDTO.IdsFonemas, "84,84,", "84,")
    
    While Right$(ObjDTO.IdsFonemas, 1) = "|" Or Right$(ObjDTO.IdsFonemas, 1) = "#" Or Right$(ObjDTO.IdsFonemas, 1) = ","
        ObjDTO.IdsFonemas = Left$(ObjDTO.IdsFonemas, Len(ObjDTO.IdsFonemas) - 1)
    Wend

    ' ---------------------------------------------------------
    ' 7. Construcción IPA final (mismo esquema que ES/CA)
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
'   PROCESAR SÍLABA — EUSKERA
' ============================================================
Private Sub ProcesarSilaba(ByVal silabaCruda As String)

    Dim silaba As String
    Dim ligaduraID As Byte
    Dim i As Long
    Dim grafema As String
    Dim id As Byte
    Dim antCh As String
    Dim sigCh As String

    ' En EU no hay schwa: no necesitamos esAtona
    silaba = silabaCruda
    silaba = Replace(silaba, "(", "")
    silaba = Replace(silaba, ")", "")
    silaba = Replace(silaba, "[", "")
    silaba = Replace(silaba, "]", "")
    silaba = Trim$(silaba)

    ' 1. Ligadura manual (Modo 1)
    'If CFG.ModoLigadura = 1 Then
        ligaduraID = DetectarLigaduraManual(silaba)
        If ligaduraID <> 0 Then
            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & CStr(ligaduraID) & ","
        End If
    'End If

    ' 2. Procesar grafemas
    i = 1
    Do While i <= Len(silaba)

        grafema = DetectarDigrafo(silaba, i)

        ' Contexto anterior
        If i > 1 Then
            antCh = Mid$(silaba, i - 1, 1)
        Else
            antCh = ""
        End If

        ' Contexto siguiente
        If i + Len(grafema) <= Len(silaba) Then
            sigCh = Mid$(silaba, i + Len(grafema), 1)
        Else
            sigCh = ""
        End If

        ' 3. Fonema base euskera
        id = AsignarFonemaBase(grafema, sigCh, antCh)

        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & CStr(id) & ","

        i = i + Len(grafema)

    Loop

End Sub

' ============================================================
'   DETECTAR DÍGRAFS — EUSKERA
' ============================================================
Private Function DetectarDigrafo(t As String, pos As Long) As String

    Dim L As Long: L = Len(t)
    Dim par As String

    ' 1. Dígrafos de 2 letras
    If pos + 1 <= L Then
        par = Mid$(t, pos, 2)

        Select Case par

            ' Africada postalveolar sorda
            Case "tx": DetectarDigrafo = par: Exit Function

            ' Africada alveolar /ts/
            Case "tz", "ts": DetectarDigrafo = par: Exit Function

            ' Vibrante múltiple (ya la trataremos como "rr")
            Case "rr": DetectarDigrafo = par: Exit Function

            ' Geminadas ocasionales (préstamos)
            Case "dd", "tt": DetectarDigrafo = par: Exit Function

            ' Palatal lateral /?/ (préstamos)
            Case "ll": DetectarDigrafo = par: Exit Function

            ' Palatal nasal /?/ (préstamos)
            Case "ñ": DetectarDigrafo = par: Exit Function

        End Select
    End If

    ' 2. Si no es dígrafo ? 1 letra
    DetectarDigrafo = Mid$(t, pos, 1)

End Function

' ============================================================
'   FONEMA BASE — EUSKERA
' ============================================================
Private Function AsignarFonemaBase(grafema As String, _
                                      Optional sig As String = "", _
                                      Optional ant As String = "") As Byte

    grafema = LCase$(grafema)
    sig = LCase$(sig)
    ant = LCase$(ant)

    ' VOCALes
    Select Case grafema
        Case "a": AsignarFonemaBase = 1: Exit Function   ' /a/
        Case "e": AsignarFonemaBase = 2: Exit Function   ' /e/
        Case "i": AsignarFonemaBase = 4: Exit Function   ' /i/
        Case "o": AsignarFonemaBase = 5: Exit Function   ' /o/
        Case "u": AsignarFonemaBase = 7: Exit Function   ' /u/
    End Select

    ' SEMIVOCALES
    If grafema = "i" And sig Like "[aeou]" Then
        AsignarFonemaBase = 21: Exit Function   ' /j/
    End If

    If grafema = "u" And sig Like "[aeio]" Then
        AsignarFonemaBase = 22: Exit Function   ' /w/
    End If

    ' NASALES
    Select Case grafema
        Case "m": AsignarFonemaBase = 36: Exit Function
        Case "n": AsignarFonemaBase = 37: Exit Function
        Case "ñ": AsignarFonemaBase = 38: Exit Function
    End Select

    ' LATERALES
    If grafema = "l" Then AsignarFonemaBase = 62: Exit Function
    If grafema = "ll" Then AsignarFonemaBase = 63: Exit Function

    ' VIBRANTES
    If grafema = "rr" Then AsignarFonemaBase = 60: Exit Function

    If grafema = "r" Then
        If ant = "" Or Not ant Like "[aeiou]" Then
            AsignarFonemaBase = 60   ' inicial / postconsonántica ? múltiple
        Else
            AsignarFonemaBase = 59   ' intervocálica ? simple
        End If
        Exit Function
    End If

    ' AFRICADAS
    If grafema = "tx" Then AsignarFonemaBase = 57: Exit Function   ' /t?/

    If grafema = "ts" Then
        AsignarFonemaBase = 72
        Exit Function
    End If

    If grafema = "tz" Then
        AsignarFonemaBase = 73
        Exit Function
    End If

    ' FRICATIVAS
    If grafema = "s" Then AsignarFonemaBase = 42: Exit Function   ' /s/
    If grafema = "z" Then AsignarFonemaBase = 43: Exit Function   ' /z/
    If grafema = "x" Then AsignarFonemaBase = 44: Exit Function   ' /?/
    If grafema = "j" Then AsignarFonemaBase = 48: Exit Function   ' /x/
    If grafema = "h" Then AsignarFonemaBase = 49: Exit Function   ' /h/ (o muda si luego la filtras)
    If grafema = "f" Then AsignarFonemaBase = 40: Exit Function


    ' g + e/i ? fricativa velar sorda (como en ES)
    If grafema = "g" And (sig = "e" Or sig = "i") Then
        AsignarFonemaBase = 48      ' /x/
        Exit Function
    End If

    ' OCLUSIVAS
    Select Case grafema
        Case "p": AsignarFonemaBase = 30: Exit Function
        Case "b": AsignarFonemaBase = 31: Exit Function
        Case "t", "tt": AsignarFonemaBase = 32: Exit Function
        Case "d", "dd": AsignarFonemaBase = 33: Exit Function
        Case "k": AsignarFonemaBase = 34: Exit Function
        Case "g": AsignarFonemaBase = 35: Exit Function
    End Select

    ' Detectar ligadura manual
    If grafema = "_" Then AsignarFonemaBase = 84: Exit Function
    
    ' No reconocido ? 255
    AsignarFonemaBase = 255

End Function

' ============================================================
'   NORMALIZAR VOCALes — EUSKERA (neutro)
' ============================================================
Private Function NormalizarVocales(ByVal texto As String) As String
    ' No tocamos grafías aquí; cualquier cosa rara ? ID 255 en el mapeo.
    NormalizarVocales = texto
End Function

Private Function BuscarSilabaRealAnterior(arr As Variant, pos As Long) As String
    Dim j As Long
    For j = pos - 1 To LBound(arr) Step -1
        ' Si encuentro frontera de palabra, paro
        If Trim$(arr(j)) = "" Then Exit For
        
        ' Si encuentro sílaba real, la devuelvo
        BuscarSilabaRealAnterior = Trim$(arr(j))
        Exit Function
    Next j
    BuscarSilabaRealAnterior = ""
End Function

Private Function BuscarSilabaRealPosterior(arr As Variant, pos As Long) As String
    Dim j As Long
    For j = pos + 1 To UBound(arr)
        ' Si encuentro frontera de palabra, paro
        If Trim$(arr(j)) = "" Then Exit For
        
        ' Si encuentro sílaba real, la devuelvo
        BuscarSilabaRealPosterior = Trim$(arr(j))
        Exit Function
    Next j
    BuscarSilabaRealPosterior = ""
End Function

'Private Function BuscarSilabaRealAnterior(arr As Variant, pos As Long) As String
'    Dim j As Long
'    For j = pos - 1 To 0 Step -1
'        If Trim$(arr(j)) <> "" Then
'            BuscarSilabaRealAnterior = Trim$(arr(j))
'            Exit Function
'        End If
'    Next j
'    BuscarSilabaRealAnterior = ""
'End Function
'
'Private Function BuscarSilabaRealPosterior(arr As Variant, pos As Long) As String
'    Dim j As Long
'    For j = pos + 1 To UBound(arr)
'        If Trim$(arr(j)) <> "" Then
'            BuscarSilabaRealPosterior = Trim$(arr(j))
'            Exit Function
'        End If
'    Next j
'    BuscarSilabaRealPosterior = ""
'End Function

Private Function HayLigaduraAutomatica(ultimaSilaba As String, primeraSilaba As String) As Boolean

    Dim ult As String
    Dim pri As String

    HayLigaduraAutomatica = False

    If ultimaSilaba = "" Or primeraSilaba = "" Then Exit Function

    ult = Right$(LimpiaSilaba(ultimaSilaba), 1)
    pri = Left$(LimpiaSilaba(primeraSilaba), 1)

    If ult Like "[aeiouáéíóúü]" And pri Like "[aeiouáéíóúü]" Then
        HayLigaduraAutomatica = True
        Exit Function
    End If

    If ult Like "[aeiouáéíóúü]" And primeraSilaba Like "h[aeiouáéíóúü]*" Then
        HayLigaduraAutomatica = True
        Exit Function
    End If

End Function

Private Function LimpiaSilaba(ByVal strCad As String) As String
    
    strCad = Replace(strCad, "(", "")
    strCad = Replace(strCad, ")", "")
    strCad = Replace(strCad, "[", "")
    strCad = Replace(strCad, "]", "")
    LimpiaSilaba = Trim$(strCad)
    
End Function
'Private Function HayLigaduraAutomatica(ultimaSilaba As String, primeraSilaba As String) As Boolean
'
'    Dim ult As String
'    Dim pri As String
'
'    HayLigaduraAutomatica = False
'
'    If ultimaSilaba = "" Or primeraSilaba = "" Then Exit Function
'
'    ult = Right$(Trim$(ultimaSilaba), 1)
'    pri = Left$(Trim$(primeraSilaba), 1)
'
'    ' Vocal + vocal
'    If ult Like "[aeiou]" And pri Like "[aeiou]" Then
'        HayLigaduraAutomatica = True
'        Exit Function
'    End If
'
'    ' Vocal + h + vocal (h muda en medio)
'    If ult Like "[aeiou]" And primeraSilaba Like "h[aeiou]*" Then
'        HayLigaduraAutomatica = True
'        Exit Function
'    End If
'
'End Function

' ============================================================
'   DETECTAR LIGADURA MANUAL — EUSKERA
' ============================================================
Private Function DetectarLigaduraManual(ByRef silaba As String) As Byte

    DetectarLigaduraManual = 0
    
    ' Si la sílaba contiene "_", se elimina y se marca ligadura
    If InStr(silaba, "_") > 0 Then
        silaba = Replace(silaba, "_", "")
        DetectarLigaduraManual = 84   ' ID de ligadura manual
    End If

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


