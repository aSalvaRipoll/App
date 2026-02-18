Attribute VB_Name = "bas_Motor_CA_Fon"

Option Compare Database
Option Explicit

' ============================================================
'   MOTOR FONÉTICO — CATALÁN
'   Construcción de IdsFonemas + IPA final
'   MISMA ARQUITECTURA QUE EL MOTOR ES
' ============================================================

Public Sub ConstruirCadenaFonemas_CA()

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

    ' 0. Normalitzar (en aquest punt, neutra: no toquem grafies no catalanes)
    frase = NormalizarVocales_CA(ObjDTO.SilabasFinal)

    ' 1. Separar síl·labes per "|"
    arrSilabas = Split(frase, "|")

    For i = 0 To UBound(arrSilabas)

        silabaCruda = arrSilabas(i)

        ' ---------------------------------------------------------
        ' 2. Separador de paraula (síl·laba buida)
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
        ProcesarSilaba_CA silabaCruda

        ' ---------------------------------------------------------
        ' 5. Separador sil·làbic
        ' ---------------------------------------------------------
        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "|"

SiguienteIteracion:
    Next i

    ' ---------------------------------------------------------
    ' 6. Neteja final (idèntica al motor ES)
    ' ---------------------------------------------------------
    ObjDTO.IdsFonemas = Replace(ObjDTO.IdsFonemas, "|#", "#")
    ObjDTO.IdsFonemas = Replace(ObjDTO.IdsFonemas, "#|", "#")

    If Right$(ObjDTO.IdsFonemas, 1) = "|" Or _
       Right$(ObjDTO.IdsFonemas, 1) = "#" Or _
       Right$(ObjDTO.IdsFonemas, 1) = "," Then

        ObjDTO.IdsFonemas = Left$(ObjDTO.IdsFonemas, Len(ObjDTO.IdsFonemas) - 1)
    End If

    ObjDTO.IdsFonemas = Replace(ObjDTO.IdsFonemas, ",|", "|")

    ' ---------------------------------------------------------
    ' 7. Construcció IPA final (mateix esquema que ES)
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
'   PROCESSAR SÍL·LABA — CATALÀ (amb schwa)
' ============================================================
Private Sub ProcesarSilaba_CA(ByVal silabaCruda As String)

    Dim silaba As String
    Dim ligaduraID As Byte
    Dim i As Long
    Dim grafema As String
    Dim id As Byte
    Dim antCh As String
    Dim sigCh As String
    Dim esAtona As Boolean

    ' Detectar si la síl·laba és àtona (sense parèntesi ni claudàtor)
    esAtona = True
    If InStr(silabaCruda, "(") > 0 Then esAtona = False
    If InStr(silabaCruda, "[") > 0 Then esAtona = False

    ' Netejar marcadors per treballar amb la grafia neta
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

    ' 2. Processar grafemes
    i = 1
    Do While i <= Len(silaba)

        grafema = DetectarDigrafo_CA(silaba, i)

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

        ' --- Tractament del schwa /?/ ---
        If esAtona Then
            If grafema = "a" Or grafema = "e" Then
                id = 8   ' schwa
                GoTo Afegir
            End If
        End If

        ' 3. Fonema base català
        id = AsignarFonemaBase_CA(grafema, sigCh, antCh)

Afegir:
        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & CStr(id) & ","

        i = i + Len(grafema)

    Loop

End Sub

' ============================================================
'   DETECTAR DÍGRAFS — CATALÀ
' ============================================================
Private Function DetectarDigrafo_CA(t As String, pos As Long) As String

    Dim L As Long: L = Len(t)
    Dim par As String, tri As String, sig As String

    ' 1. Trigrama "l·l"
    If pos + 2 <= L Then
        tri = Mid$(t, pos, 3)
        If tri = "l·l" Then
            DetectarDigrafo_CA = tri
            Exit Function
        End If
    End If

    ' 2. Dígrafs de 2 lletres
    If pos + 1 <= L Then
        par = Mid$(t, pos, 2)
        sig = Mid$(t, pos + 2, 1)

        Select Case par

            ' Laterals
            Case "ll": DetectarDigrafo_CA = par: Exit Function

            ' Nasal palatal
            Case "ny": DetectarDigrafo_CA = par: Exit Function

            ' Africada sorda
            Case "tx": DetectarDigrafo_CA = par: Exit Function

            ' Africada sonora
            Case "tg", "tj": DetectarDigrafo_CA = par: Exit Function

            ' Fricativa palatal (ix)
            Case "ix": DetectarDigrafo_CA = par: Exit Function

            ' Qu / qü
            Case "qu": DetectarDigrafo_CA = par: Exit Function
            Case "qü": DetectarDigrafo_CA = par: Exit Function

            ' Gu / güe / güi
            Case "gu"
                If sig = "e" Or sig = "i" Then
                    DetectarDigrafo_CA = par
                    Exit Function
                End If

            Case "gü"
                If sig = "e" Or sig = "i" Then
                    DetectarDigrafo_CA = par
                    Exit Function
                End If

        End Select
    End If

    ' 3. "ig" final ? /t??/ (tractat com "tx")
    If pos + 1 = L Then
        If Mid$(t, pos, 2) = "ig" Then
            DetectarDigrafo_CA = "tx"
            Exit Function
        End If
    End If

    ' 4. Si no és dígraf ? 1 lletra
    DetectarDigrafo_CA = Mid$(t, pos, 1)

End Function

' ============================================================
'   FONEMA BASE — CATALÀ (sense schwa aquí)
' ============================================================
Private Function AsignarFonemaBase_CA(grafema As String, _
                                      Optional sig As String = "", _
                                      Optional ant As String = "") As Byte

    grafema = LCase$(grafema)
    sig = LCase$(sig)
    ant = LCase$(ant)

    ' VOCALS (tòniques / secundàries)
    Select Case grafema
        Case "a", "à": AsignarFonemaBase_CA = 1: Exit Function
        Case "e", "é": AsignarFonemaBase_CA = 2: Exit Function
        Case "è": AsignarFonemaBase_CA = 3: Exit Function
        Case "i", "í", "ï": AsignarFonemaBase_CA = 4: Exit Function
        Case "o", "ó": AsignarFonemaBase_CA = 5: Exit Function
        Case "ò": AsignarFonemaBase_CA = 6: Exit Function
        Case "u", "ú", "ü": AsignarFonemaBase_CA = 7: Exit Function
    End Select

    ' SEMIVOCALS
    If grafema = "j" Or grafema = "y" Then
        AsignarFonemaBase_CA = 21: Exit Function   ' /j/
    End If

    If grafema = "w" Then
        AsignarFonemaBase_CA = 22: Exit Function   ' /w/
    End If

    ' NASALS
    Select Case grafema
        Case "m": AsignarFonemaBase_CA = 36: Exit Function
        Case "n": AsignarFonemaBase_CA = 37: Exit Function
        Case "ny": AsignarFonemaBase_CA = 38: Exit Function
        Case "ng": AsignarFonemaBase_CA = 39: Exit Function
    End Select

    ' LATERALS
    If grafema = "l" Then AsignarFonemaBase_CA = 62: Exit Function
    If grafema = "ll" Then AsignarFonemaBase_CA = 63: Exit Function
    If grafema = "l·l" Then AsignarFonemaBase_CA = 63: Exit Function

    ' VIBRANTS
    If grafema = "rr" Then AsignarFonemaBase_CA = 60: Exit Function

    If grafema = "r" Then
        If ant = "" Or Not ant Like "[aeiouàèéíïòóúü]" Then
            AsignarFonemaBase_CA = 60   ' inicial / postconsonàntica ? múltiple
        Else
            AsignarFonemaBase_CA = 59   ' intervocàlica ? simple
        End If
        Exit Function
    End If

    ' AFRICADES
    If grafema = "tx" Then AsignarFonemaBase_CA = 57: Exit Function
    If grafema = "tg" Or grafema = "tj" Then AsignarFonemaBase_CA = 58: Exit Function

    ' FRICATIVES
    If grafema = "s" Then AsignarFonemaBase_CA = 42: Exit Function
    If grafema = "z" Then AsignarFonemaBase_CA = 43: Exit Function
    If grafema = "x" Then AsignarFonemaBase_CA = 44: Exit Function
    If grafema = "ix" Then AsignarFonemaBase_CA = 44: Exit Function
    If grafema = "j" Then AsignarFonemaBase_CA = 45: Exit Function
    If grafema = "ge" Or grafema = "gi" Then AsignarFonemaBase_CA = 45: Exit Function

    ' OCLUSIVES
    Select Case grafema
        Case "p": AsignarFonemaBase_CA = 30: Exit Function
        Case "b": AsignarFonemaBase_CA = 31: Exit Function
        Case "t": AsignarFonemaBase_CA = 32: Exit Function
        Case "d": AsignarFonemaBase_CA = 33: Exit Function
        Case "c", "k", "qu": AsignarFonemaBase_CA = 34: Exit Function
        Case "g", "gu": AsignarFonemaBase_CA = 35: Exit Function
    End Select

    ' No reconegut ? 255 (tu el marcaràs com vulguis)
    AsignarFonemaBase_CA = 255

End Function

' ============================================================
'   NORMALITZAR VOCALS — CATALÀ (neutre)
' ============================================================
Private Function NormalizarVocales_CA(ByVal texto As String) As String
    ' No corregim grafies no catalanes aquí.
    ' Si apareixen "á", "â", etc., quedaran com a grafemes no reconeguts (ID 255).
    NormalizarVocales_CA = texto
End Function



' ============================================================
' ============================================================
'                Reutilizadas del ES
' ============================================================
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

    ' Inicializar caché si es la primera vez
    If IPA_Cache Is Nothing Then
        Set IPA_Cache = CreateObject("Scripting.Dictionary")
    End If

    ' Si ya está en caché ? devolverlo directamente
    If IPA_Cache.Exists(idFonema) Then
        ObtenerIPA = IPA_Cache(idFonema)
        Exit Function
    End If

    ' ID desconocido o especial
    If idFonema = 255 Then
        IPA_Cache.Add idFonema, ""   ' vacío
        ObtenerIPA = ""
        Exit Function
    End If

    ' Buscar en la tabla qryFonemasValor
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT IPA FROM qryFonemasValor WHERE ID=" & idFonema & ";", _
        dbOpenSnapshot)

    If Not (rs.EOF And rs.BOF) Then
        IPA_Cache.Add idFonema, Nz(rs!ipa, "")
        ObtenerIPA = Nz(rs!ipa, "")
    Else
        ' Si no existe el ID ? devolver vacío
        IPA_Cache.Add idFonema, ""
        ObtenerIPA = ""
    End If

    rs.Close
    Set rs = Nothing
End Function



