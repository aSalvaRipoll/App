Attribute VB_Name = "bas_Motor_CA_Fon"

Option Compare Database
Option Explicit

' ============================================================
'   MOTOR FONOLÒGIC CATALÀ — INDEPENDENT
' ============================================================

Public Sub ConstruirCadenaFonemas_CA()

    Dim arrSilabas As Variant
    Dim i As Long
    Dim silaba As Variant
    Dim ultimaSilaba As String
    Dim siguienteSilaba As String
    Dim frase As String
    Dim arrFon As Variant
    Dim Fon As Variant
    Dim res As String

    ObjDTO.IdsFonemas = ""
    ObjDTO.FonemasFinal = ""

    ' Normalització catalana
    frase = NormalizarVocales_CA(ObjDTO.SilabasFinal)

    arrSilabas = Split(frase, "|")

    For i = 0 To UBound(arrSilabas)

        silaba = Trim$(arrSilabas(i))

        ' Separador de paraula
        If silaba = "" Then

            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "#"

            If CFG.ModoLigadura = 2 Then
                ultimaSilaba = BuscarSilabaRealAnterior_CA(arrSilabas, i)
                siguienteSilaba = BuscarSilabaRealPosterior_CA(arrSilabas, i)

                If HayLigaduraAutomatica_CA(ultimaSilaba, siguienteSilaba) Then
                    ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "84,"
                End If
            End If

            GoTo SiguienteIteracion
        End If

        ' Modificadors prosòdics
        If InStr(silaba, "(") > 0 Then
            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "80,"
        ElseIf InStr(silaba, "[") > 0 Then
            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "81,"
        Else
            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "82,"
        End If

        silaba = Replace(Replace(Replace(Replace(silaba, "(", ""), ")", ""), "[", ""), "]", "")

        ' Processar grafemes
        ProcesarSilaba_CA silaba

        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "|"

SiguienteIteracion:
    Next i

    ' Neteja final
    ObjDTO.IdsFonemas = Replace(ObjDTO.IdsFonemas, "|#", "#")
    ObjDTO.IdsFonemas = Replace(ObjDTO.IdsFonemas, "#|", "#")

    If Right$(ObjDTO.IdsFonemas, 1) Like "[|,#]" Then
        ObjDTO.IdsFonemas = Left$(ObjDTO.IdsFonemas, Len(ObjDTO.IdsFonemas) - 1)
    End If

    ObjDTO.IdsFonemas = Replace(ObjDTO.IdsFonemas, ",|", "|")

    ' Reconstrucció IPA
    ObjDTO.FonemasFinal = GenerarIPA()

End Sub


' ============================================================
'   PROCESSAR SÍL·LABA
' ============================================================
Private Sub ProcesarSilaba_CA(ByVal silabaCruda As String)

    Dim silaba As String
    Dim ligaduraID As Byte
    Dim i As Long
    Dim grafema As String
    Dim id As Byte
    Dim antCh As String
    Dim sigCh As String

    silaba = Trim$(silabaCruda)

    If CFG.ModoLigadura = 1 Then
        ligaduraID = DetectarLigaduraManual(silaba)
        If ligaduraID <> 0 Then
            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & ligaduraID & ","
        End If
    End If

    i = 1
    Do While i <= Len(silaba)

        grafema = DetectarDigrafo(silaba, i)

        If i > 1 Then antCh = Mid$(silaba, i - 1, 1) Else antCh = ""
        If i + Len(grafema) <= Len(silaba) Then
            sigCh = Mid$(silaba, i + Len(grafema), 1)
        Else
            sigCh = ""
        End If

        id = AsignarFonemaBaseCa(grafema, sigCh, antCh)

        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & id & ","

        i = i + Len(grafema)

    Loop

End Sub

Private Function BuscarSilabaRealAnterior_CA(arr As Variant, pos As Long) As String
    Dim j As Long
    For j = pos - 1 To 0 Step -1
        If Trim$(arr(j)) <> "" Then
            BuscarSilabaRealAnterior_CA = Trim$(arr(j))
            Exit Function
        End If
    Next j
    BuscarSilabaRealAnterior_CA = ""
End Function


Private Function BuscarSilabaRealPosterior_CA(arr As Variant, pos As Long) As String
    Dim j As Long
    For j = pos + 1 To UBound(arr)
        If Trim$(arr(j)) <> "" Then
            BuscarSilabaRealPosterior_CA = Trim$(arr(j))
            Exit Function
        End If
    Next j
    BuscarSilabaRealPosterior_CA = ""
End Function

' ============================================================
'   DETECTAR DÍGRAFS CATALANS
' ============================================================
Private Function DetectarDigrafo(texto As String, pos As Long) As String

    Dim L As Long
    Dim par As String, tri As String

    L = Len(texto)

    If pos >= L Then
        DetectarDigrafo = Mid$(texto, pos, 1)
        Exit Function
    End If

    ' Trígraf l·l
    If pos + 2 <= L Then
        tri = Mid$(texto, pos, 3)
        If tri = "l·l" Then
            DetectarDigrafo = tri
            Exit Function
        End If
    End If

    par = Mid$(texto, pos, 2)

    Select Case par
        Case "ny", "ll", "tx", "tg", "tj", "ix", "ss", "qu", "gu", "rr", "ch"
            DetectarDigrafo = par
            Exit Function
    End Select

    DetectarDigrafo = Mid$(texto, pos, 1)

End Function

Private Function DetectarLigaduraManual(ByRef silaba As String) As Byte

    DetectarLigaduraManual = 0
    
    If InStr(silaba, "_") > 0 Then
        silaba = Replace(silaba, "_", "")
        DetectarLigaduraManual = 84
    End If

End Function


' ============================================================
'   ASSIGNACIÓ DE FONEMES CATALANS (IDs REALS)
' ============================================================
Function AsignarFonemaBaseCa(grafema As String, _
                             Optional ByVal sig As String = "", _
                             Optional ByVal ant As String = "") As Byte

    grafema = LCase$(grafema)
    sig = LCase$(sig)
    ant = LCase$(ant)

    ' VOCALS
    Select Case grafema
        Case "a", "á", "à", "â": AsignarFonemaBaseCa = 1: Exit Function
        Case "e", "é", "ê": AsignarFonemaBaseCa = 2: Exit Function
        Case "è": AsignarFonemaBaseCa = 3: Exit Function
        Case "i", "í", "ï": AsignarFonemaBaseCa = 4: Exit Function
        Case "o", "ó", "ô": AsignarFonemaBaseCa = 5: Exit Function
        Case "ò": AsignarFonemaBaseCa = 6: Exit Function
        Case "u", "ú", "ü": AsignarFonemaBaseCa = 7: Exit Function
    End Select

    ' SEMIVOCALS
    If grafema = "i" And sig Like "[aeou]" Then AsignarFonemaBaseCa = 21: Exit Function
    If grafema = "u" And sig Like "[aeio]" Then AsignarFonemaBaseCa = 22: Exit Function

    ' NASALS
    Select Case grafema
        Case "m": AsignarFonemaBaseCa = 36: Exit Function
        Case "n": AsignarFonemaBaseCa = 37: Exit Function
        Case "ny": AsignarFonemaBaseCa = 38: Exit Function
        Case "ng": AsignarFonemaBaseCa = 39: Exit Function
    End Select

    ' LATERALS
    If grafema = "l" Then AsignarFonemaBaseCa = 62: Exit Function
    If grafema = "ll" Or grafema = "l·l" Then AsignarFonemaBaseCa = 63: Exit Function

    ' FRICATIVES
    If grafema = "f" Then AsignarFonemaBaseCa = 40: Exit Function
    If grafema = "v" Then AsignarFonemaBaseCa = 41: Exit Function

    Select Case grafema
        Case "s", "ss", "c", "ç": AsignarFonemaBaseCa = 42: Exit Function
        Case "z": AsignarFonemaBaseCa = 43: Exit Function
    End Select

    If grafema = "x" Or grafema = "ix" Then AsignarFonemaBaseCa = 44: Exit Function
    If grafema = "j" Then AsignarFonemaBaseCa = 45: Exit Function
    If grafema = "g" And (sig = "e" Or sig = "i") Then AsignarFonemaBaseCa = 45: Exit Function

    ' AFRICADES
    If grafema = "tx" Then AsignarFonemaBaseCa = 57: Exit Function
    If grafema = "tg" Or grafema = "tj" Then AsignarFonemaBaseCa = 58: Exit Function

    ' OCLUSIVES
    Select Case grafema
        Case "p": AsignarFonemaBaseCa = 30: Exit Function
        Case "b": AsignarFonemaBaseCa = 31: Exit Function
        Case "t": AsignarFonemaBaseCa = 32: Exit Function
        Case "d": AsignarFonemaBaseCa = 33: Exit Function
        Case "k", "c", "qu": AsignarFonemaBaseCa = 34: Exit Function
        Case "g": AsignarFonemaBaseCa = 35: Exit Function
    End Select

    ' VIBRANTS
    If grafema = "rr" Then AsignarFonemaBaseCa = 60: Exit Function

    If grafema = "r" Then
        If ant = "" Or ant Like "[bcdfghjklmnpqrstvwxyz]" Then
            AsignarFonemaBaseCa = 60
        Else
            AsignarFonemaBaseCa = 59
        End If
        Exit Function
    End If

    ' H muda o aspirada
    If grafema = "h" Then AsignarFonemaBaseCa = 49: Exit Function

    AsignarFonemaBaseCa = 255

End Function


' ============================================================
'   NORMALITZACIÓ CATALANA
' ============================================================
Function NormalizarVocales_CA(ByVal texto As String) As String

    texto = LCase$(texto)

    texto = Replace(texto, "á", "a")
    texto = Replace(texto, "â", "a")

    texto = Replace(texto, "é", "e")
    texto = Replace(texto, "ê", "e")

    texto = Replace(texto, "ó", "o")
    texto = Replace(texto, "ô", "o")

    texto = Replace(texto, "ú", "u")
    texto = Replace(texto, "û", "u")

    texto = Replace(texto, "h", "")

    texto = Replace(texto, "’", "'")

    NormalizarVocales_CA = texto

End Function


' ============================================================
'   LIGADURA AUTOMÀTICA
' ============================================================
Private Function HayLigaduraAutomatica_CA(ultimaSilaba As String, primeraSilaba As String) As Boolean

    Dim ult As String
    Dim pri As String

    HayLigaduraAutomatica_CA = False

    If ultimaSilaba = "" Or primeraSilaba = "" Then Exit Function

    ult = Right$(Trim$(ultimaSilaba), 1)
    pri = Left$(Trim$(primeraSilaba), 1)

    If ult Like "[aeiouàèéíïòóúü]" And pri Like "[aeiouàèéíïòóúü]" Then
        HayLigaduraAutomatica_CA = True
        Exit Function
    End If

    If ult Like "[aeiouàèéíïòóúü]" And primeraSilaba Like "h[aeiouàèéíïòóúü]*" Then
        HayLigaduraAutomatica_CA = True
        Exit Function
    End If

End Function

Private Function GenerarIPA() As String

    Dim salida As String
    Dim arrPalabras As Variant
    Dim arrSilabas As Variant
    Dim arrFonemas As Variant
    Dim palabra As Variant
    Dim silaba As Variant
    Dim f As Variant
    Dim id As Long
    Dim ipa As String

    salida = ""

    ' 1. Separar palabras por "#"
    arrPalabras = Split(ObjDTO.IdsFonemas, "#")

    For Each palabra In arrPalabras

        If Trim$(palabra) = "" Then GoTo SiguientePalabra

        ' 2. Separar sílabas por "|"
        arrSilabas = Split(palabra, "|")

        For Each silaba In arrSilabas

            If Trim$(silaba) = "" Then GoTo siguienteSilaba

            ' 3. Separar fonemas/modificadores por ","
            arrFonemas = Split(Trim$(silaba), ",")

            ipa = ""

            For Each f In arrFonemas

                If Trim$(f) = "" Then GoTo SiguienteFonema

                If IsNumeric(f) Then
                    id = CLng(f)

                    Select Case id

                        ' ------------------------------
                        ' MODIFICADORES PROSÓDICOS
                        ' ------------------------------
                        Case 80: ipa = ipa & "'"     ' acento primario
                        Case 81: ipa = ipa & "?"     ' acento secundario
                        Case 82:                     ' átona ? no se muestra
                        Case 84: ipa = ipa & "?"     ' ligadura

                        ' ------------------------------
                        ' FONEMAS REALES
                        ' ------------------------------
                        Case Else
                            ipa = ipa & ObtenerIPA(id)

                    End Select

                End If

SiguienteFonema:
            Next f

            ' 4. Añadir punto silábico
            salida = salida & ipa & "."

siguienteSilaba:
        Next silaba

        ' 5. Espacio entre palabras
        salida = salida & " "

SiguientePalabra:
    Next palabra

    ' 6. Limpieza final
    salida = Trim$(salida)

    ' Quitar punto final si existe
    If Right$(salida, 1) = "." Then
        salida = Left$(salida, Len(salida) - 1)
    End If

    GenerarIPA = salida

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


