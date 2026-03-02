Attribute VB_Name = "bas_Motor_VA_Fon"

Option Compare Database
Option Explicit

' ============================================================
'   MOTOR FONÉTICO — VALENCIANO (VA)
'   Construcción de IdsFonemas + IPA final
'   Arquitectura paralela a CA e IB
' ============================================================
Public Sub ConstruirCadenaFonemas_VA()

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

    ' 0. Normalización valenciana (no tocamos grafías)
    frase = NormalizarVocales(ObjDTO.SilabasFinal)

    ' 1. Separar sílabas
    arrSilabas = Split(frase, "|")

    For i = 0 To UBound(arrSilabas)

        silabaCruda = arrSilabas(i)

        ' 2. Separador de palabra
        If Trim$(silabaCruda) = "" Then

            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "#"

            If CFG.ModoLigadura = 2 Then
                ultimaSilaba = BuscarSilabaRealAnterior(arrSilabas, i)
                siguienteSilaba = BuscarSilabaRealPosterior(arrSilabas, i)

                If HayLigaduraAutomatica(ultimaSilaba, siguienteSilaba) Then
                    ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "84,"
                End If
            End If

            GoTo SiguienteIteracion
        End If

        ' 3. Prosodia
        If InStr(silabaCruda, "(") > 0 Then
            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "80,"
        ElseIf InStr(silabaCruda, "[") > 0 Then
            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "81,"
        Else
            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "82,"
        End If

        ' 4. Procesar grafemas
        ProcesarSilaba silabaCruda

        ' 5. Separador silábico
        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "|"

SiguienteIteracion:
    Next i

    ' 6. Limpieza final
    ObjDTO.IdsFonemas = Replace(ObjDTO.IdsFonemas, "|#", "#")
    ObjDTO.IdsFonemas = Replace(ObjDTO.IdsFonemas, "#|", "#")
    ObjDTO.IdsFonemas = Replace(ObjDTO.IdsFonemas, ",|", "|")
    
    ' 6.5 Asimilaciones valencianas
    ObjDTO.IdsFonemas = AplicarAssimilaciones(ObjDTO.IdsFonemas)


    While Right$(ObjDTO.IdsFonemas, 1) Like "[|#,]"
        ObjDTO.IdsFonemas = Left$(ObjDTO.IdsFonemas, Len(ObjDTO.IdsFonemas) - 1)
    Wend

    ' 7. Construcción IPA final
    arrSilabas = Split(ObjDTO.IdsFonemas, "|")
    res = ""

    For i = 0 To UBound(arrSilabas)

        If i = 0 Then res = res & "/"

        arrFon = Split(arrSilabas(i), ",")

        For Each Fon In arrFon

            If Fon = "#84" Then Fon = 84
            If Fon = "#82" Then res = res & "/ /": Fon = 82

            If Left$(Fon, 1) = "#" Then
                res = res & "/ /"
                Fon = Replace(Fon, "#", "")
            End If

            If Trim$(Fon) <> "" And Trim$(Fon) <> "82" Then
                res = res & Replace(ObtenerIPA(Fon), "/", "")
            End If

        Next Fon

        If i = UBound(arrSilabas) Then res = res & "/ "
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
    Dim esTonica As Boolean

    ' Detectar prosodia
    esAtona = True
    esTonica = False

    If InStr(silabaCruda, "(") > 0 Then
        esAtona = False
        esTonica = True
    End If

    If InStr(silabaCruda, "[") > 0 Then
        esAtona = False
    End If

    ' Limpiar marcadores
    silaba = silabaCruda
    silaba = Replace(silaba, "(", "")
    silaba = Replace(silaba, ")", "")
    silaba = Replace(silaba, "[", "")
    silaba = Replace(silaba, "]", "")
    silaba = Trim$(silaba)

    ' 1. Ligadura manual
    If CFG.ModoLigadura = 1 Then
        ligaduraID = DetectarLigaduraManual(silaba)
        If ligaduraID <> 0 Then
            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & ligaduraID & ","
        End If
    End If

    ' 2. Procesar grafemas
    i = 1
    Do While i <= Len(silaba)

        grafema = DetectarDigrafo(silaba, i)

        ' Contexto
        If i > 1 Then antCh = Mid$(silaba, i - 1, 1) Else antCh = ""
        If i + Len(grafema) <= Len(silaba) Then
            sigCh = Mid$(silaba, i + Len(grafema), 1)
        Else
            sigCh = ""
        End If

        ' 3. Apertura vocálica valenciana real
        If esTonica Then
            If grafema = "e" Then
                id = 3   ' ?
                GoTo Agregar
            End If
            If grafema = "o" Then
                id = 6   ' ?
                GoTo Agregar
            End If
        End If

        ' 4. Fonema base valenciano
        id = AsignarFonemaBase(grafema, sigCh, antCh)

Agregar:
        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & id & ","

        i = i + Len(grafema)

    Loop

End Sub

Private Function DetectarDigrafo(t As String, pos As Long) As String

    Dim L As Long: L = Len(t)
    Dim par As String, tri As String, sig As String

    ' 1. l·l
    If pos + 2 <= L Then
        tri = Mid$(t, pos, 3)
        If tri = "l·l" Then DetectarDigrafo = tri: Exit Function
    End If

    ' 2. Dígrafos valencianos
    If pos + 1 <= L Then
        par = Mid$(t, pos, 2)
        sig = Mid$(t, pos + 2, 1)

        Select Case par
            Case "ny", "ll", "rr", "tx", "tg", "tj", "ix"
                DetectarDigrafo = par: Exit Function

            Case "qu", "qü"
                DetectarDigrafo = par: Exit Function

            Case "gu"
                If sig = "e" Or sig = "i" Then DetectarDigrafo = par: Exit Function

            Case "gü"
                If sig = "e" Or sig = "i" Then DetectarDigrafo = par: Exit Function
        End Select
    End If

    ' 3. "ig" final ? /t??/
    If pos + 1 = L Then
        If Mid$(t, pos, 2) = "ig" Then DetectarDigrafo = "tx": Exit Function
    End If

    DetectarDigrafo = Mid$(t, pos, 1)

End Function

Private Function AsignarFonemaBase(grafema As String, _
                                      Optional sig As String = "", _
                                      Optional ant As String = "") As Byte

    grafema = LCase$(grafema)
    sig = LCase$(sig)
    ant = LCase$(ant)

    ' VOCALS
    Select Case grafema
        Case "a", "à": AsignarFonemaBase = 1: Exit Function
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
        If ant = "" Or Not ant Like "[aeiouàèéíïòóúü]" Then
            AsignarFonemaBase = 60
        Else
            AsignarFonemaBase = 59
        End If
        Exit Function
    End If

    ' AFRICADES
    If grafema = "tx" Then AsignarFonemaBase = 57: Exit Function
    If grafema = "tg" Or grafema = "tj" Then AsignarFonemaBase = 58: Exit Function

    ' FRICATIVES
    If grafema = "s" Then AsignarFonemaBase = 42: Exit Function
    If grafema = "z" Then AsignarFonemaBase = 43: Exit Function
    If grafema = "x" Then AsignarFonemaBase = 44: Exit Function
    If grafema = "ix" Then AsignarFonemaBase = 44: Exit Function
    If grafema = "j" Then AsignarFonemaBase = 45: Exit Function
    If grafema = "ge" Or grafema = "gi" Then AsignarFonemaBase = 45: Exit Function

    ' OCLUSIVES
    Select Case grafema
        Case "v": AsignarFonemaBase = 41: Exit Function   ' /v/
        Case "p": AsignarFonemaBase = 30: Exit Function
        Case "b": AsignarFonemaBase = 31: Exit Function
        Case "t": AsignarFonemaBase = 32: Exit Function
        Case "d": AsignarFonemaBase = 33: Exit Function
        Case "c", "k", "qu": AsignarFonemaBase = 34: Exit Function
        Case "g", "gu": AsignarFonemaBase = 35: Exit Function
    End Select

    AsignarFonemaBase = 255

End Function

Private Function AplicarAssimilaciones(cadena As String) As String
    Dim silabas As Variant
    Dim fonemas As Variant
    Dim i As Long, j As Long
    Dim id As String
    Dim res As String
    
    ' Separar por sílabas
    silabas = Split(cadena, "|")
    
    For i = 0 To UBound(silabas)
        
        If Trim$(silabas(i)) <> "" Then
            fonemas = Split(silabas(i), ",")
            
            ' Recorremos fonemas dentro de la sílaba
            For j = 0 To UBound(fonemas)
                id = Trim$(fonemas(j))
                
                ' --- S intervocálica ? Z (42 ? 43) ---
                ' Miramos contexto: vocal (1–7) + 42 + vocal (1–7)
                If id = "42" Then
                    If j > 0 And j < UBound(fonemas) Then
                        If EsVocalID(fonemas(j - 1)) And EsVocalID(fonemas(j + 1)) Then
                            fonemas(j) = "43"
                        End If
                    End If
                End If
                
                ' --- ns ? s (37 42 ? 42) ---
                If id = "42" And j > 0 Then
                    If Trim$(fonemas(j - 1)) = "37" Then
                        fonemas(j - 1) = ""   ' borramos la n
                    End If
                End If
                
                ' --- mb ? m (36 31 ? 36) ---
                If id = "31" And j > 0 Then
                    If Trim$(fonemas(j - 1)) = "36" Then
                        fonemas(j) = ""       ' borramos la b
                    End If
                End If
                
                ' --- nk ? ?k (37 34 ? 39 34) ---
                If id = "34" And j > 0 Then
                    If Trim$(fonemas(j - 1)) = "37" Then
                        fonemas(j - 1) = "39"
                    End If
                End If
                
                ' --- ng ? ?g (37 35 ? 39 35) ---
                If id = "35" And j > 0 Then
                    If Trim$(fonemas(j - 1)) = "37" Then
                        fonemas(j - 1) = "39"
                    End If
                End If
                
                ' --- ld ? l (62 33 ? 62) ---
                If id = "33" And j > 0 Then
                    If Trim$(fonemas(j - 1)) = "62" Then
                        fonemas(j) = ""
                    End If
                End If
                
                ' --- l·l ? ? (62 62 ? 63) ---
                If id = "62" And j > 0 Then
                    If Trim$(fonemas(j - 1)) = "62" Then
                        fonemas(j - 1) = "63"
                        fonemas(j) = ""
                    End If
                End If
                
            Next j
            
            ' Reconstruir sílaba sin huecos vacíos
            silabas(i) = ReconstruirSilabaDesdeArray(fonemas)
        End If
        
    Next i
    
    ' Reconstruir cadena completa
    res = ""
    For i = 0 To UBound(silabas)
        If i > 0 Then res = res & "|"
        res = res & silabas(i)
    Next i
    
    AplicarAssimilaciones = res
End Function

Private Function EsVocalID(ByVal id As String) As Boolean
    id = Trim$(id)
    EsVocalID = (id = "1" Or id = "2" Or id = "3" Or _
                 id = "4" Or id = "5" Or id = "6" Or id = "7" Or id = "8")
End Function

Private Function ReconstruirSilabaDesdeArray(fonemas As Variant) As String
    Dim j As Long
    Dim tmp As String
    
    For j = 0 To UBound(fonemas)
        If Trim$(fonemas(j)) <> "" Then
            If tmp <> "" Then tmp = tmp & ","
            tmp = tmp & Trim$(fonemas(j))
        End If
    Next j
    
    ReconstruirSilabaDesdeArray = tmp
End Function

Private Function NormalizarVocales(ByVal texto As String) As String
    NormalizarVocales = texto
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
