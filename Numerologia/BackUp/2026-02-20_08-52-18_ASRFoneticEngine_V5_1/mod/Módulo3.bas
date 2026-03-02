Attribute VB_Name = "Módulo3"
'Option Compare Database
'Option Explicit
'
'Public Sub ConstruirCadenaFonemas_VA()
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
'    ' 0. Normalización valenciana (no tocamos grafías)
'    frase = NormalizarVocales_VA(ObjDTO.SilabasFinal)
'
'    ' 1. Separar sílabas
'    arrSilabas = Split(frase, "|")
'
'    For i = 0 To UBound(arrSilabas)
'
'        silabaCruda = arrSilabas(i)
'
'        ' 2. Separador de palabra
'        If Trim$(silabaCruda) = "" Then
'
'            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "#"
'
'            If CFG.ModoLigadura = 2 Then
'                ultimaSilaba = BuscarSilabaRealAnterior(arrSilabas, i)
'                siguienteSilaba = BuscarSilabaRealPosterior(arrSilabas, i)
'
'                If HayLigaduraAutomatica(ultimaSilaba, siguienteSilaba) Then
'                    ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "84,"
'                End If
'            End If
'
'            GoTo Siguiente
'        End If
'
'        ' 3. Prosodia
'        If InStr(silabaCruda, "(") > 0 Then
'            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "80,"
'        ElseIf InStr(silabaCruda, "[") > 0 Then
'            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "81,"
'        Else
'            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "82,"
'        End If
'
'        ' 4. Procesar grafemas
'        ProcesarSilaba_VA silabaCruda
'
'        ' 5. Separador silábico
'        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "|"
'
'Siguiente:
'    Next i
'
'    ' 6. Limpieza final
'    ObjDTO.IdsFonemas = Replace(ObjDTO.IdsFonemas, "|#", "#")
'    ObjDTO.IdsFonemas = Replace(ObjDTO.IdsFonemas, "#|", "#")
'    ObjDTO.IdsFonemas = Replace(ObjDTO.IdsFonemas, ",|", "|")
'
'    While Right$(ObjDTO.IdsFonemas, 1) Like "[|#,]"
'        ObjDTO.IdsFonemas = Left$(ObjDTO.IdsFonemas, Len(ObjDTO.IdsFonemas) - 1)
'    Wend
'
'    ' 7. Construcción IPA
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
'            If Fon = "#84" Then Fon = 84
'            If Fon = "#82" Then res = res & "/ /": Fon = 82
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
'        If i = UBound(arrSilabas) Then res = res & "/ "
'    Next i
'
'    ObjDTO.FonemasFinal = Trim$(res)
'
'End Sub
'
'Private Sub ProcesarSilaba_VA(ByVal silabaCruda As String)
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
'    esAtona = True
'    If InStr(silabaCruda, "(") > 0 Then esAtona = False
'    If InStr(silabaCruda, "[") > 0 Then esAtona = False
'
'    silaba = silabaCruda
'    silaba = Replace(silaba, "(", "")
'    silaba = Replace(silaba, ")", "")
'    silaba = Replace(silaba, "[", "")
'    silaba = Replace(silaba, "]", "")
'    silaba = Trim$(silaba)
'
'    ' 1. Ligadura manual
'    If CFG.ModoLigadura = 1 Then
'        ligaduraID = DetectarLigaduraManual(silaba)
'        If ligaduraID <> 0 Then
'            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & ligaduraID & ","
'        End If
'    End If
'
'    If Not esAtona Then
'        If grafema = "e" Then id = 3   ' ?
'        If grafema = "o" Then id = 6   ' ?
'    End If
'
'    ' 2. Procesar grafemas
'    i = 1
'    Do While i <= Len(silaba)
'
'        grafema = DetectarDigrafo_VA(silaba, i)
'
'        If i > 1 Then antCh = Mid$(silaba, i - 1, 1) Else antCh = ""
'        If i + Len(grafema) <= Len(silaba) Then
'            sigCh = Mid$(silaba, i + Len(grafema), 1)
'        Else
'            sigCh = ""
'        End If
'
'        ' 3. Fonema base valenciano
'        id = AsignarFonemaBase_VA(grafema, sigCh, antCh)
'
'        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & id & ","
'
'        i = i + Len(grafema)
'
'    Loop
'
'End Sub
'
'Private Function DetectarDigrafo_VA(t As String, pos As Long) As String
'
'    Dim L As Long: L = Len(t)
'    Dim par As String, tri As String, sig As String
'
'    ' 1. l·l
'    If pos + 2 <= L Then
'        tri = Mid$(t, pos, 3)
'        If tri = "l·l" Then DetectarDigrafo_VA = tri: Exit Function
'    End If
'
'    ' 2. Dígrafos valencianos
'    If pos + 1 <= L Then
'        par = Mid$(t, pos, 2)
'        sig = Mid$(t, pos + 2, 1)
'
'        Select Case par
'            Case "ny", "ll", "rr", "tx", "tg", "tj", "ix"
'                DetectarDigrafo_VA = par: Exit Function
'
'            Case "qu", "qü"
'                DetectarDigrafo_VA = par: Exit Function
'
'            Case "gu"
'                If sig = "e" Or sig = "i" Then DetectarDigrafo_VA = par: Exit Function
'
'            Case "gü"
'                If sig = "e" Or sig = "i" Then DetectarDigrafo_VA = par: Exit Function
'        End Select
'    End If
'
'    ' 3. "ig" final ? /t??/
'    If pos + 1 = L Then
'        If Mid$(t, pos, 2) = "ig" Then DetectarDigrafo_VA = "tx": Exit Function
'    End If
'
'    DetectarDigrafo_VA = Mid$(t, pos, 1)
'
'End Function
'
'Private Function AsignarFonemaBase_VA(grafema As String, _
'                                      Optional sig As String = "", _
'                                      Optional ant As String = "") As Byte
'
'    grafema = LCase$(grafema)
'    sig = LCase$(sig)
'    ant = LCase$(ant)
'
'    ' VOCALS
'    Select Case grafema
'        Case "a", "à": AsignarFonemaBase_VA = 1: Exit Function
'        Case "e", "é": AsignarFonemaBase_VA = 2: Exit Function
'        Case "è": AsignarFonemaBase_VA = 3: Exit Function
'        Case "i", "í", "ï": AsignarFonemaBase_VA = 4: Exit Function
'        Case "o", "ó": AsignarFonemaBase_VA = 5: Exit Function
'        Case "ò": AsignarFonemaBase_VA = 6: Exit Function
'        Case "u", "ú", "ü": AsignarFonemaBase_VA = 7: Exit Function
'    End Select
'
'    ' SEMIVOCALS
'    If grafema = "j" Or grafema = "y" Then AsignarFonemaBase_VA = 21: Exit Function
'    If grafema = "w" Then AsignarFonemaBase_VA = 22: Exit Function
'
'    ' NASALS
'    Select Case grafema
'        Case "m": AsignarFonemaBase_VA = 36: Exit Function
'        Case "n": AsignarFonemaBase_VA = 37: Exit Function
'        Case "ny": AsignarFonemaBase_VA = 38: Exit Function
'        Case "ng": AsignarFonemaBase_VA = 39: Exit Function
'    End Select
'
'    ' LATERALS
'    If grafema = "l" Then AsignarFonemaBase_VA = 62: Exit Function
'    If grafema = "ll" Then AsignarFonemaBase_VA = 63: Exit Function
'    If grafema = "l·l" Then AsignarFonemaBase_VA = 63: Exit Function
'
'    ' VIBRANTS
'    If grafema = "rr" Then AsignarFonemaBase_VA = 60: Exit Function
'
'    If grafema = "r" Then
'        If ant = "" Or Not ant Like "[aeiouàèéíïòóúü]" Then
'            AsignarFonemaBase_VA = 60
'        Else
'            AsignarFonemaBase_VA = 59
'        End If
'        Exit Function
'    End If
'
'    ' AFRICADES
'    If grafema = "tx" Then AsignarFonemaBase_VA = 57: Exit Function
'    If grafema = "tg" Or grafema = "tj" Then AsignarFonemaBase_VA = 58: Exit Function
'
'    ' FRICATIVES
'    If grafema = "s" Then AsignarFonemaBase_VA = 42: Exit Function
'    If grafema = "z" Then AsignarFonemaBase_VA = 43: Exit Function
'    If grafema = "x" Then AsignarFonemaBase_VA = 44: Exit Function
'    If grafema = "ix" Then AsignarFonemaBase_VA = 44: Exit Function
'    If grafema = "j" Then AsignarFonemaBase_VA = 45: Exit Function
'    If grafema = "ge" Or grafema = "gi" Then AsignarFonemaBase_VA = 45: Exit Function
'
'    ' OCLUSIVES
'    Select Case grafema
'        Case "v": AsignarFonemaBase_VA = 41: Exit Function   ' /v/
'        Case "p": AsignarFonemaBase_VA = 30: Exit Function
'        Case "b": AsignarFonemaBase_VA = 31: Exit Function
'        Case "t": AsignarFonemaBase_VA = 32: Exit Function
'        Case "d": AsignarFonemaBase_VA = 33: Exit Function
'        Case "c", "k", "qu": AsignarFonemaBase_VA = 34: Exit Function
'        Case "g", "gu": AsignarFonemaBase_VA = 35: Exit Function
'    End Select
'
'    AsignarFonemaBase_VA = 255
'
'End Function
'
'Private Function NormalizarVocales_VA(ByVal texto As String) As String
'    NormalizarVocales_VA = texto
'End Function
'
'
