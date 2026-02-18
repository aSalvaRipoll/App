Attribute VB_Name = "Módulo2"

'Option Compare Database
'Option Explicit
'
'' ============================================================
''   MOTOR FONÉTICO — CATALÁN
''   Construcción de IdsFonemas + IPA final
''   MISMA ARQUITECTURA QUE EL MOTOR ES
'' ============================================================
'
'Public Sub ConstruirCadenaFonemas_CA()
'
'    Dim arrSilabas As Variant
'    Dim i As Long
'    Dim silaba As Variant
'    Dim ultimaSilaba As String
'    Dim siguienteSilaba As String
'
'    Dim frase As String
'    Dim arrFon As Variant
'    Dim Fon As Variant
'
'    Dim res As String
'
'    ObjDTO.IdsFonemas = ""
'    ObjDTO.FonemasFinal = ""
'
'    '0 Normalizar vocales catalanas
'    frase = NormalizarVocales_CA(ObjDTO.SilabasFinal)
'
'    ' 1. Separar sílabas por "|"
'    arrSilabas = Split(frase, "|")
'
'    For i = 0 To UBound(arrSilabas)
'
'        silaba = Trim$(arrSilabas(i))
'
'        ' ---------------------------------------------------------
'        ' 2. Detectar separador de palabra (sílaba vacía)
'        ' ---------------------------------------------------------
'        If silaba = "" Then
'
'            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "#"
'
'            ' Ligadura automática (Modo 2)
'            If CFG.ModoLigadura = 2 Then
'
'                ultimaSilaba = BuscarSilabaRealAnterior(arrSilabas, i)
'                siguienteSilaba = BuscarSilabaRealPosterior(arrSilabas, i)
'
'                If HayLigaduraAutomatica(ultimaSilaba, siguienteSilaba) Then
'                    ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "84,"
'                End If
'
'            End If
'
'            GoTo SiguienteIteracion
'        End If
'
'        ' ---------------------------------------------------------
'        ' 3. Insertar modificadores prosódicos (acento)
'        ' ---------------------------------------------------------
'        If InStr(silaba, "(") > 0 Then
'            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "80,"
'        ElseIf InStr(silaba, "[") > 0 Then
'            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "81,"
'        Else
'            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "82,"
'        End If
'
'        ' Limpiar marcadores
'        silaba = Replace(silaba, "(", "")
'        silaba = Replace(silaba, ")", "")
'        silaba = Replace(silaba, "[", "")
'        silaba = Replace(silaba, "]", "")
'
'        ' ---------------------------------------------------------
'        ' 4. Procesar grafemas (versión catalana)
'        ' ---------------------------------------------------------
'        ProcesarSilaba_CA silaba
'
'        ' ---------------------------------------------------------
'        ' 5. Separador silábico
'        ' ---------------------------------------------------------
'        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "|"
'
'SiguienteIteracion:
'    Next i
'
'    ' ---------------------------------------------------------
'    ' 6. Limpieza final (idéntica al motor ES)
'    ' ---------------------------------------------------------
'    ObjDTO.IdsFonemas = Replace(ObjDTO.IdsFonemas, "|#", "#")
'    ObjDTO.IdsFonemas = Replace(ObjDTO.IdsFonemas, "#|", "#")
'
'    If Right$(ObjDTO.IdsFonemas, 1) = "|" Or _
'       Right$(ObjDTO.IdsFonemas, 1) = "#" Or _
'       Right$(ObjDTO.IdsFonemas, 1) = "," Then
'
'        ObjDTO.IdsFonemas = Left$(ObjDTO.IdsFonemas, Len(ObjDTO.IdsFonemas) - 1)
'    End If
'
'    ObjDTO.IdsFonemas = Replace(ObjDTO.IdsFonemas, ",|", "|")
'
'    ' ---------------------------------------------------------
'    ' 7. Construcción IPA final (idéntica al ES)
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
''   PROCESAR SÍLABA — CATALÁN
''   (Clon del ES, adaptado a dígrafos catalanes)
'' ============================================================
'
'Private Sub ProcesarSilaba_CA(ByVal silabaCruda As String)
'
'    Dim silaba As String
'    Dim ligaduraID As Byte
'    Dim i As Long
'    Dim grafema As String
'    Dim id As Byte
'    Dim antCh As String
'    Dim sigCh As String
'
'    silaba = Trim$(silabaCruda)
'
'    ' 1. Detectar ligadura manual (solo Modo 1)
'    If CFG.ModoLigadura = 1 Then
'        ligaduraID = DetectarLigaduraManual(silaba)
'        If ligaduraID <> 0 Then
'            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & CStr(ligaduraID) & ","
'        End If
'    End If
'
'    ' 2. Procesar grafemas (con detección real de dígrafos catalanes)
'    i = 1
'    Do While i <= Len(silaba)
'
'        '-----------------------------------------
'        ' Detectar grafema (1, 2 o 3 letras)
'        '-----------------------------------------
'        grafema = DetectarDigrafo_CA(silaba, i)
'
'        '-----------------------------------------
'        ' Calcular contexto anterior
'        '-----------------------------------------
'        If i > 1 Then
'            antCh = Mid$(silaba, i - 1, 1)
'        Else
'            antCh = ""
'        End If
'
'        '-----------------------------------------
'        ' Calcular contexto siguiente
'        '-----------------------------------------
'        If i + Len(grafema) <= Len(silaba) Then
'            sigCh = Mid$(silaba, i + Len(grafema), 1)
'        Else
'            sigCh = ""
'        End If
'
'        '-----------------------------------------
'        ' Obtener fonema (ID catalán)
'        '-----------------------------------------
'        id = AsignarFonemaBase_CA(grafema, sigCh, antCh)
'
'        '-----------------------------------------
'        ' Añadir fonema
'        '-----------------------------------------
'        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & CStr(id) & ","
'
'        '-----------------------------------------
'        ' Avanzar según tamaño del grafema
'        '-----------------------------------------
'        i = i + Len(grafema)
'
'    Loop
'
'End Sub
'
'' ============================================================
''   DETECTAR DÍGRAFOS — CATALÁN
'' ============================================================
'
'Private Function DetectarDigrafo_CA(t As String, pos As Long) As String
'
'    Dim L As Long: L = Len(t)
'    Dim par As String, tri As String, sig As String
'
'    ' 1. Trigrama posible: "l·l"
'    If pos + 2 <= L Then
'        tri = Mid$(t, pos, 3)
'        If tri = "l·l" Then
'            DetectarDigrafo_CA = tri
'            Exit Function
'        End If
'    End If
'
'    ' 2. Dígrafos de 2 letras
'    If pos + 1 <= L Then
'        par = Mid$(t, pos, 2)
'        sig = Mid$(t, pos + 2, 1)
'
'        Select Case par
'
'            ' Laterales
'            Case "ll": DetectarDigrafo_CA = par: Exit Function
'
'            ' Nasal palatal
'            Case "ny": DetectarDigrafo_CA = par: Exit Function
'
'            ' Africada sorda
'            Case "tx": DetectarDigrafo_CA = par: Exit Function
'
'            ' Africada sonora
'            Case "tg", "tj": DetectarDigrafo_CA = par: Exit Function
'
'            ' Fricativa palatal catalana
'            Case "ix": DetectarDigrafo_CA = par: Exit Function
'
'            ' Qu / qü
'            Case "qu": DetectarDigrafo_CA = par: Exit Function
'            Case "qü": DetectarDigrafo_CA = par: Exit Function
'
'            ' Gu / güe / güi
'            Case "gu"
'                If sig = "e" Or sig = "i" Then
'                    DetectarDigrafo_CA = par
'                    Exit Function
'                End If
'
'            Case "gü"
'                If sig = "e" Or sig = "i" Then
'                    DetectarDigrafo_CA = par
'                    Exit Function
'                End If
'
'        End Select
'    End If
'
'    ' 3. "ig" final ? /t??/
'    If pos + 1 = L Then
'        If Mid$(t, pos, 2) = "ig" Then
'            DetectarDigrafo_CA = "tx"   ' se trata como /t??/
'            Exit Function
'        End If
'    End If
'
'    ' 4. Si no es dígrafo ? 1 letra
'    DetectarDigrafo_CA = Mid$(t, pos, 1)
'
'End Function
'
'' ============================================================
''   ASIGNAR FONEMA BASE — CATALÁN
'' ============================================================
'
'Private Function AsignarFonemaBase_CA(grafema As String, _
'                                      Optional sig As String = "", _
'                                      Optional ant As String = "") As Byte
'
'    grafema = LCase$(grafema)
'    sig = LCase$(sig)
'    ant = LCase$(ant)
'
'    ' -------------------------
'    ' VOCALS
'    ' -------------------------
'    Select Case grafema
'        Case "a": AsignarFonemaBase_CA = 1: Exit Function
'        Case "e": AsignarFonemaBase_CA = 2: Exit Function
'        Case "è": AsignarFonemaBase_CA = 3: Exit Function
'        Case "é": AsignarFonemaBase_CA = 2: Exit Function
'        Case "i": AsignarFonemaBase_CA = 4: Exit Function
'        Case "í": AsignarFonemaBase_CA = 4: Exit Function
'        Case "ï": AsignarFonemaBase_CA = 4: Exit Function
'        Case "o": AsignarFonemaBase_CA = 5: Exit Function
'        Case "ó": AsignarFonemaBase_CA = 5: Exit Function
'        Case "ò": AsignarFonemaBase_CA = 6: Exit Function
'        Case "u": AsignarFonemaBase_CA = 7: Exit Function
'        Case "ú": AsignarFonemaBase_CA = 7: Exit Function
'        Case "ü": AsignarFonemaBase_CA = 7: Exit Function
'        Case "a", "e": AsignarFonemaBase_CA = 8: Exit Function   ' schwa
'    End Select
'
'    ' -------------------------
'    ' SEMIVOCALS
'    ' -------------------------
'    If grafema = "j" Or grafema = "y" Then
'        AsignarFonemaBase_CA = 21: Exit Function
'    End If
'
'    If grafema = "w" Then
'        AsignarFonemaBase_CA = 22: Exit Function
'    End If
'
'    ' -------------------------
'    ' NASALS
'    ' -------------------------
'    Select Case grafema
'        Case "m": AsignarFonemaBase_CA = 36: Exit Function
'        Case "n": AsignarFonemaBase_CA = 37: Exit Function
'        Case "ny": AsignarFonemaBase_CA = 38: Exit Function
'        Case "ng": AsignarFonemaBase_CA = 39: Exit Function
'    End Select
'
'    ' -------------------------
'    ' LATERALS
'    ' -------------------------
'    If grafema = "l" Then AsignarFonemaBase_CA = 62: Exit Function
'    If grafema = "ll" Then AsignarFonemaBase_CA = 63: Exit Function
'    If grafema = "l·l" Then AsignarFonemaBase_CA = 63: Exit Function
'
'    ' -------------------------
'    ' VIBRANTS
'    ' -------------------------
'    If grafema = "rr" Then AsignarFonemaBase_CA = 60: Exit Function
'
'    If grafema = "r" Then
'        If ant = "" Or Not ant Like "[aeiouàèéíïòóúü]" Then
'            AsignarFonemaBase_CA = 60   ' inicial ? múltiple
'        Else
'            AsignarFonemaBase_CA = 59   ' intervocàlica ? simple
'        End If
'        Exit Function
'    End If
'
'    ' -------------------------
'    ' AFRICADES
'    ' -------------------------
'    If grafema = "tx" Then AsignarFonemaBase_CA = 57: Exit Function
'    If grafema = "tg" Or grafema = "tj" Then AsignarFonemaBase_CA = 58: Exit Function
'
'    ' -------------------------
'    ' FRICATIVES
'    ' -------------------------
'    If grafema = "s" Then AsignarFonemaBase_CA = 42: Exit Function
'    If grafema = "z" Then AsignarFonemaBase_CA = 43: Exit Function
'    If grafema = "x" Then AsignarFonemaBase_CA = 44: Exit Function
'    If grafema = "ix" Then AsignarFonemaBase_CA = 44: Exit Function
'    If grafema = "j" Then AsignarFonemaBase_CA = 45: Exit Function
'    If grafema = "ge" Or grafema = "gi" Then AsignarFonemaBase_CA = 45: Exit Function
'
'    ' -------------------------
'    ' OCLUSIVES
'    ' -------------------------
'    Select Case grafema
'        Case "p": AsignarFonemaBase_CA = 30: Exit Function
'        Case "b": AsignarFonemaBase_CA = 31: Exit Function
'        Case "t": AsignarFonemaBase_CA = 32: Exit Function
'        Case "d": AsignarFonemaBase_CA = 33: Exit Function
'        Case "c", "k", "qu": AsignarFonemaBase_CA = 34: Exit Function
'        Case "g", "gu": AsignarFonemaBase_CA = 35: Exit Function
'    End Select
'
'    ' -------------------------
'    ' Si no se reconoce
'    ' -------------------------
'    AsignarFonemaBase_CA = 255
'
'End Function
'
'Private Function NormalizarVocales_CA(ByVal texto As String) As String
'
'    texto = Replace(texto, "á", "a")
'    texto = Replace(texto, "à", "a")
'    texto = Replace(texto, "â", "a")
'
'    texto = Replace(texto, "é", "e")
'    texto = Replace(texto, "è", "è")   ' mantenim oberta
'    texto = Replace(texto, "ê", "e")
'
'    texto = Replace(texto, "í", "i")
'    texto = Replace(texto, "ï", "ï")
'
'    texto = Replace(texto, "ó", "o")
'    texto = Replace(texto, "ò", "ò")
'
'    texto = Replace(texto, "ú", "u")
'    texto = Replace(texto, "ü", "ü")
'
'    NormalizarVocales_CA = texto
'
'End Function


