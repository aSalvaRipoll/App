Attribute VB_Name = "Módulo1"

'Option Compare Database
'Option Explicit
'
'' ============================================
'' PUNTO DE ENTRADA AL PROCESADOR FONÉTICO (CATALÁN)
'' ============================================
'Public Sub ConstruirCadenaFonemas_CA()
'
'    Dim arrSilabas As Variant
'    Dim i As Long
'    Dim silaba As Variant
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
'    ' 0) Normalizar vocales y grafemas catalanes
'    frase = NormalizarVocales_CA(ObjDTO.SilabasFinal)
'
'    ' 1) Separar sílabas por "|"
'    arrSilabas = Split(frase, "|")
'
'    For i = 0 To UBound(arrSilabas)
'
'        silaba = Trim$(arrSilabas(i))
'
'        ' 2) Separador de palabra (sílaba vacía)
'        If silaba = "" Then
'
'            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "#"
'
'            If CFG.ModoLigadura = 2 Then
'                ultimaSilaba = BuscarSilabaRealAnterior(arrSilabas, i)
'                siguienteSilaba = BuscarSilabaRealPosterior(arrSilabas, i)
'
'                If HayLigaduraAutomatica_CA(ultimaSilaba, siguienteSilaba) Then
'                    ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "84,"
'                End If
'            End If
'
'            GoTo SiguienteIteracion
'        End If
'
'        ' 3) Modificadores prosódicos (igual que ES)
'        If InStr(silaba, "(") > 0 Then
'            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "80,"
'        ElseIf InStr(silaba, "[") > 0 Then
'            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "81,"
'        Else
'            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "82,"
'        End If
'
'        silaba = Replace(Replace(Replace(Replace(silaba, "(", ""), ")", ""), "[", ""), "]", "")
'
'        ' 4) Procesar grafemas catalanes
'        ProcesarSilaba_CA silaba
'
'        ' 5) Separador silábico
'        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "|"
'
'SiguienteIteracion:
'    Next i
'
'    ' 6) Limpieza final (igual que ES)
'    ObjDTO.IdsFonemas = Replace(ObjDTO.IdsFonemas, "|#", "#")
'    ObjDTO.IdsFonemas = Replace(ObjDTO.IdsFonemas, "#|", "#")
'
'    If Right$(ObjDTO.IdsFonemas, 1) = "|" Or Right$(ObjDTO.IdsFonemas, 1) = "#" Or Right$(ObjDTO.IdsFonemas, 1) = "," Then
'        ObjDTO.IdsFonemas = Left$(ObjDTO.IdsFonemas, Len(ObjDTO.IdsFonemas) - 1)
'    End If
'
'    ObjDTO.IdsFonemas = Replace(ObjDTO.IdsFonemas, ",|", "|")
'
'    ' 7) Reconstruir IPA (puedes reutilizar tu bucle ES cambiando solo NormalizarVocales)
'    '    o llamar directamente a GenerarIPA si usas el mismo formato de IdsFonemas
'    ObjDTO.FonemasFinal = GenerarIPA   ' si el formato de IdsFonemas es el mismo
'
'End Sub
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
'    ' 1) Ligadura manual (igual que ES)
'    If CFG.ModoLigadura = 1 Then
'        ligaduraID = DetectarLigaduraManual(silaba)
'        If ligaduraID <> 0 Then
'            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & CStr(ligaduraID) & ","
'        End If
'    End If
'
'    ' 2) Recorrido de grafemas catalanes
'    i = 1
'    Do While i <= Len(silaba)
'
'        grafema = DetectarDigrafoCa(silaba, i)
'
'        If i > 1 Then
'            antCh = Mid$(silaba, i - 1, 1)
'        Else
'            antCh = ""
'        End If
'
'        If i + Len(grafema) <= Len(silaba) Then
'            sigCh = Mid$(silaba, i + Len(grafema), 1)
'        Else
'            sigCh = ""
'        End If
'
'        id = AsignarFonemaBaseCa(grafema, sigCh, antCh)
'
'        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & CStr(id) & ","
'
'        i = i + Len(grafema)
'
'    Loop
'
'End Sub
'
'Function AsignarFonemaBaseCa(grafema As String, _
'                             Optional ByVal sig As String = "", _
'                             Optional ByVal ant As String = "") As Byte
'
'    grafema = LCase$(grafema)
'    sig = LCase$(sig)
'    ant = LCase$(ant)
'
'    ' ============================================
'    ' VOCALS CATALANES
'    ' ============================================
'    Select Case grafema
'        Case "a", "á", "à", "â": AsignarFonemaBaseCa = 1: Exit Function
'        Case "e", "é", "ê": AsignarFonemaBaseCa = 2: Exit Function
'        Case "è": AsignarFonemaBaseCa = 3: Exit Function
'        Case "i", "í", "ï": AsignarFonemaBaseCa = 4: Exit Function
'        Case "o", "ó", "ô": AsignarFonemaBaseCa = 5: Exit Function
'        Case "ò": AsignarFonemaBaseCa = 6: Exit Function
'        Case "u", "ú", "ü": AsignarFonemaBaseCa = 7: Exit Function
'
'        ' Vocal neutra (schwa)
'        Case "a", "e"
'            ' En catalán central, la vocal neutra depende de la sílaba átona
'            ' Aquí solo la asignamos si viene marcada por el silabeador
'            If ObjDTO.EsSilabaAtona Then
'                AsignarFonemaBaseCa = 8
'                Exit Function
'            End If
'    End Select
'
'    ' ============================================
'    ' SEMIVOCALS
'    ' ============================================
'    If grafema = "i" And sig Like "[aeou]" Then AsignarFonemaBaseCa = 21: Exit Function
'    If grafema = "u" And sig Like "[aeio]" Then AsignarFonemaBaseCa = 22: Exit Function
'
'    ' ============================================
'    ' NASALS
'    ' ============================================
'    Select Case grafema
'        Case "m": AsignarFonemaBaseCa = 36: Exit Function
'        Case "n": AsignarFonemaBaseCa = 37: Exit Function
'        Case "ny": AsignarFonemaBaseCa = 38: Exit Function
'        Case "ng": AsignarFonemaBaseCa = 39: Exit Function
'    End Select
'
'    ' ============================================
'    ' LATERALS
'    ' ============================================
'    If grafema = "l" Then AsignarFonemaBaseCa = 62: Exit Function
'    If grafema = "ll" Or grafema = "l·l" Then AsignarFonemaBaseCa = 63: Exit Function
'
'    ' ============================================
'    ' FRICATIVES
'    ' ============================================
'    If grafema = "f" Then AsignarFonemaBaseCa = 40: Exit Function
'    If grafema = "v" Then AsignarFonemaBaseCa = 41: Exit Function
'
'    ' Sibilants
'    Select Case grafema
'        Case "s", "ss", "c", "ç": AsignarFonemaBaseCa = 42: Exit Function
'        Case "z": AsignarFonemaBaseCa = 43: Exit Function
'    End Select
'
'    ' Palatals
'    If grafema = "x" Then AsignarFonemaBaseCa = 44: Exit Function
'    If grafema = "ix" Then AsignarFonemaBaseCa = 44: Exit Function
'
'    If grafema = "j" Then AsignarFonemaBaseCa = 45: Exit Function
'    If grafema = "g" And (sig = "e" Or sig = "i") Then AsignarFonemaBaseCa = 45: Exit Function
'
'    ' ============================================
'    ' AFRICADES
'    ' ============================================
'    If grafema = "tx" Then AsignarFonemaBaseCa = 57: Exit Function
'    If grafema = "tg" Or grafema = "tj" Then AsignarFonemaBaseCa = 58: Exit Function
'
'    ' ============================================
'    ' OCLUSIVES
'    ' ============================================
'    Select Case grafema
'        Case "p": AsignarFonemaBaseCa = 30: Exit Function
'        Case "b": AsignarFonemaBaseCa = 31: Exit Function
'        Case "t": AsignarFonemaBaseCa = 32: Exit Function
'        Case "d": AsignarFonemaBaseCa = 33: Exit Function
'        Case "k", "c", "qu": AsignarFonemaBaseCa = 34: Exit Function
'        Case "g": AsignarFonemaBaseCa = 35: Exit Function
'    End Select
'
'    ' ============================================
'    ' VIBRANTS
'    ' ============================================
'    If grafema = "rr" Then AsignarFonemaBaseCa = 60: Exit Function
'
'    If grafema = "r" Then
'        If ant = "" Or ant Like "[bcdfghjklmnpqrstvwxyz]" Then
'            AsignarFonemaBaseCa = 60   ' r múltiple en inicio
'        Else
'            AsignarFonemaBaseCa = 59   ' r simple
'        End If
'        Exit Function
'    End If
'
'    ' ============================================
'    ' H (aspirada en préstamos)
'    ' ============================================
'    If grafema = "h" Then AsignarFonemaBaseCa = 49: Exit Function
'
'    ' ============================================
'    ' DESCONOCIDO
'    ' ============================================
'    AsignarFonemaBaseCa = 255
'
'End Function
'
'Function DetectarDigrafoCa(Texto As String, pos As Long) As String
'
'    Dim L As Long
'    Dim par As String, tri As String, sig As String
'
'    L = Len(Texto)
'
'    ' 1) Si queda solo una letra ? devolverla
'    If pos >= L Then
'        DetectarDigrafoCa = Mid$(Texto, pos, 1)
'        Exit Function
'    End If
'
'    ' 2) Trígrafo posible: l·l
'    If pos + 2 <= L Then
'        tri = Mid$(Texto, pos, 3)
'        If tri = "l·l" Then
'            DetectarDigrafoCa = tri
'            Exit Function
'        End If
'    End If
'
'    ' 3) Dígrafos de dos letras
'    par = Mid$(Texto, pos, 2)
'
'    Select Case par
'
'        ' Nasal palatal
'        Case "ny"
'            DetectarDigrafoCa = "ny"
'            Exit Function
'
'        ' Lateral palatal
'        Case "ll"
'            DetectarDigrafoCa = "ll"
'            Exit Function
'
'        ' Africada sorda
'        Case "tx"
'            DetectarDigrafoCa = "tx"
'            Exit Function
'
'        ' Africada sonora
'        Case "tg", "tj"
'            DetectarDigrafoCa = par
'            Exit Function
'
'        ' Fricativa palatal sorda
'        Case "ix"
'            DetectarDigrafoCa = "ix"
'            Exit Function
'
'        ' S sorda geminada
'        Case "ss"
'            DetectarDigrafoCa = "ss"
'            Exit Function
'
'        ' Qu ? /k/
'        Case "qu"
'            DetectarDigrafoCa = "qu"
'            Exit Function
'
'        ' Gu ? /g/
'        Case "gu"
'            DetectarDigrafoCa = "gu"
'            Exit Function
'
'        ' R múltiple
'        Case "rr"
'            DetectarDigrafoCa = "rr"
'            Exit Function
'
'        ' Préstamos
'        Case "ch"
'            DetectarDigrafoCa = "ch"
'            Exit Function
'
'    End Select
'
'    ' 4) Grafema simple
'    DetectarDigrafoCa = Mid$(Texto, pos, 1)
'
'End Function
'
'Function NormalizarVocales_CA(ByVal Texto As String) As String
'
'    Texto = LCase$(Texto)
'
'    ' Mantener acentos graves y diéresis
'    ' Solo normalizamos variantes raras
'
'    Texto = Replace(Texto, "á", "a")
'    Texto = Replace(Texto, "â", "a")
'
'    Texto = Replace(Texto, "é", "e")
'    Texto = Replace(Texto, "ê", "e")
'
'    Texto = Replace(Texto, "ó", "o")
'    Texto = Replace(Texto, "ô", "o")
'
'    Texto = Replace(Texto, "ú", "u")
'    Texto = Replace(Texto, "û", "u")
'
'    ' Mantener: à è ò í ï ü ç l·l
'
'    ' Eliminar h muda (excepto en "hò", "hí", préstamos)
'    Texto = Replace(Texto, "h", "")
'
'    ' Mantener apóstrofos
'    Texto = Replace(Texto, "’", "'")
'
'    NormalizarVocales_CA = Texto
'
'End Function
'
'
