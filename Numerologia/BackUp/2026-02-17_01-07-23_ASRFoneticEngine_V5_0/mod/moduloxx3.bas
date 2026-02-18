Attribute VB_Name = "moduloxx3"

'Option Compare Database
'Option Explicit
'
'' ============================================================
''   MOTOR FONOLÒGIC CATALÀ — COMPLET
''   Autor: Alba Salvà Ripoll + Copilot
''
''   Funcions principals:
''       · ConstruirCadenaFonemas_CA
''       · ProcesarSilaba_CA
''       · DetectarDigrafoCa
''       · AsignarFonemaBaseCa
''       · NormalizarVocales_CA
''       · HayLigaduraAutomatica_CA
''
''   Totalment compatible amb:
''       · DTO (clsDTO_Motor)
''       · qryFonemasBase
''       · qryFonemasValor
'' ============================================================
'
'
'' ============================================================
''   ENTRADA PRINCIPAL DEL MOTOR FONÈTIC CATALÀ
'' ============================================================
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
'    ' 0) Normalització de vocals i grafemes catalans
'    frase = NormalizarVocales_CA(ObjDTO.SilabasFinal)
'
'    ' 1) Separar síl·labes
'    arrSilabas = Split(frase, "|")
'
'    For i = 0 To UBound(arrSilabas)
'
'        silaba = Trim$(arrSilabas(i))
'
'        ' ---------------------------------------------------------
'        ' 2. Separador de paraula (síl·laba buida)
'        ' ---------------------------------------------------------
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
'        ' ---------------------------------------------------------
'        ' 3. Modificadors prosòdics
'        ' ---------------------------------------------------------
'        If InStr(silaba, "(") > 0 Then
'            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "80,"
'        ElseIf InStr(silaba, "[") > 0 Then
'            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "81,"
'        Else
'            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "82,"
'        End If
'
'        ' Netejar marcadors
'        silaba = Replace(Replace(Replace(Replace(silaba, "(", ""), ")", ""), "[", ""), "]", "")
'
'        ' ---------------------------------------------------------
'        ' 4. Processar grafemes catalans
'        ' ---------------------------------------------------------
'        ProcesarSilaba_CA silaba
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
'    ' 6. Neteja final
'    ' ---------------------------------------------------------
'    ObjDTO.IdsFonemas = Replace(ObjDTO.IdsFonemas, "|#", "#")
'    ObjDTO.IdsFonemas = Replace(ObjDTO.IdsFonemas, "#|", "#")
'
'    If Right$(ObjDTO.IdsFonemas, 1) Like "[|,#]" Then
'        ObjDTO.IdsFonemas = Left$(ObjDTO.IdsFonemas, Len(ObjDTO.IdsFonemas) - 1)
'    End If
'
'    ObjDTO.IdsFonemas = Replace(ObjDTO.IdsFonemas, ",|", "|")
'
'    ' ---------------------------------------------------------
'    ' 7. Reconstrucció IPA (mateix motor que ES)
'    ' ---------------------------------------------------------
'    ObjDTO.FonemasFinal = GenerarIPA()
'
'End Sub
'
'
'' ============================================================
''   PROCESSAR UNA SÍL·LABA CATALANA
'' ============================================================
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
'    ' 1. Ligadura manual
'    If CFG.ModoLigadura = 1 Then
'        ligaduraID = DetectarLigaduraManual(silaba)
'        If ligaduraID <> 0 Then
'            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & ligaduraID & ","
'        End If
'    End If
'
'    ' 2. Recórrer grafemes catalans
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
'        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & id & ","
'
'        i = i + Len(grafema)
'
'    Loop
'
'End Sub
'
'
'' ============================================================
''   DETECTAR DÍGRAFS CATALANS
'' ============================================================
'Function DetectarDigrafoCa(Texto As String, pos As Long) As String
'
'    Dim L As Long
'    Dim par As String, tri As String
'
'    L = Len(Texto)
'
'    If pos >= L Then
'        DetectarDigrafoCa = Mid$(Texto, pos, 1)
'        Exit Function
'    End If
'
'    ' Trígraf: l·l
'    If pos + 2 <= L Then
'        tri = Mid$(Texto, pos, 3)
'        If tri = "l·l" Then
'            DetectarDigrafoCa = tri
'            Exit Function
'        End If
'    End If
'
'    ' Dígrafs
'    par = Mid$(Texto, pos, 2)
'
'    Select Case par
'        Case "ny", "ll", "tx", "tg", "tj", "ix", "ss", "qu", "gu", "rr", "ch"
'            DetectarDigrafoCa = par
'            Exit Function
'    End Select
'
'    DetectarDigrafoCa = Mid$(Texto, pos, 1)
'
'End Function
'
'
'' ============================================================
''   ASSIGNACIÓ DE FONEMES CATALANS (IDs REALS)
'' ============================================================
'Function AsignarFonemaBaseCa(grafema As String, _
'                             Optional ByVal sig As String = "", _
'                             Optional ByVal ant As String = "") As Byte
'
'    grafema = LCase$(grafema)
'    sig = LCase$(sig)
'    ant = LCase$(ant)
'
'    ' VOCALS
'    Select Case grafema
'        Case "a", "á", "à", "â": AsignarFonemaBaseCa = 1: Exit Function
'        Case "e", "é", "ê": AsignarFonemaBaseCa = 2: Exit Function
'        Case "è": AsignarFonemaBaseCa = 3: Exit Function
'        Case "i", "í", "ï": AsignarFonemaBaseCa = 4: Exit Function
'        Case "o", "ó", "ô": AsignarFonemaBaseCa = 5: Exit Function
'        Case "ò": AsignarFonemaBaseCa = 6: Exit Function
'        Case "u", "ú", "ü": AsignarFonemaBaseCa = 7: Exit Function
'    End Select
'
'    ' SEMIVOCALS
'    If grafema = "i" And sig Like "[aeou]" Then AsignarFonemaBaseCa = 21: Exit Function
'    If grafema = "u" And sig Like "[aeio]" Then AsignarFonemaBaseCa = 22: Exit Function
'
'    ' NASALS
'    Select Case grafema
'        Case "m": AsignarFonemaBaseCa = 36: Exit Function
'        Case "n": AsignarFonemaBaseCa = 37: Exit Function
'        Case "ny": AsignarFonemaBaseCa = 38: Exit Function
'        Case "ng": AsignarFonemaBaseCa = 39: Exit Function
'    End Select
'
'    ' LATERALS
'    If grafema = "l" Then AsignarFonemaBaseCa = 62: Exit Function
'    If grafema = "ll" Or grafema = "l·l" Then AsignarFonemaBaseCa = 63: Exit Function
'
'    ' FRICATIVES
'    If grafema = "f" Then AsignarFonemaBaseCa = 40: Exit Function
'    If grafema = "v" Then AsignarFonemaBaseCa = 41: Exit Function
'
'    Select Case grafema
'        Case "s", "ss", "c", "ç": AsignarFonemaBaseCa = 42: Exit Function
'        Case "z": AsignarFonemaBaseCa = 43: Exit Function
'    End Select
'
'    If grafema = "x" Or grafema = "ix" Then AsignarFonemaBaseCa = 44: Exit Function
'    If grafema = "j" Then AsignarFonemaBaseCa = 45: Exit Function
'    If grafema = "g" And (sig = "e" Or sig = "i") Then AsignarFon
'
