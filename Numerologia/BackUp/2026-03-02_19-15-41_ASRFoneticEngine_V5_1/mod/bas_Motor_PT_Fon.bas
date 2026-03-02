Attribute VB_Name = "bas_Motor_PT_Fon"

Option Compare Database
Option Explicit

' ============================================================
'   MOTOR FONÉTICO — PORTUGUÉS EUROPEO (PT-EU)
'   MISMA ARQUITECTURA QUE ES / CA / GL
' ============================================================

Private esFinalDePalabra As Boolean
Private SigSilaba As String
Private AntSilaba As String

'' ============================================================
''   MOTOR FONÉTICO — PORTUGUÉS EUROPEO (PT-EU)
''   MÓDULO COMPLETO — PARTE 1 / 3
'' ============================================================

Public Sub ConstruirCadenaFonemas_PT()

    Dim arrSilabas As Variant
    Dim i As Long
    Dim silabaCruda As String
    Dim ultimaSilaba As String
    Dim SiguienteSilaba As String
    Dim ligaduraID As String
    Dim arrFon As Variant
    Dim Fon As Variant
    Dim res As String

    ObjDTO.IdsFonemas = ""
    ObjDTO.FonemasFinal = ""

    ' 0. Normalización mínima (PT no necesita cambios)
    arrSilabas = Split(ObjDTO.SilabasFinal, "|")

    For i = 0 To UBound(arrSilabas)

        silabaCruda = arrSilabas(i)

        esFinalDePalabra = False
        If i = UBound(arrSilabas) Then
            esFinalDePalabra = True
        ElseIf i < UBound(arrSilabas) And Trim(arrSilabas(i + 1)) = "" Then
            esFinalDePalabra = True
        End If
        
        AntSilaba = BuscarSilabaRealAnterior(arrSilabas, i)
        SigSilaba = BuscarSilabaRealPosterior(arrSilabas, i)
        
        ' ---------------------------------------------------------
        ' 1. Separador de palabra (sílaba vacía)
        ' ---------------------------------------------------------
        If Trim$(silabaCruda) = "" Then
            
            
            SigSilaba = BuscarSilabaRealPosterior(arrSilabas, i)
            
            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "#"

            ultimaSilaba = AntSilaba 'BuscarSilabaRealAnterior(arrSilabas, i)
            SiguienteSilaba = SigSilaba 'BuscarSilabaRealPosterior(arrSilabas, i)

            If HayLigaduraAutomatica(ultimaSilaba, SiguienteSilaba) Then
                ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "84,"
            End If

            GoTo SiguienteIteracion
        End If

        ' ---------------------------------------------------------
        ' 2. Ligadura manual
        ' ---------------------------------------------------------
        ligaduraID = DetectarLigaduraManual(silabaCruda)
        If ligaduraID <> 0 Then
            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & ligaduraID & ","
        End If

        ' ---------------------------------------------------------
        ' 3. Acento
        ' ---------------------------------------------------------
        If InStr(silabaCruda, "(") > 0 Then
            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "80,"
        ElseIf InStr(silabaCruda, "[") > 0 Then
            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "81,"
        Else
            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "82,"
        End If

        ' ---------------------------------------------------------
        ' 4. Procesar grafemas PT-EU
        ' ---------------------------------------------------------
        'SiguienteSilaba = BuscarSilabaRealPosterior(arrSilabas, i)
        ProcesarSilaba_PT silabaCruda ', SiguienteSilaba

        ' ---------------------------------------------------------
        ' 5. Separador silábico
        ' ---------------------------------------------------------
        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "|"

SiguienteIteracion:
    Next i

    ' ---------------------------------------------------------
    ' 6. Limpieza final (idéntica a ES/CA/GL)
    ' ---------------------------------------------------------
    ObjDTO.IdsFonemas = Replace(ObjDTO.IdsFonemas, "|#", "#")
    ObjDTO.IdsFonemas = Replace(ObjDTO.IdsFonemas, "#|", "#")
    ObjDTO.IdsFonemas = Replace(ObjDTO.IdsFonemas, ",|", "|")
    ObjDTO.IdsFonemas = Replace(ObjDTO.IdsFonemas, "84,84,", "84,")

    While Right$(ObjDTO.IdsFonemas, 1) = "|" Or _
          Right$(ObjDTO.IdsFonemas, 1) = "#" Or _
          Right$(ObjDTO.IdsFonemas, 1) = ","
        ObjDTO.IdsFonemas = Left$(ObjDTO.IdsFonemas, Len(ObjDTO.IdsFonemas) - 1)
    Wend

    ' ---------------------------------------------------------
    ' 7. Construcción IPA final (idéntica a ES/CA/GL)
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

'Public Sub ConstruirCadenaFonemas_PT()
'
'    Dim arrSilabas As Variant
'    Dim i As Long
'    Dim silabaCruda As String
'    Dim ultimaSilaba As String
'    Dim siguienteSilaba As String
'    Dim ligaduraID As String
'
'    ObjDTO.IdsFonemas = ""
'    ObjDTO.FonemasFinal = ""
'
'    arrSilabas = Split(ObjDTO.SilabasFinal, "|")
'
'    For i = 0 To UBound(arrSilabas)
'
'        silabaCruda = Trim$(arrSilabas(i))
'
'        ' 1. Separador de palabra
'        If silabaCruda = "" Then
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
'        ' 2. Ligadura manual
'        ligaduraID = DetectarLigaduraManual(silabaCruda)
'        If ligaduraID <> 0 Then
'            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & ligaduraID & ","
'        End If
'
'        ' 3. Acento
'        If InStr(silabaCruda, "(") > 0 Then
'            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "80,"
'        ElseIf InStr(silabaCruda, "[") > 0 Then
'            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "81,"
'        Else
'            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "82,"
'        End If
'
'        silabaCruda = Replace(silabaCruda, "(", "")
'        silabaCruda = Replace(silabaCruda, ")", "")
'        silabaCruda = Replace(silabaCruda, "[", "")
'        silabaCruda = Replace(silabaCruda, "]", "")
'
'        ' 4. Procesar grafemas
'        ProcesarSilaba_PT silabaCruda
'
'        ' 5. Separador silábico
'        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "|"
'
'SiguienteIteracion:
'    Next i
'
'    ' 6. Limpieza final
'    ObjDTO.IdsFonemas = Replace(ObjDTO.IdsFonemas, "|#", "#")
'    ObjDTO.IdsFonemas = Replace(ObjDTO.IdsFonemas, "#|", "#")
'    ObjDTO.IdsFonemas = Replace(ObjDTO.IdsFonemas, ",|", "|")
'    ObjDTO.IdsFonemas = Replace(ObjDTO.IdsFonemas, "84,84,", "84,")
'
'    While Right$(ObjDTO.IdsFonemas, 1) Like "[|#,]"
'        ObjDTO.IdsFonemas = Left$(ObjDTO.IdsFonemas, Len(ObjDTO.IdsFonemas) - 1)
'    Wend
'
'    ' 7. Generar IPA
'    ObjDTO.FonemasFinal = GenerarIPA()
'
'End Sub

' ============================================================
'   PROCESAR SÍLABA PT-EU (versión final corregida)
' ============================================================

Private Sub ProcesarSilaba_PT(ByVal silabaCruda As String) ', ByVal esAtona As Boolean)

    Dim silaba As String
    Dim i As Long
    Dim grafema As String
    Dim id As Byte
    Dim antCh As String
    Dim sigCh As String
    Dim esAtona As Boolean
    
    silaba = Trim$(silabaCruda)

    esAtona = True
    
    If InStr(silabaCruda, "(") > 0 Then esAtona = False
    If InStr(silabaCruda, "[") > 0 Then esAtona = False


    ' ------------------------------------------------------------
    ' 1. LIMPIEZA COMPLETA DE LA SÍLABA
    '    (igual que en ES/CA/GL)
    ' ------------------------------------------------------------
    'silaba = silabaCruda
    silaba = Replace(silaba, "(", "")
    silaba = Replace(silaba, ")", "")
    silaba = Replace(silaba, "[", "")
    silaba = Replace(silaba, "]", "")
    'silaba = Replace(silaba, "_", "")
    silaba = Trim$(silaba)

    ' ------------------------------------------------------------
    ' 2. RECORRIDO DE GRAFEMAS
    ' ------------------------------------------------------------
    i = 1
    Do While i <= Len(silaba)

        ' --------------------------------------------------------
        ' 2.1 TRÍGRAFOS NASALES (deben ir ANTES de todo)
        ' --------------------------------------------------------
        'If i + 2 <= Len(silaba) Then
        If i + 1 <= Len(silaba) Then
            'If ProcesarDiptongoNasal_PT(Mid$(silaba, i, 3)) Then
            If ProcesarDiptongoNasal_PT(Mid$(silaba, i, 2)) Then
                'i = i + 3
                i = i + 2
                GoTo SiguienteGrafema
            End If
        End If

        ' --------------------------------------------------------
        ' 2.2 DETECTAR GRAFEMA (2 letras)
        ' --------------------------------------------------------
        grafema = DetectarDigrafo_PT(silaba, i)

        ' Caracteres anterior y siguiente
        If i > 1 Then antCh = Mid$(silaba, i - 1, 1) Else antCh = ""
        If i + Len(grafema) <= Len(silaba) Then
            sigCh = Mid$(silaba, i + Len(grafema), 1)
        Else
            sigCh = ""
        End If


        If esFinalDePalabra Then
            Select Case grafema
                Case "am" ', "an", "em", "en", "im", "in", "om", "on", "um", "un"
                    ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "25,24,"
                    i = i + 2
                    GoTo SiguienteGrafema
                Case "an"
                    ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "25,"
                    i = i + 2
                    GoTo SiguienteGrafema
                Case "em", "en"
                    ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "15,"
                    i = i + 2
                    GoTo SiguienteGrafema
                Case "im", "in"
                    ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "16,"
                    i = i + 2
                    GoTo SiguienteGrafema
                Case "om", "on"
                    ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "17,"
                    i = i + 2
                    GoTo SiguienteGrafema
                Case "um", "un"
                    ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "18,"
                    i = i + 2
                    GoTo SiguienteGrafema
                Case Else
                
            End Select
        End If

        ' --------------------------------------------------------
        ' 2.3 REDUCCIÓN VOCÁLICA PT-EU
        ' --------------------------------------------------------
        If ProcesarReduccionVocal_PT(grafema, esAtona) Then
            i = i + Len(grafema)
            GoTo SiguienteGrafema
        End If

        ' --------------------------------------------------------
        ' 2.4 PROCESAR X
        ' --------------------------------------------------------
        If grafema = "x" Then
            If ProcesarX_PT(antCh, sigCh) Then
                i = i + 1
                GoTo SiguienteGrafema
            End If
        End If

' --------------------------------------------------------
' 2.5 PROCESAR R
' --------------------------------------------------------
If grafema = "r" Or grafema = "rr" Then
    id = ProcesarR_PT(silaba, grafema, antCh, sigCh)
    addLog "id: " & id

    If id <> 0 Then
        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & id & ","
    End If

    i = i + Len(grafema)
    GoTo SiguienteGrafema
End If

        ' --------------------------------------------------------
        ' 2.6 PROCESAR S
        ' --------------------------------------------------------
        If grafema = "s" Then
            'id = ProcesarS_PT(silaba, SigSilaba, antCh, grafema)
            id = ProcesarS_PT(silaba, grafema, antCh, sigCh)
            addLog "id: " & id
            
            If id <> 0 Then
                ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & id & ","
            End If

'            If ProcesarS_PT(antCh, sigCh, silaba) Then
                i = i + 1
                GoTo SiguienteGrafema
'            End If
        End If

        ' --------------------------------------------------------
        ' 2.7 FONEMA BASE
        ' --------------------------------------------------------
        id = AsignarFonemaBase_PT(grafema, sigCh, antCh)
        If id <> 0 Then
            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & id & ","
        End If

        i = i + Len(grafema)

SiguienteGrafema:
    Loop

End Sub

'' ============================================================
''   PROCESAR SÍLABA PT-EU
'' ============================================================
'
'Private Sub ProcesarSilaba_PT(ByVal silabaCruda As String)
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
'    silaba = Trim$(silabaCruda)
'
'    ' Normalización mínima (NO eliminar h)
'    silaba = silaba
'
'    ' Detectar si la sílaba es átona (ID 82 ya añadido)
''    esAtona = (Right$(ObjDTO.IdsFonemas, 3) = "82,")
'
'    If InStr(silabaCruda, "(") > 0 Then esAtona = False
'    If InStr(silabaCruda, "[") > 0 Then esAtona = False
'
'    ' Netejar marcadors per treballar amb la grafia neta
'    silaba = silabaCruda
'    silaba = Replace(silaba, "(", "")
'    silaba = Replace(silaba, ")", "")
'    silaba = Replace(silaba, "[", "")
'    silaba = Replace(silaba, "]", "")
'    silaba = Trim$(silaba)
'
'
'    ' Ligadura manual
'    ligaduraID = DetectarLigaduraManual(silaba)
'    If ligaduraID <> 0 Then
'        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & ligaduraID & ","
'    End If
'
'    ' 2. Processar grafemas
'    i = 1
'    Do While i <= Len(silaba)
'
'        ' 0. Trigramas nasales
'        If i + 2 <= Len(silaba) Then
'            If ProcesarDiptongoNasal_PT(Mid$(silaba, i, 3)) Then
'                i = i + 3
'                GoTo SiguienteGrafema
'            End If
'        End If
'
'        grafema = DetectarDigrafo_PT(silaba, i)
'
'        If i > 1 Then antCh = Mid$(silaba, i - 1, 1) Else antCh = ""
'        If i + Len(grafema) <= Len(silaba) Then
'            sigCh = Mid$(silaba, i + Len(grafema), 1)
'        Else
'            sigCh = ""
'        End If
'
'        ' 1. Reducción vocálica
'        If ProcesarReduccionVocal_PT(grafema, esAtona) Then
'            i = i + Len(grafema)
'            GoTo SiguienteGrafema
'        End If
'
'        ' 2. Procesar X
'        If grafema = "x" Then
'            If ProcesarX_PT(antCh, sigCh) Then
'                i = i + 1
'                GoTo SiguienteGrafema
'            End If
'        End If
'
'        ' 3. Procesar S
'        If grafema = "s" Then
'            If ProcesarS_PT(antCh, sigCh) Then
'                i = i + 1
'                GoTo SiguienteGrafema
'            End If
'        End If
'
'        ' 4. Fonema base
'        id = AsignarFonemaBase_PT(grafema, sigCh, antCh)
'        If id <> 0 Then ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & id & ","
'
'        i = i + Len(grafema)
'
'SiguienteGrafema:
'    Loop
'
'End Sub

' ============================================================
'   LIGADURA AUTOMÁTICA PT-EU
' ============================================================
Private Function HayLigaduraAutomatica(ultimaSilaba As String, SiguienteSilaba As String) As Boolean

    Dim fin As String
    Dim ini As String

    If ultimaSilaba = "" Or SiguienteSilaba = "" Then
        HayLigaduraAutomatica = False
        Exit Function
    End If

    fin = Right$(LimpiaSilaba(ultimaSilaba), 1)
    ini = Left$(LimpiaSilaba(SiguienteSilaba), 1)

    ' Vocales y semivocales portuguesas
    Dim vocales As String
    vocales = "aeiouáéíóúâêôãõ"

    ' Condición PT-EU:
    ' vocal/semivocal final + vocal/semivocal inicial ? ligadura
    If InStr(vocales, fin) > 0 And InStr(vocales, ini) > 0 Then
        HayLigaduraAutomatica = True
    Else
        HayLigaduraAutomatica = False
    End If

End Function
Private Function LimpiaSilaba(ByVal strCad As String) As String
    
    strCad = Replace(strCad, "(", "")
    strCad = Replace(strCad, ")", "")
    strCad = Replace(strCad, "[", "")
    strCad = Replace(strCad, "]", "")
    LimpiaSilaba = Trim$(strCad)
    
End Function


' ============================================================
'   LIGADURA MANUAL (símbolo "_")
' ============================================================

Private Function DetectarLigaduraManual(ByVal silaba As String) As Byte

    ' El usuario puede escribir:  a_mi  ? a?mi
    ' El símbolo "_" NO debe aparecer en la salida,
    ' solo sirve para insertar el fonema ID 84 (?)

    If InStr(silaba, "_") > 0 Then
        DetectarLigaduraManual = 84   ' ID de ligadura manual
    Else
        DetectarLigaduraManual = 0
    End If

End Function

Private Function BuscarSilabaRealAnterior(arr As Variant, pos As Long) As String
    Dim j As Long
    If pos > LBound(arr) Then
    For j = pos - 1 To LBound(arr) Step -1
        ' Si encuentro frontera de palabra, paro
        If Trim$(arr(j)) = "" Then Exit For
        
        ' Si encuentro sílaba real, la devuelvo
        BuscarSilabaRealAnterior = Trim$(arr(j))
        Exit Function
    Next j
    End If
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

' ============================================================
'   DETECTAR GRAFEMAS PT-EU
' ============================================================
Private Function DetectarDigrafo_PT(texto As String, pos As Long) As String

    Dim L As Long
    Dim g3 As String, g2 As String

    L = Len(texto)

    ' Trigramas
    If pos + 2 <= L Then
        g3 = Mid$(texto, pos, 3)
        Select Case g3
            Case "ão", "õe", "ãe"
                DetectarDigrafo_PT = g3
                Exit Function
        End Select
    End If

    ' Dígrafos
    If pos + 1 <= L Then
        g2 = Mid$(texto, pos, 2)

        Select Case g2
            Case "nh", "lh", "ch", "rr", "ss", "gu", "qu"
                DetectarDigrafo_PT = g2
                Exit Function
        End Select

        ' Nasales finales
        Select Case g2
            Case "am", "an", "em", "en", "im", "in", "om", "on", "um", "un"
                If pos + 2 > L Or Not (Mid$(texto, pos + 2, 1) Like "[aeiouáéíóúâêôãõ]") Then
                    DetectarDigrafo_PT = g2
                    Exit Function
                End If
        End Select
    End If

    DetectarDigrafo_PT = Mid$(texto, pos, 1)

End Function

' ============================================================
'   FONEMAS BASE PT-EU
' ============================================================

Private Function AsignarFonemaBase_PT(grafema As String, _
                             Optional ByVal sig As String = "", _
                             Optional ByVal ant As String = "") As Byte

    ' ============================================================
    ' SEMIVOCALES /w/ y /j/
    ' ============================================================

    ' /w/ ? u + vocal fuerte
    If grafema = "u" And (sig Like "[aeoáéóâêôãõ]") Then
        AsignarFonemaBase_PT = 22
        Exit Function
    End If

    ' /j/ ? i + vocal fuerte
    If grafema = "i" And (sig Like "[aeouáéóâêôãõ]") Then
        AsignarFonemaBase_PT = 21
        Exit Function
    End If

    ' ============================================================
    ' VOCALES ORALES
    ' ============================================================

    Select Case grafema
        Case "a", "á", "â": AsignarFonemaBase_PT = 1: Exit Function
        Case "e", "é", "ê": AsignarFonemaBase_PT = 2: Exit Function
        Case "i", "í":      AsignarFonemaBase_PT = 4: Exit Function
        Case "o", "ó", "ô": AsignarFonemaBase_PT = 5: Exit Function
        Case "u", "ú":      AsignarFonemaBase_PT = 7: Exit Function
    End Select

    ' ============================================================
    ' VOCALES NASALES (IDs 14–18)
    ' ============================================================

    Select Case grafema
        Case "ã", "am", "an": AsignarFonemaBase_PT = 25: Exit Function ' AsignarFonemaBase_PT = 14
        Case "em", "en":      AsignarFonemaBase_PT = 15: Exit Function
        Case "im", "in":      AsignarFonemaBase_PT = 16: Exit Function
        Case "õ", "om", "on": AsignarFonemaBase_PT = 17: Exit Function
        Case "um", "un":      AsignarFonemaBase_PT = 18: Exit Function
    End Select

    ' ============================================================
    ' OCLUSIVAS
    ' ============================================================

    Select Case grafema
        Case "p": AsignarFonemaBase_PT = 30: Exit Function
        Case "b": AsignarFonemaBase_PT = 31: Exit Function
        Case "t": AsignarFonemaBase_PT = 32: Exit Function
        Case "d": AsignarFonemaBase_PT = 33: Exit Function
    End Select

    ' k / c / q
    If grafema = "k" Then AsignarFonemaBase_PT = 34: Exit Function

    If grafema = "c" Then
        If sig Like "[eiéê]" Then
            AsignarFonemaBase_PT = 42   ' /s/
        Else
            AsignarFonemaBase_PT = 34   ' /k/
        End If
        Exit Function
    End If

    If grafema = "qu" Then
        AsignarFonemaBase_PT = 34
        Exit Function
    End If

    If grafema = "g" Then
        If sig Like "[eiéê]" Then
            AsignarFonemaBase_PT = 45   ' /?/
        Else
            AsignarFonemaBase_PT = 35   ' /g/
        End If
        Exit Function
    End If

    If grafema = "gu" Then
        AsignarFonemaBase_PT = 35
        Exit Function
    End If

    ' ============================================================
    ' FRICATIVAS
    ' ============================================================

    If grafema = "f" Then AsignarFonemaBase_PT = 40: Exit Function
    If grafema = "v" Then AsignarFonemaBase_PT = 41: Exit Function
    'If grafema = "ss" Then AsignarFonemaBase_PT = 42: Exit Function   ' /s/
    'If grafema = "s" Then AsignarFonemaBase_PT = 42: Exit Function
    If grafema = "z" Then AsignarFonemaBase_PT = 43: Exit Function
    'If grafema = "x" Then AsignarFonemaBase_PT = 44: Exit Function   ' /?/ por defecto

    ' ============================================================
    ' NASALES
    ' ============================================================

    If grafema = "nh" Then AsignarFonemaBase_PT = 38: Exit Function
    If grafema = "m" Then AsignarFonemaBase_PT = 36: Exit Function
    If grafema = "n" Then AsignarFonemaBase_PT = 37: Exit Function

    ' ============================================================
    ' LATERALES
    ' ============================================================

    If grafema = "lh" Then AsignarFonemaBase_PT = 63: Exit Function
    If grafema = "l" Then AsignarFonemaBase_PT = 62: Exit Function

    ' ============================================================
    ' VIBRANTES PORTUGUESAS
    ' ============================================================

    ' rr ? /?/
    If grafema = "rr" Then
        AsignarFonemaBase_PT = 51
        Exit Function
    End If

    ' r inicial o postconsonántica ? /?/
    If grafema = "r" Then
        If ant = "" Or Not (ant Like "[aeiouáéíóúâêôãõ]") Then
            AsignarFonemaBase_PT = 51
        Else
            AsignarFonemaBase_PT = 59   ' intervocálica ? /?/
        End If
        Exit Function
    End If

    ' ============================================================
    ' AFRICADA
    ' ============================================================

    If grafema = "ch" Then
        AsignarFonemaBase_PT = 57
        Exit Function
    End If

    ' ============================================================
    ' H MUDA
    ' ============================================================

    If grafema = "h" Then
        AsignarFonemaBase_PT = 0
        Exit Function
    End If

    ' ============================================================
    ' DESCONOCIDO
    ' ============================================================

    AsignarFonemaBase_PT = 255

End Function

' ============================================================
'   TRÍGRAFOS NASALES PT-EU
' ============================================================

Private Function ProcesarDiptongoNasal_PT(grafema As String) As Boolean

    Select Case grafema

        ' ão ? /?~/ + /w~/ ? 14 + 24
        Case "ão"
            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "25,24,"
            'ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "14,24,"
            ProcesarDiptongoNasal_PT = True
            Exit Function

        ' õe ? /õ/ + /j~/ ? 17 + 23
        Case "õe"
            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "17,23,"
            ProcesarDiptongoNasal_PT = True
            Exit Function

        ' ãe ? /?~/ + /j~/ ? 14 + 23
        Case "ãe"
            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "25,23,"
            'ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "14,23,"
            ProcesarDiptongoNasal_PT = True
            Exit Function

    End Select

    ProcesarDiptongoNasal_PT = False

End Function


' ============================================================
'   PROCESAR X PORTUGUÉS (5 valores)
' ============================================================

Private Function ProcesarX_PT(ant As String, sig As String) As Boolean

    ' ex- + vocal ? /z/
    If ant = "e" And (sig Like "[aeiouáéíóúâêôãõ]") Then
        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "43,"
        ProcesarX_PT = True
        Exit Function
    End If

    ' intervocálica ? /z/
    If (ant Like "[aeiouáéíóúâêôãõ]") And (sig Like "[aeiouáéíóúâêôãõ]") Then
        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "43,"
        ProcesarX_PT = True
        Exit Function
    End If

    ' antes de consonante sonora ? /z/
    If sig Like "[bdgvzmnrl]" Then
        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "43,"
        ProcesarX_PT = True
        Exit Function
    End If

    ' antes de consonante sorda ? /?/
    If sig Like "[ptkfscx]" Then
        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "44,"
        ProcesarX_PT = True
        Exit Function
    End If

    ' final de sílaba ? /?/
    If sig = "" Then
        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "44,"
        ProcesarX_PT = True
        Exit Function
    End If

    ' /ks/
    If sig = "c" Or sig = "s" Then
        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "34,42,"
        ProcesarX_PT = True
        Exit Function
    End If

    ' /gz/
    If sig = "g" Or sig = "z" Then
        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "35,43,"
        ProcesarX_PT = True
        Exit Function
    End If

    ' por defecto ? /?/
    ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "44,"
    ProcesarX_PT = True

End Function

Private Function EsVocal(ByVal c As String) As Boolean
    EsVocal = InStr("aeiouáéíóúâêôãõ", LCase$(c)) > 0
End Function

' ============================================================
'   PROCESAR R PORTUGUÉS
' ============================================================
Private Function ProcesarR_PT(ByVal silabaActual As String, ByVal grafema As String, ByVal antCh As String, ByVal sigCh As String) As Byte

    Dim sigLetra As String
    Dim antLetra As String
    Dim antVocal As Boolean
    Dim sigVocal As Boolean

    Dim SilabaAnterior As String
    Dim silabaSiguiente As String

    silabaActual = LimpiaSilaba(silabaActual)
    SilabaAnterior = LimpiaSilaba(AntSilaba)
    silabaSiguiente = LimpiaSilaba(SigSilaba)


' Vocal anterior
    antLetra = ""
    antVocal = False
    If antCh = "" Then
        If SilabaAnterior <> "" Then
            antLetra = Right$(SilabaAnterior, 1)
            antVocal = EsVocal(antLetra)
        End If
    Else
        antLetra = antCh
        antVocal = EsVocal(antLetra)
    End If
    
    ' Vocal siguiente
    sigLetra = ""
    sigVocal = False
    
    If sigCh = "" Then
        If silabaSiguiente <> "" Then
            sigLetra = Left$(silabaSiguiente, 1)
            sigVocal = EsVocal(sigLetra)
        End If
    Else
        sigLetra = sigCh
        sigVocal = EsVocal(sigLetra)
    End If
    
'    ' Vocal anterior
'    antLetra = ""
'    antVocal = False
'    If SilabaAnterior <> "" Then
'        antLetra = Right$(SilabaAnterior, 1)
'        antVocal = EsVocal(antLetra)
'    End If

'    ' Vocal siguiente (primero dentro de la misma sílaba)
'    sigLetra = ""
'    sigVocal = False
'
'    If Len(silabaActual) > 1 Then
'        sigLetra = Mid$(silabaActual, 2, 1)
'        sigVocal = EsVocal(sigLetra)
'    End If
'
'    ' Si no hay vocal dentro de la misma sílaba, mirar la siguiente
'    If sigLetra = "" And silabaSiguiente <> "" Then
'        sigLetra = Left$(silabaSiguiente, 1)
'        sigVocal = EsVocal(sigLetra)
'    End If

    ' ============================
    '   REGLAS DE LA "R" PT-EU
    ' ============================

    ' 1) rr ? /?/
    If grafema = "rr" Then
        ProcesarR_PT = 51
        Exit Function
    End If

    ' 2) Inicio de palabra ? /?/
    If antLetra = "" Then
        ProcesarR_PT = 51
        Exit Function
    End If

    ' 3) Intervocálica ? /?/
    If antVocal And sigVocal Then
        ProcesarR_PT = 59
        Exit Function
    End If

    ' 4) Inicio de sílaba tras consonante ? /?/
    If Not antVocal And sigVocal Then
        ProcesarR_PT = 59
        Exit Function
    End If

    ' 5) Final de sílaba o final de palabra ? /?/
    If sigLetra = "" Or Not sigVocal Then
        ProcesarR_PT = 59
        Exit Function
    End If

    ' 6) Por defecto ? /?/
    ProcesarR_PT = 51

End Function

' ============================================================
'   PROCESAR S PORTUGUÉS
' ============================================================
Private Function ProcesarS_PT(ByVal silabaActual As String, ByVal grafema As String, ByVal antCh As String, ByVal sigCh As String) As Byte

    Dim sigLetra As String
    Dim antLetra As String
    Dim antVocal As Boolean
    Dim sigVocal As Boolean

    Dim SilabaAnterior As String
    Dim silabaSiguiente As String
    
    silabaActual = LimpiaSilaba(silabaActual)
    SilabaAnterior = LimpiaSilaba(AntSilaba)
    silabaSiguiente = LimpiaSilaba(SigSilaba)

    ' Vocal anterior
    antLetra = ""
    antVocal = False
    If antCh = "" Then
        If SilabaAnterior <> "" Then
            antLetra = Right$(SilabaAnterior, 1)
            antVocal = EsVocal(antLetra)
        End If
    Else
        antLetra = antCh
        antVocal = EsVocal(antLetra)
    End If
    
    ' Vocal siguiente
    sigLetra = ""
    sigVocal = False
    
    If sigCh = "" Then
        If silabaSiguiente <> "" Then
            sigLetra = Left$(silabaSiguiente, 1)
            sigVocal = EsVocal(sigLetra)
        End If
    Else
        sigLetra = sigCh
        sigVocal = EsVocal(sigLetra)
    End If
    
    addLog "PROCEDIMIENTO: " & "ProcesarS_PT"
    addLog
    addLog "silabaActual: " & silabaActual
    addLog "silabaSiguiente: " & silabaSiguiente
    addLog "SilabaAnterior: " & SilabaAnterior
    addLog "grafema: " & grafema

    'addLog "sigLetra: " & sigLetra

    ' 1) ss ? /s/
    If antLetra = "s" Then
        ProcesarS_PT = 42
        'ProcesarS_PT = 44
        Exit Function
    End If
    
    If sigLetra = "s" Then
        'ProcesarS_PT = 42
        ProcesarS_PT = 44
        Exit Function
    End If

If antLetra = "" Then
    ProcesarS_PT = 42
    Exit Function
End If

    ' 2) intervocálica ? /z/
    If antVocal And sigVocal Then
        ProcesarS_PT = 43
        Exit Function
    End If

    ' 3) antes de consonante sorda ? /?/
    If sigLetra Like "[ptkf]" Then
        ProcesarS_PT = 44
        Exit Function
    End If

    ' 4) antes de consonante sonora ? /?/
    If sigLetra Like "[bdgvzmnrl]" Then
        ProcesarS_PT = 45 '46
        Exit Function
    End If

    ' 5) final de sílaba ? /?/
    If silabaSiguiente = "" Then
        ProcesarS_PT = 44
        'ProcesarS_PT = 42
        Exit Function
    End If

    ' 6) por defecto ? /s/
    ProcesarS_PT = 42

End Function

'Private Function ProcesarS_PT(ByVal silabaActual As String, ByVal grafema As String) As Byte
'
'    Dim sigLetra As String
'    Dim antLetra As String
'    Dim antVocal As Boolean
'    Dim sigVocal As Boolean
'
'    Dim SilabaAnterior As String
'    Dim silabaSiguiente As String
'
'    silabaActual = LimpiaSilaba(silabaActual)
'    silabaSiguiente = LimpiaSilaba(SigSilaba)
'    SilabaAnterior = LimpiaSilaba(AntSilaba)
'
'
'    addLog "PROCEDIMIENTO: " & "ProcesarS_PT"
'    addLog
'    addLog "silabaActual: " & silabaActual
'    addLog "silabaSiguiente: " & silabaSiguiente
'    addLog "SilabaAnterior: " & SilabaAnterior
'    addLog "grafema: " & grafema
'
'
'    ' Última letra de la sílaba anterior
'    antLetra = ""
'    If SilabaAnterior <> "" Then
'        antLetra = Right$(SilabaAnterior, 1)
'    End If
'
'    ' Primera letra de la sílaba siguiente
'    sigLetra = ""
'    If silabaSiguiente <> "" Then
'        sigLetra = Left$(silabaSiguiente, 1)
'    End If
'
'    antVocal = EsVocal(antLetra)
'    sigVocal = EsVocal(sigLetra)
'
'    ' ============================================
'    ' 1) "ss" ? siempre /s/
'    ' ============================================
'    If grafema = "ss" Then
'        ProcesarS_PT = 42   ' /s/
'        Exit Function
'    End If
'
'    ' ============================================
'    ' 2) Intervocálica ? /z/
'    '    (vocal + s + vocal)
'    ' ============================================
'    If antVocal And sigVocal Then
'        ProcesarS_PT = 43   ' /z/
'        Exit Function
'    End If
'
'    ' ============================================
'    ' 3) Final de sílaba ? /?/
'    '    (no hay sílaba siguiente o empieza por consonante)
'    ' ============================================
'    If silabaSiguiente = "" Or Not sigVocal Then
'        ProcesarS_PT = 44   ' /?/
'        Exit Function
'    End If
'
'    ' ============================================
'    ' 4) Antes de consonante sorda ? /?/
'    ' ============================================
'    If sigLetra Like "[ptkf]" Then
'        ProcesarS_PT = 44   ' /?/
'        Exit Function
'    End If
'
'    ' ============================================
'    ' 5) Antes de consonante sonora ? /?/
'    ' ============================================
'    If sigLetra Like "[bdgvzmnrl]" Then
'        ProcesarS_PT = 46   ' /?/
'        Exit Function
'    End If
'
'    ' ============================================
'    ' 6) Por defecto ? /s/
'    ' ============================================
'    ProcesarS_PT = 42
'
'
'End Function

'Private Function ProcesarS_PT(ant As String, sig As String, silaba As String) As Boolean
'
'
'Dim antVocal As Boolean
'Dim sigVocal As Boolean
'Dim sigLetra As String
'
'sigLetra = ""
'If SigSilaba <> "" Then sigLetra = Left$(SigSilaba, 1)
'
'antVocal = EsVocal(ant)
'sigVocal = EsVocal(sigLetra)
'
'' 1) Intervocálica ? /z/
'If antVocal And sigVocal Then
'    ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "43,"   ' /z/
'    Exit Function
'End If
'
'' 2) Final de sílaba ? /?/
'If SigSilaba = "" Or Not sigVocal Then
'    ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "44,"   ' /?/
'    Exit Function
'End If
'
'' 3) Antes de consonante sorda ? /?/
'If sigLetra Like "[ptkf]" Then
'    ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "44,"   ' /?/
'    Exit Function
'End If
'
'' 4) Antes de consonante sonora ? /?/
'If sigLetra Like "[bdgvzmnrl]" Then
'    ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "46,"   ' /?/
'    Exit Function
'End If
'
'' 5) Por defecto ? /s/
'ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "42,"
'End Function

'Private Function ProcesarS_PT(ant As String, sig As String) As Boolean
'
'    ' intervocálica ? /z/
'    If EsVocal(ant) And EsVocal(sig) Then
'        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "43,"
'        ProcesarS_PT = True
'        Exit Function
'    End If
'
'    ' final de sílaba ? /?/
'    If sig = "" Then
'        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "44,"
'        ProcesarS_PT = True
'        Exit Function
'    End If
'
'    ' antes de consonante sorda ? /?/
'    If sig Like "[ptkfscx]" Then
'        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "44,"
'        ProcesarS_PT = True
'        Exit Function
'    End If
'
'    ' antes de consonante sonora ? /z/
'    If sig Like "[bdgvzmnrl]" Then
'        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "43,"
'        ProcesarS_PT = True
'        Exit Function
'    End If
'
'    ' por defecto ? /s/
'    ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "42,"
'    ProcesarS_PT = True
'
'End Function


' ============================================================
'   REDUCCIÓN VOCÁLICA PT-EU
' ============================================================

Private Function ProcesarReduccionVocal_PT(grafema As String, esAtona As Boolean) As Boolean

    If Not esAtona Then
        ProcesarReduccionVocal_PT = False
        Exit Function
    End If

    Select Case grafema

        ' a átona ? /?/ ? ID 9
        Case "a"
            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "9,"
            ProcesarReduccionVocal_PT = True
            Exit Function

        ' e átona ? /?/ ? ID 10
        Case "e"
            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "10,"
            ProcesarReduccionVocal_PT = True
            Exit Function

        ' o átono ? /u/ ? ID 7
        Case "o"
            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "7,"
            ProcesarReduccionVocal_PT = True
            Exit Function

    End Select

    ProcesarReduccionVocal_PT = False

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


'' ============================================================
''   GENERAR IPA A PARTIR DE ObjDTO.IdsFonemas
''   (Función universal para todos los idiomas)
'' ============================================================
'
'Private Function GenerarIPA() As String
'
'    Dim arr As Variant
'    Dim out As String
'    Dim i As Long
'    Dim id As String
'    Dim ipa As String
'
'    If ObjDTO.IdsFonemas = "" Then
'        GenerarIPA = ""
'        Exit Function
'    End If
'
'    arr = Split(ObjDTO.IdsFonemas, ",")
'
'    For i = LBound(arr) To UBound(arr)
'
'        id = Trim$(arr(i))
'        If id = "" Then GoTo siguiente
'
'        ipa = ObtenerIPA_DesdeID(id)
'
'        If ipa <> "" Then
'            out = out & ipa
'        End If
'
'siguiente:
'    Next i
'
'    GenerarIPA = out
'
'End Function
'
'
'' ============================================================
''   MAPEO ID ? SÍMBOLO AFI
''   (Usa tu tabla qryFonemasValor)
'' ============================================================
'
'Private Function ObtenerIPA_DesdeID(ByVal id As String) As String
'
'    Static dic As Object
'    Dim rs As DAO.Recordset
'
'    If dic Is Nothing Then
'        Set dic = CreateObject("Scripting.Dictionary")
'
'        Set rs = CurrentDb.OpenRecordset( _
'            "SELECT ID, IPA FROM qryFonemasValor")
'
'        Do Until rs.EOF
'            dic(CStr(rs!id)) = rs!ipa
'            rs.MoveNext
'        Loop
'
'        rs.Close
'    End If
'
'    If dic.Exists(id) Then
'        ObtenerIPA_DesdeID = dic(id)
'    Else
'        ObtenerIPA_DesdeID = ""
'    End If
'
'End Function


