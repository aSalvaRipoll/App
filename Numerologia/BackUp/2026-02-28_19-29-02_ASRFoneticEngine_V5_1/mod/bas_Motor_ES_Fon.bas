Attribute VB_Name = "bas_Motor_ES_Fon"

Option Compare Database
Option Explicit


' ============================================
' PUNTO DE ENTRADA AL PROCESADOR FONÉTICO
' ============================================
Public Sub ConstruirCadenaFonemas_ES()

    Dim arrSilabas As Variant
    Dim i As Long
    Dim silabaCruda As String
    Dim ultimaSilaba As String
    Dim siguienteSilaba As String
    
    Dim frase As String
    Dim arrFon As Variant
    Dim Fon As Variant
    
    Dim ligaduraID As String
    Dim res As String
    
    ObjDTO.IdsFonemas = ""
    ObjDTO.FonemasFinal = ""

    '0 Normalizar vocales
    frase = NormalizarVocales(ObjDTO.SilabasFinal)

    ' 1. Separar sílabas por "|"
    arrSilabas = Split(frase, "|")

    For i = 0 To UBound(arrSilabas)

        silabaCruda = arrSilabas(i)

        ' ---------------------------------------------------------
        ' 2. Detectar separador de palabra (sílaba vacía)
        ' ---------------------------------------------------------
        If Trim$(silabaCruda) = "" Then

            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "#"

            ' Ligadura automàtica (Mode 2)
            'If CFG.ModoLigadura = 2 Then

                ultimaSilaba = BuscarSilabaRealAnterior(arrSilabas, i)
                siguienteSilaba = BuscarSilabaRealPosterior(arrSilabas, i)

                If HayLigaduraAutomatica(ultimaSilaba, siguienteSilaba) Then
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
        ' 3. Insertar modificadores prosódicos (acento)
        ' ---------------------------------------------------------
        If InStr(silabaCruda, "(") > 0 Then
            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "80,"
        ElseIf InStr(silabaCruda, "[") > 0 Then
            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "81,"
        Else
            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "82,"
        End If

        ' ---------------------------------------------------------
        ' 4. Procesar grafemas
        ' ---------------------------------------------------------
        ProcesarSilaba silabaCruda

        ' ---------------------------------------------------------
        ' 5. Separador silábico
        ' ---------------------------------------------------------
        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "|"

SiguienteIteracion:
    Next i

'    For i = 0 To UBound(arrSilabas)
'
'        silabaCruda = Trim$(arrSilabas(i))
'
'        ' ---------------------------------------------------------
'        ' 2. Detectar separador de palabra (sílaba vacía)
'        ' ---------------------------------------------------------
'        If silabaCruda = "" Then
'
'            ' Insertar separador estructural de palabra
'            If Trim$(silabaCruda) = "" Then
'
'                ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "#"
'
'                ' Ligadura automàtica (Mode 2)
'                'If CFG.ModoLigadura = 2 Then
'
'                    ultimaSilaba = BuscarSilabaRealAnterior(arrSilabas, i)
'                    siguienteSilaba = BuscarSilabaRealPosterior(arrSilabas, i)
'
'                    If HayLigaduraAutomatica(ultimaSilaba, siguienteSilaba) Then
'                        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "84,"
'                    End If
'
'                'End If
'
'                GoTo SiguienteIteracion
'            End If
'
'            ' 1. Ligadura manual (Mode 1)
'            'If CFG.ModoLigadura = 1 Then
'                ligaduraID = DetectarLigaduraManual(silabaCruda)
'                If ligaduraID <> 0 Then
'                    ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & CStr(ligaduraID) & ","
'                End If
'        'End If
'
'        ' ---------------------------------------------------------
'        ' 3. Insertar modificadores prosódicos (acento)
'        ' ---------------------------------------------------------
'        If InStr(silabaCruda, "(") > 0 Then
'            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "80,"
'        ElseIf InStr(silabaCruda, "[") > 0 Then
'            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "81,"
'        Else
'            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "82,"
'        End If
'
'        ' Limpiar marcadores
'        silabaCruda = Replace(silabaCruda, "(", "")
'        silabaCruda = Replace(silabaCruda, ")", "")
'        silabaCruda = Replace(silabaCruda, "[", "")
'        silabaCruda = Replace(silabaCruda, "]", "")
'
'        ' ---------------------------------------------------------
'        ' 4. Procesar grafemas
'        ' ---------------------------------------------------------
'        ProcesarSilaba silabaCruda
'
'        ' ---------------------------------------------------------
'        ' 5. Separador silábico
'        ' ---------------------------------------------------------
'        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "|"
'
'SiguienteIteracion:
'    Next i

    ' ---------------------------------------------------------
    ' 6. Limpieza final
    ' ---------------------------------------------------------
    ObjDTO.IdsFonemas = Replace(ObjDTO.IdsFonemas, "|#", "#")
    ObjDTO.IdsFonemas = Replace(ObjDTO.IdsFonemas, "#|", "#")
    ObjDTO.IdsFonemas = Replace(ObjDTO.IdsFonemas, ",|", "|")
    ObjDTO.IdsFonemas = Replace(ObjDTO.IdsFonemas, "84,84,", "84,")

    While Right$(ObjDTO.IdsFonemas, 1) = "|" Or Right$(ObjDTO.IdsFonemas, 1) = "#" Or Right$(ObjDTO.IdsFonemas, 1) = ","
        ObjDTO.IdsFonemas = Left$(ObjDTO.IdsFonemas, Len(ObjDTO.IdsFonemas) - 1)
    Wend

    arrSilabas = Split(ObjDTO.IdsFonemas, "|")
    res = ""
        
        For i = 0 To UBound(arrSilabas)
            If i = 0 Then
                res = res & "/"
            End If
            
            arrFon = Split(arrSilabas(i), ",")
            For Each Fon In arrFon
                If Fon = "#84" Then
                    Fon = 84
                ElseIf Fon = "#82" Then
                    res = res & "/ /"
                    Fon = 82
                End If
                
                If Left(Fon, 1) = "#" Then
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

'---------------------------------------------------------
'                   FUNCIONES AUXILIARES
'---------------------------------------------------------
Private Sub ProcesarSilaba(ByVal silabaCruda As String)

    Dim silaba As String
    Dim ligaduraID As Byte
    Dim i As Long
    Dim grafema As String
    Dim id As Byte
    Dim antCh As String
    Dim sigCh As String

    silaba = Trim$(silabaCruda)
    
'    ' 1. Detectar ligadura manual (solo Modo 1)
'    'If CFG.ModoLigadura = 1 Then
'        ligaduraID = DetectarLigaduraManual(silaba)
'        If ligaduraID <> 0 Then
'            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & CStr(ligaduraID) & ","
'        End If
'    'End If

    ' 2. Procesar grafemas (con detección real de dígrafos)
    i = 1
    Do While i <= Len(silaba)

        '-----------------------------------------
        ' Detectar grafema (1 o 2 letras)
        '-----------------------------------------
        grafema = DetectarDigrafo(silaba, i)

        '-----------------------------------------
        ' Calcular contexto anterior
        '-----------------------------------------
        If i > 1 Then
            antCh = Mid$(silaba, i - 1, 1)
        Else
            antCh = ""
        End If

        '-----------------------------------------
        ' Calcular contexto siguiente
        '-----------------------------------------
        If i + Len(grafema) <= Len(silaba) Then
            sigCh = Mid$(silaba, i + Len(grafema), 1)
        Else
            sigCh = ""
        End If

        '-----------------------------------------
        ' Obtener fonema
        '-----------------------------------------
        id = AsignarFonemaBase(grafema, sigCh, antCh)

        '-----------------------------------------
        ' Añadir fonema
        '-----------------------------------------
        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & CStr(id) & ","

        '-----------------------------------------
        ' Avanzar según tamaño del grafema
        '-----------------------------------------
        i = i + Len(grafema)

    Loop

End Sub

Private Function NormalizarGrafema(ByVal texto As String) As String

    texto = LCase$(texto)

    texto = Replace(texto, "á", "a")
    texto = Replace(texto, "é", "e")
    texto = Replace(texto, "í", "i")
    texto = Replace(texto, "ó", "o")
    texto = Replace(texto, "ú", "u")

    texto = Replace(texto, "h", "")

    NormalizarGrafema = texto
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

Private Function DetectarDigrafo(texto As String, pos As Long) As String

    Dim par As String, sig As String

    If pos + 1 > Len(texto) Then
        DetectarDigrafo = Mid$(texto, pos, 1)
        Exit Function
    End If

    par = Mid$(texto, pos, 2)
    sig = Mid$(texto, pos + 2, 1)

    Select Case par
        Case "ch", "ll", "rr", "qu"
            DetectarDigrafo = par
            Exit Function
    End Select

    If par = "gu" And (sig = "e" Or sig = "i") Then
        DetectarDigrafo = "gu"
        Exit Function
    End If

    If par = "hi" And sig Like "[aeiou]" Then
        DetectarDigrafo = "hi"
        Exit Function
    End If

    ' Si no es dígrafo ? 1 letra
    DetectarDigrafo = Mid$(texto, pos, 1)

End Function

'===========================================================
' MOTOR FONÉTICO ÚNICO (CASTELLANO)
'===========================================================
Private Function AsignarFonemaBase(grafema As String, _
                             Optional ByVal sig As String = "", _
                             Optional ByVal ant As String = "") As Byte

    grafema = LCase$(grafema)
    sig = LCase$(sig)
    ant = LCase$(ant)

    '===========================================================
    ' VOCAL
    '===========================================================
    Select Case grafema
        Case "a": AsignarFonemaBase = 1: Exit Function
        Case "e": AsignarFonemaBase = 2: Exit Function
        Case "i": AsignarFonemaBase = 4: Exit Function
        Case "o": AsignarFonemaBase = 5: Exit Function
        Case "u": AsignarFonemaBase = 7: Exit Function
    End Select

    '===========================================================
    ' OCLUSIVAS
    '===========================================================
    Select Case grafema
        Case "p": AsignarFonemaBase = 30: Exit Function
        Case "b", "v": AsignarFonemaBase = 31: Exit Function
        Case "t": AsignarFonemaBase = 32: Exit Function
        Case "d": AsignarFonemaBase = 33: Exit Function
    End Select

    '===========================================================
    ' K
    '===========================================================
    If grafema = "k" Then AsignarFonemaBase = 34: Exit Function

    '===========================================================
    ' C
    '===========================================================
    If grafema = "c" Then
        If sig = "e" Or sig = "i" Then AsignarFonemaBase = 46 Else AsignarFonemaBase = 34
        Exit Function
    End If

    '===========================================================
    ' QU
    '===========================================================
    If grafema = "qu" Then AsignarFonemaBase = 34: Exit Function

    '===========================================================
    ' GU
    '===========================================================
    If grafema = "gu" Then AsignarFonemaBase = 35: Exit Function

    '===========================================================
    ' G
    '===========================================================
    If grafema = "g" Then
        If sig = "e" Or sig = "i" Then AsignarFonemaBase = 48 Else AsignarFonemaBase = 35
        Exit Function
    End If

    '===========================================================
    ' FRICATIVAS (solo las no dialectales)
    '===========================================================
    Select Case grafema
        Case "f": AsignarFonemaBase = 40: Exit Function   ' /f/
        Case "j": AsignarFonemaBase = 48: Exit Function   ' /x/
    End Select

    '===========================================================
    ' X (contextual)
    '===========================================================
    If grafema = "x" Then

        ' 1) X inicial + vocal ? /s/
        If ant = "" And sig Like "[aeiouáéíóúü]" Then
            AsignarFonemaBase = 42
            Exit Function
        End If

        ' 2) Vocal + X + vocal ? /gz/
        If ant Like "[aeiouáéíóúü]" And sig Like "[aeiouáéíóúü]" Then
            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "35,43,"
            AsignarFonemaBase = 0
            Exit Function
        End If

        ' 3) Vocal + X + consonante ? /ks/
        If ant Like "[aeiouáéíóúü]" And sig Like "[bcdfghjklmnñpqrstvwxyz]" Then
            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "34,42,"
            AsignarFonemaBase = 0
            Exit Function
        End If

        ' 4) Topónimos mexicanos ? /x/
'        If LCase$(ObjDTO.palabra) = "méxico" Or _
'           LCase$(ObjDTO.palabra) = "oaxaca" Or _
'           LCase$(ObjDTO.palabra) = "xola" Then
'            AsignarFonemaBase = 48
'            Exit Function
'        End If

        ' 5) Por defecto ? /ks/
        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "34,42,"
        AsignarFonemaBase = 0
        Exit Function

    End If

    '===========================================================
    ' NASALES
    '===========================================================
    Select Case grafema
        Case "m": AsignarFonemaBase = 36: Exit Function
        Case "n": AsignarFonemaBase = 37: Exit Function
        Case "ñ": AsignarFonemaBase = 38: Exit Function
    End Select

    '===========================================================
    ' LATERALES
    '===========================================================
    If grafema = "l" Then AsignarFonemaBase = 62: Exit Function
    If grafema = "ll" Then AsignarFonemaBase = 63: Exit Function

    '===========================================================
    ' VIBRANTES
    '===========================================================
    If grafema = "rr" Then AsignarFonemaBase = 60: Exit Function

    If grafema = "r" Then
        If ant = "" Or Not ant Like "[aeiou]" Then AsignarFonemaBase = 60 Else AsignarFonemaBase = 59
        Exit Function
    End If

    '===========================================================
    ' AFRICADA
    '===========================================================
    If grafema = "ch" Then AsignarFonemaBase = 57: Exit Function

    '===========================================================
    ' Y /?/
    '===========================================================
    If grafema = "y" Or grafema = "hi" Then AsignarFonemaBase = 45: Exit Function

    '===========================================================
    ' MODOS DE SIBILANTES (C/Z/S)
    '===========================================================
    Select Case CFG.ModoSibilantes

        Case 0, 1   ' Distinción
            If grafema = "c" And (sig = "e" Or sig = "i") Then AsignarFonemaBase = 46: Exit Function
            If grafema = "z" Then AsignarFonemaBase = 46: Exit Function
            If grafema = "s" Then AsignarFonemaBase = 42: Exit Function

        Case 2       ' Seseo
            If grafema = "c" And (sig = "e" Or sig = "i") Then AsignarFonemaBase = 42: Exit Function
            If grafema = "z" Then AsignarFonemaBase = 42: Exit Function
            If grafema = "s" Then AsignarFonemaBase = 42: Exit Function

        Case 3       ' Ceceo
            If grafema = "c" And (sig = "e" Or sig = "i") Then AsignarFonemaBase = 46: Exit Function
            If grafema = "z" Then AsignarFonemaBase = 46: Exit Function
            If grafema = "s" Then AsignarFonemaBase = 46: Exit Function

    End Select

    '===========================================================
    ' MODOS LATERALES (LL/Y)
    '===========================================================
    Select Case CFG.ModoLateral

        Case 0   ' Normal
            If grafema = "ll" Then AsignarFonemaBase = 63: Exit Function
            If grafema = "y" Then AsignarFonemaBase = 45: Exit Function

        Case 1   ' Lleísmo
            If grafema = "ll" Then AsignarFonemaBase = 63: Exit Function
            If grafema = "y" Then AsignarFonemaBase = 63: Exit Function

        Case 2   ' Yeísmo
            If grafema = "ll" Then AsignarFonemaBase = 45: Exit Function
            If grafema = "y" Then AsignarFonemaBase = 45: Exit Function

        Case 3   ' Yeísmo rehilado
            If grafema = "ll" Then AsignarFonemaBase = 72: Exit Function
            If grafema = "y" Then AsignarFonemaBase = 72: Exit Function

    End Select

    '===========================================================
    ' H MUDA
    '===========================================================
    If grafema = "h" Then AsignarFonemaBase = 0: Exit Function
    
    '===========================================================
    ' Si no se reconoce
    '===========================================================
    AsignarFonemaBase = 255

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

Private Function NormalizarVocales(ByVal texto As String) As String

    texto = Replace(texto, "á", "a")
    texto = Replace(texto, "à", "a")
    texto = Replace(texto, "ä", "a")
    texto = Replace(texto, "â", "a")

    texto = Replace(texto, "é", "e")
    texto = Replace(texto, "è", "e")
    texto = Replace(texto, "ë", "e")
    texto = Replace(texto, "ê", "e")

    texto = Replace(texto, "í", "i")
    texto = Replace(texto, "ì", "i")
    texto = Replace(texto, "ï", "i")
    texto = Replace(texto, "î", "i")

    texto = Replace(texto, "ó", "o")
    texto = Replace(texto, "ò", "o")
    texto = Replace(texto, "ö", "o")
    texto = Replace(texto, "ô", "o")

    texto = Replace(texto, "ú", "u")
    texto = Replace(texto, "ù", "u")
    texto = Replace(texto, "û", "u")

    ' ü se mantiene
    NormalizarVocales = texto

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
