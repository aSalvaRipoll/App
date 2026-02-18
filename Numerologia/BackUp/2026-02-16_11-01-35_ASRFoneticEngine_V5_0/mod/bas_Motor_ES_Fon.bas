Attribute VB_Name = "bas_Motor_ES_Fon"

Option Compare Database
Option Explicit

' Caché en memoria para acelerar búsquedas
Private IPA_Cache As Scripting.Dictionary

' ============================================
' PUNTO DE ENTRADA AL PROCESADOR FONÉTICO
' ============================================
Public Sub ConstruirCadenaFonemas_ES()

    Dim arrSilabas As Variant
    Dim i As Long
    Dim silaba As Variant
    Dim ultimaSilaba As String
    Dim SiguienteSilaba As String
    
    Dim frase As String
    Dim arrFon As Variant
    Dim Fon As Variant
    
    Dim res As String
    
    ObjDTO.IdsFonemas = ""
    ObjDTO.FonemasFinal = ""

    '0 Normalizar vocales
    frase = NormalizarVocales_ES(ObjDTO.SilabasFinal)

    ' 1. Separar sílabas por "|"
    arrSilabas = Split(frase, "|")

    For i = 0 To UBound(arrSilabas)

        silaba = Trim$(arrSilabas(i))

        ' ---------------------------------------------------------
        ' 2. Detectar separador de palabra (sílaba vacía)
        ' ---------------------------------------------------------
        If silaba = "" Then

            ' Insertar separador estructural de palabra
            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "#"

            ' Ligadura automática (Modo 2)
            If CFG.ModoLigadura = 2 Then

                ultimaSilaba = BuscarSilabaRealAnterior(arrSilabas, i)
                SiguienteSilaba = BuscarSilabaRealPosterior(arrSilabas, i)

                If HayLigaduraAutomatica(ultimaSilaba, SiguienteSilaba) Then
                    ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "84,"
                End If

            End If

            GoTo SiguienteIteracion
        
        End If

        ' ---------------------------------------------------------
        ' 3. Insertar modificadores prosódicos (acento)
        ' ---------------------------------------------------------
        If InStr(silaba, "(") > 0 Then
            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "80,"
        ElseIf InStr(silaba, "[") > 0 Then
            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "81,"
        Else
            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "82,"
        End If

        ' Limpiar marcadores
        silaba = Replace(silaba, "(", "")
        silaba = Replace(silaba, ")", "")
        silaba = Replace(silaba, "[", "")
        silaba = Replace(silaba, "]", "")

        ' ---------------------------------------------------------
        ' 4. Procesar grafemas
        ' ---------------------------------------------------------
        ProcesarSilaba silaba

        ' ---------------------------------------------------------
        ' 5. Separador silábico
        ' ---------------------------------------------------------
        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "|"

SiguienteIteracion:
    Next i

    ' ---------------------------------------------------------
    ' 6. Limpieza final
    ' ---------------------------------------------------------
    ObjDTO.IdsFonemas = Replace(ObjDTO.IdsFonemas, "|#", "#")
    ObjDTO.IdsFonemas = Replace(ObjDTO.IdsFonemas, "#|", "#")

    If Right$(ObjDTO.IdsFonemas, 1) = "|" Or Right$(ObjDTO.IdsFonemas, 1) = "#" Or Right$(ObjDTO.IdsFonemas, 1) = "," Then
        ObjDTO.IdsFonemas = Left$(ObjDTO.IdsFonemas, Len(ObjDTO.IdsFonemas) - 1)
    End If

    ObjDTO.IdsFonemas = Replace(ObjDTO.IdsFonemas, ",|", "|")
    
    arrSilabas = Split(ObjDTO.IdsFonemas, "|")
    res = ""
    'ObjDTO.FonemasFinal = ObjDTO.FonemasFinal & "/"
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
        'ObjDTO.FonemasFinal = ObjDTO.FonemasFinal & "/ "
        
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
    
    ' Normalizar vocales
'    silaba = NormalizarVocales_ES(silaba)
    
    ' 1. Detectar ligadura manual (solo Modo 1)
    If CFG.ModoLigadura = 1 Then
        ligaduraID = DetectarLigaduraManual(silaba)
        If ligaduraID <> 0 Then
            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & CStr(ligaduraID) & ","
        End If
    End If

    ' 2. Procesar grafemas (con detección real de dígrafos)
    i = 1
    Do While i <= Len(silaba)

        '-----------------------------------------
        ' Detectar grafema (1 o 2 letras)
        '-----------------------------------------
        grafema = DetectarDigrafoEs(silaba, i)

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
        id = AsignarFonemaBaseEs(grafema, sigCh, antCh)

        '-----------------------------------------
        ' Añadir fonema
        '-----------------------------------------
        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & CStr(id) & ","

        '-----------------------------------------
        ' Avanzar según tamaño del grafema
        '-----------------------------------------
        i = i + Len(grafema)

    Loop

'    If Right(ObjDTO.IdsFonemas, 1) = "," Then _
        ObjDTO.IdsFonemas = Left(ObjDTO.IdsFonemas, Len(ObjDTO.IdsFonemas) - 1)

End Sub

'Private Sub ProcesarSilaba_1(ByVal silabaCruda As String)
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
'    ' 2. Procesar grafemas (incluye dígrafos)
'    i = 1
'    Do While i <= Len(silaba)
'
'        '-----------------------------
'        ' Intentar grafema de 2 letras
'        '-----------------------------
'        If i < Len(silaba) Then
'
'            grafema = Mid$(silaba, i, 2)
'
'            ' contexto anterior
'            If i > 1 Then
'                antCh = Mid$(silaba, i - 1, 1)
'            Else
'                antCh = ""
'            End If
'
'            ' contexto siguiente (después del dígrafo)
'            If i + 2 <= Len(silaba) Then
'                sigCh = Mid$(silaba, i + 2, 1)
'            Else
'                sigCh = ""
'            End If
'
'            id = AsignarFonemaBaseEs(grafema, sigCh, antCh)
'
'        Else
'            id = 0
'        End If
'
'        '-----------------------------
'        ' Si no existe dígrafo, 1 letra
'        '-----------------------------
'        If id = 0 Then
'
'            grafema = Mid$(silaba, i, 1)
'
'            ' contexto anterior
'            If i > 1 Then
'                antCh = Mid$(silaba, i - 1, 1)
'            Else
'                antCh = ""
'            End If
'
'            ' contexto siguiente
'            If i + 1 <= Len(silaba) Then
'                sigCh = Mid$(silaba, i + 1, 1)
'            Else
'                sigCh = ""
'            End If
'
'            id = AsignarFonemaBaseEs(grafema, sigCh, antCh)
'
'        End If
'
'        ' Añadir fonema
'        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & CStr(id) & ","
'
'        ' Avanzar según el tamaño del grafema
'        i = i + Len(grafema)
'
'    Loop
'
'End Sub

'Private Sub ProcesarSilaba_0(ByVal silabaCruda As String)
'
'    Dim silaba As String
'    Dim ligaduraID As Byte
'    Dim i As Long
'    Dim grafema As String
'    Dim id As Byte
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
'    ' 2. Procesar grafemas (incluye dígrafos)
'    i = 1
'    Do While i <= Len(silaba)
'
'        ' Intentar grafema de 2 letras
'        If i < Len(silaba) Then
'            grafema = Mid$(silaba, i, 2)
'            id = AsignarFonemaBaseEs(grafema, Mid$(silaba, i + 2, 1), Mid$(silaba, i - 1, 1))
'        Else
'            id = 0
'        End If
'
'        ' Si no existe dígrafo, usar grafema de 1 letra
'        If id = 0 Then
'            grafema = Mid$(silaba, i, 1)
'            id = AsignarFonemaBaseEs(grafema, Mid$(silaba, i + 1, 1), Mid$(silaba, i - 1, 1))
'        End If
'
'        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & CStr(id) & ","
'
'        i = i + Len(grafema)
'
'    Loop
'
'End Sub


Function NormalizarGrafemaEs(ByVal Texto As String) As String

    Texto = LCase$(Texto)

    Texto = Replace(Texto, "á", "a")
    Texto = Replace(Texto, "é", "e")
    Texto = Replace(Texto, "í", "i")
    Texto = Replace(Texto, "ó", "o")
    Texto = Replace(Texto, "ú", "u")

    Texto = Replace(Texto, "h", "")

    NormalizarGrafemaEs = Texto
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

Function DetectarDigrafoEs(Texto As String, pos As Long) As String

    Dim par As String, sig As String

    If pos + 1 > Len(Texto) Then
        DetectarDigrafoEs = Mid$(Texto, pos, 1)
        Exit Function
    End If

    par = Mid$(Texto, pos, 2)
    sig = Mid$(Texto, pos + 2, 1)

    Select Case par
        Case "ch", "ll", "rr", "qu"
            DetectarDigrafoEs = par
            Exit Function
    End Select

    If par = "gu" And (sig = "e" Or sig = "i") Then
        DetectarDigrafoEs = "gu"
        Exit Function
    End If

    If par = "hi" And sig Like "[aeiou]" Then
        DetectarDigrafoEs = "hi"
        Exit Function
    End If

    ' Si no es dígrafo ? 1 letra
    DetectarDigrafoEs = Mid$(Texto, pos, 1)

End Function

'Function DetectarDigrafoEs(texto As String, pos As Long) As String
'
'    Dim par As String, sig As String
'
'    If pos + 1 > Len(texto) Then
'        DetectarDigrafoEs = Mid$(texto, pos, 1)
'        Exit Function
'    End If
'
'    par = Mid$(texto, pos, 2)
'    sig = Mid$(texto, pos + 2, 1)
'
'    Select Case par
'        Case "ch", "ll", "rr", "qu"
'            DetectarDigrafoEs = par
'            Exit Function
'    End Select
'
'    If par = "gu" And (sig = "e" Or sig = "i") Then
'        DetectarDigrafoEs = "gu"
'        Exit Function
'    End If
'
'    If par = "hi" And sig Like "[aeiou]" Then
'        DetectarDigrafoEs = "hi"
'        Exit Function
'    End If
'
'    DetectarDigrafoEs = Mid$(texto, pos, 1)
'End Function


'===========================================================
' MOTOR FONÉTICO ÚNICO (CASTELLANO)
'===========================================================
Function AsignarFonemaBaseEs(grafema As String, _
                             Optional ByVal sig As String = "", _
                             Optional ByVal ant As String = "") As Byte

    grafema = LCase$(grafema)
    sig = LCase$(sig)
    ant = LCase$(ant)

    '===========================================================
    ' VOCAL
    '===========================================================
    Select Case grafema
        Case "a": AsignarFonemaBaseEs = 1: Exit Function
        Case "e": AsignarFonemaBaseEs = 2: Exit Function
        Case "i": AsignarFonemaBaseEs = 4: Exit Function
        Case "o": AsignarFonemaBaseEs = 5: Exit Function
        Case "u": AsignarFonemaBaseEs = 7: Exit Function
    End Select

    '===========================================================
    ' OCLUSIVAS
    '===========================================================
    Select Case grafema
        Case "p": AsignarFonemaBaseEs = 30: Exit Function
        Case "b", "v": AsignarFonemaBaseEs = 31: Exit Function
        Case "t": AsignarFonemaBaseEs = 32: Exit Function
        Case "d": AsignarFonemaBaseEs = 33: Exit Function
    End Select

    '===========================================================
    ' K
    '===========================================================
    If grafema = "k" Then AsignarFonemaBaseEs = 34: Exit Function

    '===========================================================
    ' C
    '===========================================================
    If grafema = "c" Then
        If sig = "e" Or sig = "i" Then AsignarFonemaBaseEs = 46 Else AsignarFonemaBaseEs = 34
        Exit Function
    End If

    '===========================================================
    ' QU
    '===========================================================
    If grafema = "qu" Then AsignarFonemaBaseEs = 34: Exit Function

    '===========================================================
    ' G
    '===========================================================
    If grafema = "g" Then
        If sig = "e" Or sig = "i" Then AsignarFonemaBaseEs = 48 Else AsignarFonemaBaseEs = 35
        Exit Function
    End If

    '===========================================================
    ' GU
    '===========================================================
    If grafema = "gu" Then AsignarFonemaBaseEs = 35: Exit Function

    '===========================================================
    ' FRICATIVAS (solo las no dialectales)
    '===========================================================
    Select Case grafema
        Case "f": AsignarFonemaBaseEs = 40: Exit Function   ' /f/
        Case "j": AsignarFonemaBaseEs = 48: Exit Function   ' /x/
    End Select

    '===========================================================
    ' X (contextual)
    '===========================================================
    If grafema = "x" Then

        ' 1) X inicial + vocal ? /s/
        If ant = "" And sig Like "[aeiouáéíóúü]" Then
            AsignarFonemaBaseEs = 42
            Exit Function
        End If

        ' 2) Vocal + X + vocal ? /gz/
        If ant Like "[aeiouáéíóúü]" And sig Like "[aeiouáéíóúü]" Then
            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "35,43,"
            AsignarFonemaBaseEs = 0
            Exit Function
        End If

        ' 3) Vocal + X + consonante ? /ks/
        If ant Like "[aeiouáéíóúü]" And sig Like "[bcdfghjklmnñpqrstvwxyz]" Then
            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "34,42,"
            AsignarFonemaBaseEs = 0
            Exit Function
        End If

        ' 4) Topónimos mexicanos ? /x/
'        If LCase$(ObjDTO.palabra) = "méxico" Or _
'           LCase$(ObjDTO.palabra) = "oaxaca" Or _
'           LCase$(ObjDTO.palabra) = "xola" Then
'            AsignarFonemaBaseEs = 48
'            Exit Function
'        End If

        ' 5) Por defecto ? /ks/
        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "34,42,"
        AsignarFonemaBaseEs = 0
        Exit Function

    End If

    '===========================================================
    ' NASALES
    '===========================================================
    Select Case grafema
        Case "m": AsignarFonemaBaseEs = 36: Exit Function
        Case "n": AsignarFonemaBaseEs = 37: Exit Function
        Case "ñ": AsignarFonemaBaseEs = 38: Exit Function
    End Select

    '===========================================================
    ' LATERALES
    '===========================================================
    If grafema = "l" Then AsignarFonemaBaseEs = 62: Exit Function
    If grafema = "ll" Then AsignarFonemaBaseEs = 63: Exit Function

    '===========================================================
    ' VIBRANTES
    '===========================================================
    If grafema = "rr" Then AsignarFonemaBaseEs = 60: Exit Function

    If grafema = "r" Then
        If ant = "" Or Not ant Like "[aeiou]" Then AsignarFonemaBaseEs = 60 Else AsignarFonemaBaseEs = 59
        Exit Function
    End If

    '===========================================================
    ' AFRICADA
    '===========================================================
    If grafema = "ch" Then AsignarFonemaBaseEs = 57: Exit Function

    '===========================================================
    ' Y /?/
    '===========================================================
    If grafema = "y" Or grafema = "hi" Then AsignarFonemaBaseEs = 45: Exit Function

    '===========================================================
    ' MODOS DE SIBILANTES (C/Z/S)
    '===========================================================
    Select Case CFG.ModoSibilantes

        Case 0, 1   ' Distinción
            If grafema = "c" And (sig = "e" Or sig = "i") Then AsignarFonemaBaseEs = 46: Exit Function
            If grafema = "z" Then AsignarFonemaBaseEs = 46: Exit Function
            If grafema = "s" Then AsignarFonemaBaseEs = 42: Exit Function

        Case 2       ' Seseo
            If grafema = "c" And (sig = "e" Or sig = "i") Then AsignarFonemaBaseEs = 42: Exit Function
            If grafema = "z" Then AsignarFonemaBaseEs = 42: Exit Function
            If grafema = "s" Then AsignarFonemaBaseEs = 42: Exit Function

        Case 3       ' Ceceo
            If grafema = "c" And (sig = "e" Or sig = "i") Then AsignarFonemaBaseEs = 46: Exit Function
            If grafema = "z" Then AsignarFonemaBaseEs = 46: Exit Function
            If grafema = "s" Then AsignarFonemaBaseEs = 46: Exit Function

    End Select

    '===========================================================
    ' MODOS LATERALES (LL/Y)
    '===========================================================
    Select Case CFG.ModoLateral

        Case 0   ' Normal
            If grafema = "ll" Then AsignarFonemaBaseEs = 63: Exit Function
            If grafema = "y" Then AsignarFonemaBaseEs = 45: Exit Function

        Case 1   ' Lleísmo
            If grafema = "ll" Then AsignarFonemaBaseEs = 63: Exit Function
            If grafema = "y" Then AsignarFonemaBaseEs = 63: Exit Function

        Case 2   ' Yeísmo
            If grafema = "ll" Then AsignarFonemaBaseEs = 45: Exit Function
            If grafema = "y" Then AsignarFonemaBaseEs = 45: Exit Function

        Case 3   ' Yeísmo rehilado
            If grafema = "ll" Then AsignarFonemaBaseEs = 72: Exit Function
            If grafema = "y" Then AsignarFonemaBaseEs = 72: Exit Function

    End Select

    '===========================================================
    ' Si no se reconoce
    '===========================================================
    AsignarFonemaBaseEs = 255

End Function

'Function AsignarFonemaBaseEs(grafema As String, _
'                             Optional ByVal sig As String = "", _
'                             Optional ByVal ant As String = "") As Byte
'
'    grafema = LCase$(grafema)
'    sig = LCase$(sig)
'    ant = LCase$(ant)
'
'    '===========================================================
'    ' VOCAL
'    '===========================================================
'    Select Case grafema
'        Case "a": AsignarFonemaBaseEs = 1: Exit Function
'        Case "e": AsignarFonemaBaseEs = 2: Exit Function
'        Case "i": AsignarFonemaBaseEs = 4: Exit Function
'        Case "o": AsignarFonemaBaseEs = 5: Exit Function
'        Case "u": AsignarFonemaBaseEs = 7: Exit Function
'    End Select
'
'    '===========================================================
'    ' OCLUSIVAS
'    '===========================================================
'    Select Case grafema
'        Case "p": AsignarFonemaBaseEs = 30: Exit Function
'        Case "b", "v": AsignarFonemaBaseEs = 31: Exit Function
'        Case "t": AsignarFonemaBaseEs = 32: Exit Function
'        Case "d": AsignarFonemaBaseEs = 33: Exit Function
'    End Select
'
'    '===========================================================
'    ' K
'    '===========================================================
'    If grafema = "k" Then AsignarFonemaBaseEs = 34: Exit Function
'
'    '===========================================================
'    ' C
'    '===========================================================
'    If grafema = "c" Then
'        If sig = "e" Or sig = "i" Then AsignarFonemaBaseEs = 46 Else AsignarFonemaBaseEs = 34
'        Exit Function
'    End If
'
'    '===========================================================
'    ' QU
'    '===========================================================
'    If grafema = "qu" Then AsignarFonemaBaseEs = 34: Exit Function
'
'    '===========================================================
'    ' G
'    '===========================================================
'    If grafema = "g" Then
'        If sig = "e" Or sig = "i" Then AsignarFonemaBaseEs = 48 Else AsignarFonemaBaseEs = 35
'        Exit Function
'    End If
'
'    '===========================================================
'    ' GU
'    '===========================================================
'    If grafema = "gu" Then AsignarFonemaBaseEs = 35: Exit Function
'
'    '===========================================================
'    ' FRICATIVAS
'    '===========================================================
'    Select Case grafema
'        Case "f": AsignarFonemaBaseEs = 40: Exit Function
'        Case "z": AsignarFonemaBaseEs = 46: Exit Function
'        Case "s": AsignarFonemaBaseEs = 42: Exit Function
'        Case "j": AsignarFonemaBaseEs = 48: Exit Function
'    End Select
'
'    '===========================================================
'    ' X (contextual)
'    '===========================================================
'    If grafema = "x" Then
'
'        ' 1) X inicial + vocal ? /s/
'        If ant = "" And sig Like "[aeiouáéíóúü]" Then
'            AsignarFonemaBaseEs = 42
'            Exit Function
'        End If
'
'        ' 2) Vocal + X + vocal ? /gz/
'        If ant Like "[aeiouáéíóúü]" And sig Like "[aeiouáéíóúü]" Then
'            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "35,43,"
'            AsignarFonemaBaseEs = 0
'            Exit Function
'        End If
'
'        ' 3) Vocal + X + consonante ? /ks/
'        If ant Like "[aeiouáéíóúü]" And sig Like "[bcdfghjklmnñpqrstvwxyz]" Then
'            ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "34,42,"
'            AsignarFonemaBaseEs = 0
'            Exit Function
'        End If
'
'        ' 4) Topónimos mexicanos ? /x/
'        If LCase$(ObjDTO.palabra) = "méxico" Or _
'           LCase$(ObjDTO.palabra) = "oaxaca" Or _
'           LCase$(ObjDTO.palabra) = "xola" Then
'            AsignarFonemaBaseEs = 48
'            Exit Function
'        End If
'
'        ' 5) Por defecto ? /ks/
'        ObjDTO.IdsFonemas = ObjDTO.IdsFonemas & "34,42,"
'        AsignarFonemaBaseEs = 0
'        Exit Function
'
'    End If
'
'    '===========================================================
'    ' NASALES
'    '===========================================================
'    Select Case grafema
'        Case "m": AsignarFonemaBaseEs = 36: Exit Function
'        Case "n": AsignarFonemaBaseEs = 37: Exit Function
'        Case "ñ": AsignarFonemaBaseEs = 38: Exit Function
'    End Select
'
'    '===========================================================
'    ' LATERALES
'    '===========================================================
'    If grafema = "l" Then AsignarFonemaBaseEs = 62: Exit Function
'    If grafema = "ll" Then AsignarFonemaBaseEs = 63: Exit Function
'
'    '===========================================================
'    ' VIBRANTES
'    '===========================================================
'    If grafema = "rr" Then AsignarFonemaBaseEs = 60: Exit Function
'
'    If grafema = "r" Then
'        If ant = "" Or Not ant Like "[aeiou]" Then AsignarFonemaBaseEs = 60 Else AsignarFonemaBaseEs = 59
'        Exit Function
'    End If
'
'    '===========================================================
'    ' AFRICADA
'    '===========================================================
'    If grafema = "ch" Then AsignarFonemaBaseEs = 57: Exit Function
'
'    '===========================================================
'    ' Y /?/
'    '===========================================================
'    If grafema = "y" Or grafema = "hi" Then AsignarFonemaBaseEs = 45: Exit Function
'
'    '===========================================================
'    ' MODOS DE SIBILANTES (C/Z/S)
'    '===========================================================
'    Select Case CFG.ModoSibilantes
'
'        Case 0, 1   ' Distinción
'            If grafema = "c" And (sig = "e" Or sig = "i") Then AsignarFonemaBaseEs = 46: Exit Function
'            If grafema = "z" Then AsignarFonemaBaseEs = 46: Exit Function
'            If grafema = "s" Then AsignarFonemaBaseEs = 42: Exit Function
'
'        Case 2       ' Seseo
'            If grafema = "c" And (sig = "e" Or sig = "i") Then AsignarFonemaBaseEs = 42: Exit Function
'            If grafema = "z" Then AsignarFonemaBaseEs = 42: Exit Function
'            If grafema = "s" Then AsignarFonemaBaseEs = 42: Exit Function
'
'        Case 3       ' Ceceo
'            If grafema = "c" And (sig = "e" Or sig = "i") Then AsignarFonemaBaseEs = 46: Exit Function
'            If grafema = "z" Then AsignarFonemaBaseEs = 46: Exit Function
'            If grafema = "s" Then AsignarFonemaBaseEs = 46: Exit Function
'
'    End Select
'
'    '===========================================================
'    ' MODOS LATERALES (LL/Y)
'    '===========================================================
'    Select Case CFG.ModoLateral
'
'        Case 0   ' Normal
'            If grafema = "ll" Then AsignarFonemaBaseEs = 63: Exit Function
'            If grafema = "y" Then AsignarFonemaBaseEs = 45: Exit Function
'
'        Case 1   ' Lleísmo
'            If grafema = "ll" Then AsignarFonemaBaseEs = 63: Exit Function
'            If grafema = "y" Then AsignarFonemaBaseEs = 63: Exit Function
'
'        Case 2   ' Yeísmo
'            If grafema = "ll" Then AsignarFonemaBaseEs = 45: Exit Function
'            If grafema = "y" Then AsignarFonemaBaseEs = 45: Exit Function
'
'        Case 3   ' Yeísmo rehilado
'            If grafema = "ll" Then AsignarFonemaBaseEs = 72: Exit Function
'            If grafema = "y" Then AsignarFonemaBaseEs = 72: Exit Function
'
'    End Select
'
'    '===========================================================
'    ' Si no se reconoce
'    '===========================================================
'    AsignarFonemaBaseEs = 255
'
'End Function


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

Private Function NormalizarVocales_ES(ByVal Texto As String) As String

    Texto = Replace(Texto, "á", "a")
    Texto = Replace(Texto, "à", "a")
    Texto = Replace(Texto, "ä", "a")
    Texto = Replace(Texto, "â", "a")

    Texto = Replace(Texto, "é", "e")
    Texto = Replace(Texto, "è", "e")
    Texto = Replace(Texto, "ë", "e")
    Texto = Replace(Texto, "ê", "e")

    Texto = Replace(Texto, "í", "i")
    Texto = Replace(Texto, "ì", "i")
    Texto = Replace(Texto, "ï", "i")
    Texto = Replace(Texto, "î", "i")

    Texto = Replace(Texto, "ó", "o")
    Texto = Replace(Texto, "ò", "o")
    Texto = Replace(Texto, "ö", "o")
    Texto = Replace(Texto, "ô", "o")

    Texto = Replace(Texto, "ú", "u")
    Texto = Replace(Texto, "ù", "u")
    Texto = Replace(Texto, "û", "u")

    ' ü se mantiene
    NormalizarVocales_ES = Texto

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

            If Trim$(silaba) = "" Then GoTo SiguienteSilaba

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

SiguienteSilaba:
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


'-------------------------------------------------------------
'-------------------------------------------------------------

Private Sub CargaCombos(frm As Form)

    With frm

        ' Limpiar combos
        .cmbVocales.Value = ""
        .cmbTrataH.Value = ""
        .cmbGenero.Value = ""
        .cmbNumero.Value = ""
        ' ... los que correspondan

        ' Cargar valores según idioma
        ' Ejemplo para Castellano
        .cmbVocales.Clear
        .cmbVocales.AddItem "0;Ninguna"
        .cmbVocales.AddItem "1;Acentos Castellanos"
        .cmbVocales.AddItem "2;Acentos Graves y Agudos"
        .cmbVocales.AddItem "3;Todos los acentos y marcadores"

        ' ... resto de combos según idioma

        ' Cargar el ListBox con SQL dinámico
        Dim strSQL As String
        strSQL = CreaSQL("Castellano")
        .lstDigrafos.RowSource = strSQL

    End With

End Sub

