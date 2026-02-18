Attribute VB_Name = "modMotorES_Main"

' ============================
'  modMotorES_Main
'  Motor fonético español
'  (flujo principal, DTO-céntrico)
' ============================

Option Compare Database
Option Explicit

' Estado interno del motor (privado)
Private IndiceSilabaActual As Long
Private EsTonicaActual As Boolean
Private GrafAnterior As String
Private GrafActual As String
Private GrafSiguiente As String

' DTO
Private objDTO As clsMotorFonetico

' ============================================================
'   ENTRADA PRINCIPAL DEL MOTOR
' ============================================================
Public Function EntradaMotor_ES(texto As String) As String

    Set objDTO = New clsMotorFonetico
    objDTO.TextoOriginal = texto

    Call NormalizarEntrada
    'Call MF_DebugDTO("NormalizarEntrada")
    
    Call SilabearAuto
    'Call MF_DebugDTO("SilabearAuto")
    
    Call DetectarTonica
    'Call MF_DebugDTO("DetectarTonica")
    
    Call MarcarSilabasTonicas
    'Call MF_DebugDTO("MarcarSilabasTonicas")
    
    Call RevisionSilabeo
'    Call MF_DebugDTO("RevisionSilabeo")
'    Call ReconstruirSilabasFinales

    Call ConvertirSilabasAFonemas
'    Call MF_DebugDTO("ConvertirSilabasAFonemas")
    
    EntradaMotor_ES = objDTO.TextoFinal
    Call MF_DebugDTO("Proceso finalizado")
    
End Function

' ============================================================
'   NORMALIZACIÓN
' ============================================================
Private Sub NormalizarEntrada()

    Dim s As String

    s = objDTO.TextoOriginal

    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop

    s = Replace(s, vbTab, " ")
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")

    s = Replace(s, "–", "-")
    s = Replace(s, "—", "-")
    s = Replace(s, "“", """")
    s = Replace(s, "”", """")

    s = Replace(s, " -", "-")
    s = Replace(s, "- ", "-")

    s = LCase$(Trim$(s))

    objDTO.TextoNormalizado = s

End Sub


' ============================================================
'   SILABEO AUTOMÁTICO
' ============================================================
Private Sub SilabearAuto()

    Dim texto As String
    Dim col As New Collection
    Dim i As Long, ini As Long
    Dim c1 As String, c2 As String, c3 As String, c4 As String
    Dim par As String
    
    Dim arr() As String
    Dim strSilabas As String
    
    texto = objDTO.TextoNormalizado
    texto = Trim$(texto)

    If Len(texto) = 0 Then
        objDTO.SilabasAuto = ""
        Exit Sub
    End If

    ini = 1

    For i = 2 To Len(texto)

        c1 = Mid$(texto, i - 1, 1)
        c2 = Mid$(texto, i, 1)
        par = c1 & c2

        ' 0. ESPACIOS
        If c1 = " " Then
            ' Cierra la sílaba anterior, pero NO añade sílaba vacía aquí
            If i - 2 >= ini Then col.Add Array(ini, i - 2)
            ini = i
            GoTo siguiente
        End If
        
        If c2 = " " Then
            ' Cierra la sílaba anterior
            col.Add Array(ini, i - 1)
            ' Añade UNA sola sílaba vacía
            col.Add Array(i, i)
            ini = i + 1
            GoTo siguiente
        End If

        ' 1. DÍGRAFOS INSEPARABLES
        If par = "ch" Or par = "ll" Or par = "rr" Then GoTo siguiente

        ' 2. GRUPOS CONSONÁNTICOS INSEPARABLES
        If EsConsonante_ES(c1) And EsConsonante_ES(c2) Then
            If EsGrupoInseparable_ES(par) Then GoTo siguiente
        End If

        ' 3. REGLAS VOCÁLICAS (VV / VVV)

        ' TRIPTONGO
        If i < Len(texto) - 1 Then
            c3 = Mid$(texto, i + 1, 1)
            If EsTriptongo(c1, c2, c3) Then GoTo siguiente
        End If

        ' HIATO POR TILDE
        If EsVocal_ES(c1) And EsVocal_ES(c2) Then
            If (c1 = "í" Or c1 = "ú") Or (c2 = "í" Or c2 = "ú") Then
                col.Add Array(ini, i - 1)
                ini = i
                GoTo siguiente
            End If
        End If

        ' HIATO FUERTE + FUERTE
        If EsHiatoFuerteFuerte(c1, c2) Then
            col.Add Array(ini, i - 1)
            ini = i
            GoTo siguiente
        End If

        ' DIPTONGO
        If EsDiptongo(c1, c2) Then GoTo siguiente

        ' VV --> separar
        If EsVocal_ES(c1) And EsVocal_ES(c2) Then
            col.Add Array(ini, i - 1)
            ini = i
            GoTo siguiente
        End If

        ' *** REGLA ESPECIAL: V + RR ***
        If EsVocal_ES(c1) And c2 = "r" Then
            If i < Len(texto) Then
                c3 = Mid$(texto, i + 1, 1)
                If c3 = "r" Then
                    col.Add Array(ini, i - 1)
                    ini = i
                    GoTo siguiente
                End If
            End If
        End If

        ' 4. CCV ? C | CV
        If EsConsonante_ES(c1) And EsConsonante_ES(c2) Then
            If i < Len(texto) Then
                c3 = Mid$(texto, i + 1, 1)
                If EsVocal_ES(c3) Then
                    If Not EsGrupoInseparable_ES(par) Then
                        col.Add Array(ini, i - 1)
                        ini = i
                        GoTo siguiente
                    End If
                End If
            End If
        End If

        ' 5. VCV ? V | CV
        If EsVocal_ES(c1) And EsConsonante_ES(c2) Then
            If i < Len(texto) Then
                c3 = Mid$(texto, i + 1, 1)
                If EsVocal_ES(c3) Then
                    col.Add Array(ini, i - 1)
                    ini = i
                    GoTo siguiente
                End If
            End If
        End If

        ' 6. EXCEPCIÓN A-H-U
        If EsVocal_ES(c1) And EsConsonante_ES(c2) Then
            If i < Len(texto) Then
                c3 = Mid$(texto, i + 1, 1)
                If EsVocal_ES(c3) Then
                    If c1 = "a" And c2 = "h" And c3 = "u" Then
                        If i + 1 < Len(texto) Then
                            c4 = Mid$(texto, i + 2, 1)
                            If EsConsonante_ES(c4) Or c4 = "y" Then GoTo siguiente
                        End If
                    End If
                    col.Add Array(ini, i - 1)
                    ini = i
                    GoTo siguiente
                End If
            End If
        End If

siguiente:
    Next i

    If ini <= Len(texto) Then
        col.Add Array(ini, Len(texto))
    End If

    
    ReDim arr(1 To col.Count)
    
    For i = 1 To col.Count
        arr(i) = Mid$(texto, col(i)(0), col(i)(1) - col(i)(0) + 1)
    Next i
    
    ' Limpio los espacios dobles
    strSilabas = Join(arr, "-")
    
    While InStr(strSilabas, "- - -")
        strSilabas = Replace(strSilabas, "- - -", "- -")
    Wend
    
    objDTO.SilabasAuto = strSilabas
    
End Sub


' ============================================================
'   DETECCIÓN DE TÓNICA
' ============================================================
Private Sub DetectarTonica()

    Dim partes() As String
    Dim palabras As New Collection
    Dim palabraActual As New Collection
    Dim i As Long, j As Long
    Dim s As String
    Dim idx As Long
    Dim offset As Long
    Dim arrTonica() As String
    Dim countTonica As Long

    ' Si no hay sílabas, no hay nada que hacer
    If Len(objDTO.SilabasAuto) = 0 Then
        objDTO.SilabaTonica = ""
        Exit Sub
    End If

    ' 1. Dividir sílabas globales
    partes = Split(objDTO.SilabasAuto, "-")

    ' 2. Reconstruir palabras a partir de la sílaba vacía " "
    For i = LBound(partes) To UBound(partes)
        s = partes(i)

        If Trim$(s) = "" Then
            ' Fin de palabra
            If palabraActual.Count > 0 Then
                palabras.Add palabraActual
                Set palabraActual = New Collection
            End If
        Else
            palabraActual.Add s
        End If
    Next i

    ' Añadir la última palabra si existe
    If palabraActual.Count > 0 Then palabras.Add palabraActual

    ' 3. Preparar array dinámico para tónicas
    ReDim arrTonica(0 To 0)
    countTonica = 0
    offset = 0

    ' 4. Procesar palabra por palabra
    For i = 1 To palabras.Count

        ' Detectar tónica de esta palabra (tu función existente)
        idx = DetectarTonicaDeUnaPalabra_DesdeSilabas(palabras(i))

        If idx > 0 Then
            countTonica = countTonica + 1
            ReDim Preserve arrTonica(0 To countTonica - 1)
            arrTonica(countTonica - 1) = CStr(offset + idx)
        End If

        ' Avanzar offset global (+1 por la sílaba vacía entre palabras)
        offset = offset + palabras(i).Count + 1
    Next i

    ' 5. Guardar resultado final
    If countTonica = 0 Then
        objDTO.SilabaTonica = ""
    Else
        objDTO.SilabaTonica = Join(arrTonica, ",")
    End If

End Sub


Private Sub MarcarSilabasTonicas()

Dim marcadas As String
    
    marcadas = MarcarTonicas(objDTO.SilabasAuto, objDTO.SilabaTonica)
    objDTO.SilabasAuto = marcadas

End Sub

' ============================================================
'   REVISIÓN MANUAL
' ============================================================
Private Sub RevisionSilabeo()

    Dim texto As String
    Dim s As String
    Dim partes() As String
    Dim resultado As String
    Dim idxTonica As New Collection
    Dim raw As String
    Dim i As Long

    'texto = objDTO.TextoNormalizado
    texto = objDTO.TextoOriginal
    
    ' SilabasAuto ya es un string con guiones
    s = objDTO.SilabasAuto

    ' Abrir formulario de revisión
    s = RevisarSilabas_EnFormulario(texto, s)

    ' Si el usuario cancela, conservar lo anterior
    If s = "" Then
        objDTO.SilabasFinal = objDTO.SilabasAuto
        objDTO.SilabaTonica = objDTO.SilabaTonica
        Exit Sub
    End If

    ' Dividir sílabas revisadas
    partes = Split(s, "-")

    ' Procesar cada sílaba
    For i = LBound(partes) To UBound(partes)

        raw = partes(i)

        If EsSilabaMarcada(raw) Then
            idxTonica.Add i + 1
            partes(i) = LimpiarSilabaMarcada(raw)
        Else
            partes(i) = raw
        End If

    Next i

    ' Reconstruir SilabasFinal como string
    resultado = Join(partes, "-")
    objDTO.SilabasFinal = resultado

    ' Reconstruir SilabaTonica como string
    If idxTonica.Count > 0 Then
        Dim arr() As String
        ReDim arr(0 To idxTonica.Count - 1)

        For i = 1 To idxTonica.Count
            arr(i - 1) = CStr(idxTonica(i))
        Next i

        objDTO.SilabaTonica = Join(arr, ",")
    End If

End Sub


' ============================================================
'   RECONSTRUCCIÓN FINAL DE SÍLABAS
' ============================================================
Private Sub ReconstruirSilabasFinales()
    ' Esta función ahora no hace nada porque la reconstrucción
    ' ya se hace en RevisionSilabeo.
    ' La dejamos por si en el futuro quieres añadir lógica adicional.
End Sub


' ============================================================
'   CONVERSIÓN SÍLABAS ? FONEMAS
' ============================================================
Private Sub ConvertirSilabasAFonemas()

    Dim arrSilabas() As String
    Dim arrTonica() As String
    Dim i As Long, j As Long
    Dim esTonica As Boolean
    Dim sil As String

    Dim listaSilabas As New Collection
    Dim silabaFonemas As Collection
    Dim resultado As Collection
    Dim f As Variant

    arrSilabas = Split(objDTO.SilabasFinal, "-")
    arrTonica = Split(objDTO.SilabaTonica, ",")

    ' ============================================
    '   1. Procesar cada sílaba
    ' ============================================
    For i = 0 To UBound(arrSilabas)

        sil = arrSilabas(i)
        esTonica = False

        ' ¿Es tónica?
        For j = 0 To UBound(arrTonica)
            If Trim$(arrTonica(j)) <> "" Then
                If CLng(arrTonica(j)) = i + 1 Then
                    esTonica = True
                    Exit For
                End If
            End If
        Next j

        ' Nueva colección para esta sílaba
        Set silabaFonemas = New Collection

        ' 61 si es tónica
        If esTonica Then silabaFonemas.Add 61

        ' Espacio ? 0
        If Trim$(sil) = "" Then
            silabaFonemas.Add 0
        Else
            ' ============================================
            '   LLAMADA AL MOTOR FONÉTICO MODULAR
            ' ============================================
            Set resultado = Conv_Silaba(sil)

            ' Añadir todos los fonemas devueltos
            For Each f In resultado
                silabaFonemas.Add f
            Next f
        End If

        listaSilabas.Add silabaFonemas

    Next i

    ' ============================================
    '   2. Construir salida final
    ' ============================================
    Dim salida As String
    Dim parte As String

    salida = ""

    For Each silabaFonemas In listaSilabas

        parte = ""

        For Each f In silabaFonemas
            If parte = "" Then
                parte = CStr(f)
            Else
                parte = parte & "," & CStr(f)
            End If
        Next f

        If salida = "" Then
            salida = parte
        Else
            salida = salida & "-" & parte
        End If

    Next silabaFonemas

    objDTO.TextoFinal = salida

End Sub

'======================================================================================
' Conversor Fonético
'======================================================================================

' ============================================
'   MÓDULO PRINCIPAL: Conv_Silaba
'   Orquesta la conversión grafema ? fonema
'   Devuelve:
'       - Collection de fonemas
' ============================================

Private Function Conv_Silaba(silaba As String) As Collection

    Dim col As New Collection
    Dim i As Long
    Dim g As String, g2 As String
    Dim anterior As String, siguiente As String
    Dim resultado As Variant
    Dim f As Variant

    silaba = NormalizaVocales(LCase$(silaba))

    For i = 1 To Len(silaba)

        g = Mid$(silaba, i, 1)
        
        'anterior = IIf(i > 1, Mid$(silaba, i - 1, 1), "")
        anterior = ""
        If i > 1 Then
            anterior = Mid$(silaba, i - 1, 1)
        End If
        
        'siguiente = IIf(i < Len(silaba), Mid$(silaba, i + 1, 1), "")
        siguiente = ""
        If i < Len(silaba) Then
            siguiente = Mid$(silaba, i + 1, 1) ', "")
        End If
        
        ' ============================================
        '   1. DÍGRAFOS (2 grafemas ? 1 fonema)
        ' ============================================
        If i < Len(silaba) Then
            g2 = Mid$(silaba, i, 2)
            resultado = Conv_Digrafos(g2, siguiente)

            If Not IsNull(resultado) Then
                col.Add resultado
                i = i + 1   ' Consumimos 2 grafemas
                GoTo SiguienteGrafema
            End If
        End If

        ' ============================================
        '   2. VOCALES / DIPTONGOS / HIATOS
        ' ============================================
        If TypeOf Conv_Vocales(g, siguiente) Is Collection Then
            Set resultado = Conv_Vocales(g, siguiente)
        Else
            resultado = Conv_Vocales(g, siguiente)
        End If

        If Not IsNull(resultado) Then
            If TypeName(resultado) = "Collection" Then
                For Each f In resultado
                    col.Add f
                Next f
            Else
                col.Add resultado
            End If
            GoTo SiguienteGrafema
        End If

        ' ============================================
        '   3. REGLAS CONTEXTUALES
        ' ============================================
        resultado = Conv_Contexto(g, anterior, siguiente)

        If Not IsNull(resultado) Then
            col.Add resultado
            GoTo SiguienteGrafema
        End If

        ' ============================================
        '   4. MONÓGRAFOS
        ' ============================================
        resultado = Conv_Monografos(g)

        If Not IsNull(resultado) Then
            col.Add resultado
            GoTo SiguienteGrafema
        End If

        ' ============================================
        '   5. Si nada aplica ? placeholder
        ' ============================================
        col.Add 99

SiguienteGrafema:
    Next i

    Set Conv_Silaba = col

End Function

' ============================================
'   MÓDULO: Conv_Digrafos
'   Convierte dígrafos (2 grafemas ? 1 fonema)
'   Devuelve:
'       - ID fonema (Long)
'       - Null si no aplica
' ============================================
Private Function Conv_Digrafos(g2 As String, siguiente As String) As Variant

    g2 = LCase$(g2)
    siguiente = LCase$(siguiente)

    Select Case g2

        ' ===== Dígrafos reales =====

        Case "ch"
            Conv_Digrafos = 41      ' /t??/

        Case "ll"
            Conv_Digrafos = 42      ' /?/ o /?/ según dialecto

        Case "rr"
            Conv_Digrafos = 43      ' vibrante múltiple


        ' ===== Dígrafos ortográficos =====

        Case "gu"
            If siguiente = "e" Or siguiente = "i" Then
                Conv_Digrafos = 24  ' /g/
            Else
                Conv_Digrafos = Null
            End If

        Case "qu"
            If siguiente = "e" Or siguiente = "i" Then
                Conv_Digrafos = 31  ' /k/
            Else
                Conv_Digrafos = Null
            End If


        ' ===== No reconocido =====
        Case Else
            Conv_Digrafos = Null

    End Select

End Function

' ============================================
'   MÓDULO: Conv_Vocales
'   Vocales, diptongos, hiatos, semivocales
'   Devuelve:
'       - Collection (1 o más fonemas)
'       - Null si no aplica
' ============================================

Private Function Conv_Vocales(g As String, siguiente As String) As Variant

    g = LCase$(g)
    siguiente = LCase$(siguiente)

    Dim col As Collection
    Set col = New Collection

    ' ============================
    '   1. Diptongos crecientes
    ' ============================
    ' i/u + a/e/o
    If (g = "i" Or g = "u") And (siguiente = "a" Or siguiente = "e" Or siguiente = "o") Then
        col.Add 14   ' i ? semivocal
        col.Add VocalID(siguiente)
        Set Conv_Vocales = col
        Exit Function
    End If

    ' ============================
    '   2. Diptongos decrecientes
    ' ============================
    ' a/e/o + i/u
    If (g = "a" Or g = "e" Or g = "o") And (siguiente = "i" Or siguiente = "u") Then
        col.Add VocalID(g)
        col.Add 14   ' i/u ? semivocal
        Set Conv_Vocales = col
        Exit Function
    End If

    ' ============================
    '   3. Hiatos obligatorios
    ' ============================
    ' Vocal abierta + vocal abierta
    If (g = "a" Or g = "e" Or g = "o") And (siguiente = "a" Or siguiente = "e" Or siguiente = "o") Then
        col.Add VocalID(g)
        Set Conv_Vocales = col
        Exit Function
    End If

    ' ============================
    '   4. Vocal simple
    ' ============================
    If g Like "[aeiou]" Then
        col.Add VocalID(g)
        Set Conv_Vocales = col
        Exit Function
    End If

    ' ============================
    '   5. No aplica
    ' ============================
    Conv_Vocales = Null

End Function


' ============================================
'   Función auxiliar: VocalID
' ============================================
Private Function VocalID(v As String) As Long
    Select Case v
        Case "a": VocalID = 12
        Case "e": VocalID = 13
        Case "i": VocalID = 14
        Case "o": VocalID = 15
        Case "u": VocalID = 16
    End Select
End Function


' ============================================
'   MÓDULO: Conv_Contexto
'   Reglas contextuales (c, g, y, x...)
'   Devuelve:
'       - ID fonema (Long)
'       - Null si no aplica
' ============================================

Private Function Conv_Contexto(g As String, anterior As String, siguiente As String) As Variant

    g = LCase$(g)
    anterior = LCase$(anterior)
    siguiente = LCase$(siguiente)

    ' ============================
    '   1. C ? /k/ o /?/
    ' ============================
    If g = "c" Then
        If siguiente = "e" Or siguiente = "i" Then
            Conv_Contexto = 33   ' /?/ (o /s/ si seseo)
        Else
            Conv_Contexto = 31   ' /k/
        End If
        Exit Function
    End If

    ' ============================
    '   2. G ? /g/ o /x/
    ' ============================
    If g = "j" Then
        Conv_Contexto = 34    ' /x/
        Exit Function
    
    ElseIf g = "g" Then
    
        ' --- g + ü ? /gw/ ---
        If siguiente = "ü" Then
            Conv_Contexto = 57   ' /gw/
            Exit Function
        End If
    
        ' --- g + e/i ? /x/ ---
        If siguiente = "e" Or siguiente = "i" Then
            Conv_Contexto = 34   ' /x/
            Exit Function
        End If
    
        ' --- g + a/o/u ? /g/ ---
        Conv_Contexto = 24       ' /g/
        Exit Function
    
    End If

    
'    If g = "j" Then
'        Conv_Contexto = 34    ' /x/
'        Exit Function
'    ElseIf g = "g" Then
'        If siguiente = "e" Or siguiente = "i" Then
'            Conv_Contexto = 34   ' /x/
'        Else
'            Conv_Contexto = 24   ' /g/
'        End If
'        Exit Function
'    End If

    ' ============================
    '   3. Y ? /i/ o /?/
    ' ============================
    If g = "y" Then
        If siguiente = "" Then
            Conv_Contexto = 14   ' /i/ final
        Else
            Conv_Contexto = 35   ' /?/
        End If
        Exit Function
    End If

    ' ============================
    '   4. X ? /ks/ o /x/
    ' ============================
    If g = "x" Then
        If siguiente Like "[aeiou]" Then
            Conv_Contexto = 36   ' /ks/
        Else
            Conv_Contexto = 32   ' /x/
        End If
        Exit Function
    End If

    ' ============================
    '   5. S ? /s/
    ' ============================
    If g = "s" Then
        Conv_Contexto = 29
        Exit Function
    End If

    ' ============================
    '   6. No aplica
    ' ============================
    Conv_Contexto = Null

End Function

' ============================================
'   MÓDULO: Conv_Monografos
'   Convierte monógrafos (1 grafema ? 1 fonema)
'   Devuelve:
'       - ID fonema (Long)
'       - Null si no aplica
' ============================================

Private Function Conv_Monografos(g As String) As Variant

    g = LCase$(g)

    Select Case g

        ' ===== Vocales =====
        Case "a": Conv_Monografos = 12
        Case "e": Conv_Monografos = 13
        Case "i": Conv_Monografos = 14
        Case "o": Conv_Monografos = 15
        Case "u": Conv_Monografos = 16

        ' ===== Consonantes =====
        Case "m": Conv_Monografos = 21
        Case "n": Conv_Monografos = 22
        Case "p": Conv_Monografos = 23
        Case "b": Conv_Monografos = 24
        Case "d": Conv_Monografos = 25
        Case "f": Conv_Monografos = 26
        Case "l": Conv_Monografos = 27
        Case "r": Conv_Monografos = 28
        Case "s": Conv_Monografos = 29
        Case "t": Conv_Monografos = 30
        Case "k": Conv_Monografos = 31
        Case "x": Conv_Monografos = 32
        Case "z": Conv_Monografos = 33

        ' ===== No reconocido =====
        Case Else
            Conv_Monografos = Null

    End Select

End Function

Private Function NormalizaVocales(ByVal texto As String) As String

    Dim t As String
    t = texto

    Select Case texto
        Case "á", "à", "ä", "â"
            t = "a"

        Case "é", "è", "ë", "ê"
            t = "e"

        Case "í", "ì", "ï", "î"
            t = "i"

        Case "ó", "ò", "ö", "ô"
            t = "o"

        Case "ú", "ù", "û"   ' ü se mantiene
            t = "u"
    End Select

    NormalizaVocales = t

End Function

'Private Sub ConvertirSilabasAFonemas()
'
'    Dim i As Long, j As Long
'    Dim strFinal As String
'    Dim esTonica As Boolean
'    Dim sil As String
'    Dim arrFon() As Byte
'    Dim f As Variant
'
'    Dim arrSilabas() As String
'    Dim arrTonica() As String
'
'    arrSilabas = Split(objDTO.SilabasFinal, "-")
'    arrTonica = Split(objDTO.SilabaTonica, ",")
'
'    If UBound(arrSilabas) < 0 Then Exit Sub
'
'    strFinal = ""
'
'    For i = 0 To UBound(arrSilabas)
'
'        sil = arrSilabas(i)
'
'        ' Detectar si es tónica
'        esTonica = False
'        If UBound(arrTonica) >= 0 Then
'            For j = 0 To UBound(arrTonica)
'                If CByte(arrTonica(j)) = i + 1 Then
'                    esTonica = True
'                    Exit For
'                End If
'            Next j
'        End If
'
'        EsTonicaActual = esTonica
'
'        If esTonica Then
'            strFinal = strFinal & "61, "
'        End If
'
'        If sil = " " Then
'            strFinal = strFinal & "0 - "
'            GoTo siguiente
'        End If
'
'        ' ?? Ahora pasamos la sílaba directamente
'        arrFon = ConvertirGrafemasDeSilabaAIdFonemas(sil)
'
'        For Each f In arrFon
'            strFinal = strFinal & CStr(f) & ", "
'        Next f
'
'        strFinal = strFinal & "- "
'
'siguiente:
'    Next i
'
'    objDTO.TextoFinal = Trim$(strFinal)
'
'End Sub



' ============================================================
'   CONVERSIÓN GRAFEMAS ? IDFONEMAS
' ============================================================
Private Function ConvertirGrafemasDeSilabaAIdFonemas(ByVal sil As String) As Byte()

'    Dim s As String
'    Dim i As Long
'    Dim graf As String
'    Dim fon As Byte
'    Dim arr() As Byte
'    Dim idx As Long
'
'    ' Normalizar sílaba
'    s = LCase$(sil)
'    s = MF_NormalizarVocales_ES(s)
'
'    ReDim arr(1 To 1)
'    idx = 1
'    i = 1
'
'    Do While i <= Len(s)
'
'        GrafAnterior = ""
'        GrafActual = ""
'        GrafSiguiente = ""
'
'        ' ============================
'        '   TRIGRAFEMAS
'        ' ============================
'        If i <= Len(s) - 2 Then
'            graf = Mid$(s, i, 3)
'
'            If graf = "güe" Or graf = "güi" Or _
'               graf = "gue" Or graf = "gui" Or _
'               graf = "que" Or graf = "qui" Then
'
'                If i > 1 Then GrafAnterior = Mid$(s, i - 1, 1)
'                If i < Len(s) - 2 Then GrafSiguiente = Mid$(s, i + 3, 1)
'                GrafActual = graf
'
'                fon = ReglasCastellano(graf, GrafAnterior, GrafSiguiente, EsTonicaActual)
'
'                If fon > 0 Then
'                    arr(idx) = fon
'                    idx = idx + 1
'                    ReDim Preserve arr(1 To idx)
'                    i = i + 3
'                    GoTo siguiente
'                End If
'            End If
'        End If
'
'        ' ============================
'        '   DÍGRAFOS
'        ' ============================
'        If i <= Len(s) - 1 Then
'            graf = Mid$(s, i, 2)
'
'            If graf = "ch" Or graf = "ll" Or graf = "rr" Or _
'               graf = "gu" Or graf = "qu" Or _
'               graf = "ai" Or graf = "ei" Or graf = "oi" Or graf = "ou" Or graf = "au" Then
'
'                If i > 1 Then GrafAnterior = Mid$(s, i - 1, 1)
'                If i < Len(s) - 1 Then GrafSiguiente = Mid$(s, i + 2, 1)
'                GrafActual = graf
'
'                fon = ReglasCastellano(graf, GrafAnterior, GrafSiguiente, EsTonicaActual)
'
'                If fon > 0 Then
'                    arr(idx) = fon
'                    idx = idx + 1
'                    ReDim Preserve arr(1 To idx)
'                    i = i + 2
'                    GoTo siguiente
'                End If
'            End If
'        End If
'
'        ' ============================
'        '   MONÓGRAFOS
'        ' ============================
'        graf = Mid$(s, i, 1)
'        GrafActual = graf
'
'        If i > 1 Then GrafAnterior = Mid$(s, i - 1, 1)
'        If i < Len(s) Then GrafSiguiente = Mid$(s, i + 1, 1)
'
'        fon = ReglasCastellano(graf, GrafAnterior, GrafSiguiente, EsTonicaActual)
'
'        If fon > 0 Then
'            arr(idx) = fon
'            idx = idx + 1
'            ReDim Preserve arr(1 To idx)
'        End If
'
'        i = i + 1
'
'siguiente:
'    Loop
'
'    ' Ajuste final
'    If idx > 1 Then
'        ReDim Preserve arr(1 To idx - 1)
'    Else
'        ReDim arr(1 To 0)
'    End If
'
'    ConvertirGrafemasDeSilabaAIdFonemas = arr

End Function

'=================================================================================================

'-----------------
' Auxiliares
'-----------------

Private Function ContarSilabasDePalabra(ByVal palabra As String, ByRef sils() As String, ByVal offset As Long) As Long
    Dim total As Long
    Dim i As Long
    Dim lenAcum As Long
    
    total = 0
    lenAcum = 0
    
    For i = offset To UBound(sils)
        lenAcum = lenAcum + Len(sils(i))
        
        If lenAcum <= Len(palabra) Then
            total = total + 1
        Else
            Exit For
        End If
    Next i
    
    ContarSilabasDePalabra = total
End Function

Private Function DetectarTonicaDeUnaPalabra(ByVal palabra As String, ByVal numSilabas As Long) As Long
    
    ' 1. Si hay tilde explícita ? esa sílaba es la tónica
    Dim i As Long
    Dim sils() As String
    sils = Split(objDTO.SilabasAuto, "-")
    
    For i = 0 To numSilabas - 1
        If TieneTilde(sils(i)) Then
            DetectarTonicaDeUnaPalabra = i + 1
            Exit Function
        End If
    Next i
    
    ' 2. Si no hay tilde ? aplicar reglas generales
    ' Palabras terminadas en vocal, n, s ? llana
    Dim ultima As String
    ultima = Right$(palabra, 1)
    
If ultima Like "[aeiousn]" Then
    If numSilabas > 1 Then
        DetectarTonicaDeUnaPalabra = numSilabas - 1
    Else
        DetectarTonicaDeUnaPalabra = 1
    End If
Else
    DetectarTonicaDeUnaPalabra = numSilabas
End If

End Function

Private Function DetectarTonicaDeUnaPalabra_DesdeSilabas(ByVal colSilabas As Collection) As Long
    Dim i As Long
    Dim sil As String

    ' 1. Buscar tilde explícita
    For i = 1 To colSilabas.Count
        sil = colSilabas(i)
        If TieneTilde(sil) Then
            DetectarTonicaDeUnaPalabra_DesdeSilabas = i
            Exit Function
        End If
    Next i

    ' 2. Si no hay tilde, aplicar reglas generales:
    '    - aguda si termina en vocal, n o s
    '    - llana en caso contrario

    Dim ultima As String
    ultima = colSilabas(colSilabas.Count)

    If TerminaEnVocalNoSNoN(ultima) Then
        ' Palabra llana ? tónica en penúltima
        If colSilabas.Count >= 2 Then
            DetectarTonicaDeUnaPalabra_DesdeSilabas = colSilabas.Count - 1
        Else
            DetectarTonicaDeUnaPalabra_DesdeSilabas = colSilabas.Count
        End If
    Else
        ' Palabra aguda ? tónica en última
        DetectarTonicaDeUnaPalabra_DesdeSilabas = colSilabas.Count
    End If
End Function

Private Function TieneTilde(ByVal silaba As String) As Boolean
    TieneTilde = (InStr(silaba, "á") > 0 Or _
                  InStr(silaba, "é") > 0 Or _
                  InStr(silaba, "í") > 0 Or _
                  InStr(silaba, "ó") > 0 Or _
                  InStr(silaba, "ú") > 0)
End Function

Private Function TerminaEnVocalNoSNoN(ByVal silaba As String) As Boolean
    Dim c As String
    c = Right$(silaba, 1)
    TerminaEnVocalNoSNoN = (c Like "[aeiouáéíóú]")
End Function


Public Function EsSilabaMarcada(ByVal s As String) As Boolean
    Dim t As String
    t = Trim$(s)
    EsSilabaMarcada = (Left$(t, 1) = "*" And Right$(t, 1) = "*")
End Function

Public Function LimpiarSilabaMarcada(ByVal s As String) As String
    Dim t As String
    t = Trim$(s)
    t = Mid$(t, 2, Len(t) - 2)   ' quitar los dos *
    LimpiarSilabaMarcada = Trim$(t)
End Function

' ============================================================
'   AUXILIARES FONÉTICAS (ES)
' ============================================================

Public Function EsVocal_ES(c As String) As Boolean
    EsVocal_ES = (InStr("aeiouáéíóú", c) > 0)
End Function

Public Function EsConsonante_ES(c As String) As Boolean
    EsConsonante_ES = (c <> " " And Not EsVocal_ES(c))
End Function

Public Function EsVocalFuerte(c As String) As Boolean
    EsVocalFuerte = (InStr("aeoáéó", c) > 0)
End Function

Public Function EsVocalDebil(c As String) As Boolean
    EsVocalDebil = (InStr("iuíú", c) > 0)
End Function

Public Function EsDiptongo(c1 As String, c2 As String) As Boolean

    If EsVocalDebil(c1) And EsVocalDebil(c2) Then
        If c1 <> "í" And c1 <> "ú" And c2 <> "í" And c2 <> "ú" Then
            EsDiptongo = True
            Exit Function
        End If
    End If

    If EsVocalDebil(c1) And EsVocalFuerte(c2) Then
        If c1 <> "í" And c1 <> "ú" Then
            EsDiptongo = True
            Exit Function
        End If
    End If

    If EsVocalFuerte(c1) And EsVocalDebil(c2) Then
        If c2 <> "í" And c2 <> "ú" Then
            EsDiptongo = True
            Exit Function
        End If
    End If

End Function

Public Function EsTriptongo(c1 As String, c2 As String, c3 As String) As Boolean
    If EsVocalDebil(c1) And EsVocalFuerte(c2) And EsVocalDebil(c3) Then
        If c1 <> "í" And c1 <> "ú" And c3 <> "í" And c3 <> "ú" Then
            EsTriptongo = True
        End If
    End If
End Function

Public Function EsHiatoFuerteFuerte(c1 As String, c2 As String) As Boolean
    If EsVocalFuerte(c1) And EsVocalFuerte(c2) Then
        EsHiatoFuerteFuerte = True
    End If
End Function

Public Function EsGrupoInseparable_ES(par As String) As Boolean

    Dim g As Variant
    Dim lista As Variant

    lista = Array("br", "bl", "cr", "cl", "dr", "fr", "gr", "gl", "pr", "pl", "tr")

    For Each g In lista
        If par = g Then
            EsGrupoInseparable_ES = True
            Exit Function
        End If
    Next g

End Function

' ============================================
'   Rutina auxiliar de diagnóstico del motor
'   Imprime el estado completo del DTO
' ============================================
Public Sub MF_DebugDTO(Proc As String)

    If objDTO Is Nothing Then
        Debug.Print "DTO no inicializado."
        Exit Sub
    End If

    Debug.Print
    Debug.Print "==============================="
    Debug.Print "   ESTADO DEL MOTOR"
    Debug.Print "==============================="
    
    Debug.Print
    Debug.Print "-------------------------------"
    Debug.Print " Proc.: "; Proc
    Debug.Print "-------------------------------"
    Debug.Print
    
    Debug.Print "Texto original:        "; objDTO.TextoOriginal
    Debug.Print "Texto normalizado:     "; objDTO.TextoNormalizado
    Debug.Print "SilabasAuto:           "; objDTO.SilabasAuto
    Debug.Print "SilabasFinal:          "; objDTO.SilabasFinal
    Debug.Print "SilabaTonica:          "; objDTO.SilabaTonica
    Debug.Print "TextoFinal (fonemas):  "; objDTO.TextoFinal

    Debug.Print "-------------------------------"
    Debug.Print "   Detalles internos"
    Debug.Print "-------------------------------"

    Debug.Print "Num sílabas auto:      "; CountItems(objDTO.SilabasAuto, "-") + 1
    Debug.Print "Num sílabas final:     "; CountItems(objDTO.SilabasFinal, "-") + 1

    Debug.Print "==============================="
    Debug.Print 'vbCrLf

    'Stop
    
End Sub

' Contador auxiliar para separar elementos
Private Function CountItems(ByVal s As String, ByVal sep As String) As Long
    If Len(Trim$(s)) = 0 Then
        CountItems = 0
    Else
        CountItems = UBound(Split(s, sep))
    End If
End Function


'Private Sub ConvertirSilabasAFonemas()
'
'    Dim arrSilabas() As String
'    Dim arrTonica() As String
'    Dim i As Long, j As Long
'    Dim esTonica As Boolean
'    Dim sil As String
'
'    Dim listaSilabas As New Collection      ' colección de colecciones
'    Dim silabaFonemas As Collection        ' colección de fonemas de UNA sílaba
'    Dim f As Variant
'
'    arrSilabas = Split(objDTO.SilabasFinal, "-")
'    arrTonica = Split(objDTO.SilabaTonica, ",")
'
'    ' === 1. Construir fonemas por sílaba ===
'    For i = 0 To UBound(arrSilabas)
'
'        sil = arrSilabas(i)
'        esTonica = False
'
'        ' ¿Es tónica esta sílaba?
'        For j = 0 To UBound(arrTonica)
'            If Trim$(arrTonica(j)) <> "" Then
'                If CLng(arrTonica(j)) = i + 1 Then
'                    esTonica = True
'                    Exit For
'                End If
'            End If
'        Next j
'
'        ' NUEVA colección por cada sílaba
'        Set silabaFonemas = New Collection
'
'        ' 61 si es tónica
'        If esTonica Then silabaFonemas.Add 61
'
'        ' Espacio ? 0, resto ? placeholder 99
'        If Trim$(sil) = "" Then
'            silabaFonemas.Add 0
'        Else
'            silabaFonemas.Add 99   ' aquí luego irán los fonemas reales
'        End If
'
'        ' Añadir esta sílaba (su colección) a la lista global
'        listaSilabas.Add silabaFonemas
'
'    Next i
'
'    ' === 2. Ensamblar salida final ===
'    Dim salida As String
'    Dim parte As String
'
'    salida = ""
'
'    For Each silabaFonemas In listaSilabas
'
'        parte = ""
'
'        For Each f In silabaFonemas
'            If parte = "" Then
'                parte = CStr(f)
'            Else
'                parte = parte & "," & CStr(f)
'            End If
'        Next f
'
'        If salida = "" Then
'            salida = parte
'        Else
'            salida = salida & "-" & parte
'        End If
'
'    Next silabaFonemas
'
'    objDTO.TextoFinal = salida
'
'End Sub


