Attribute VB_Name = "Módulo4"

Option Compare Database
Option Explicit

' ================================================================
'   Motor Silábico 2.1 — Versión limpia y autocontenida
'   - Comparación binaria (distinción entre vocales acentuadas)
'   - Sin dependencias externas
'   - Preparado para integrarse con clsDTO_Motor
' ================================================================

Public ObjDTO As clsDTO_Motor

Public Function Entrada_Motor_ES(Texto As String) As String

    Set ObjDTO = New clsDTO_Motor
    
    strDebug = ""

    ' 1) Asignamos el texto recibido y
    '    Normalización (dentro del DTO)
    ObjDTO.TextoOriginal = Texto
    ObjDTO.NormalizaEntrada

    ' 2) Silabeo automático
    Call Silabear

'    Call MF_DebugDTO("Silabear")

    ' 3) Devolver resultado
    Entrada_Motor_ES = ObjDTO.SilabasAuto

End Function

'' ================================================================
''   FUNCIÓN PRINCIPAL PARA EL DTO
'' ================================================================
'Public Sub ProcesarSilabas(ByRef dto As clsDTO_Motor)
'
'    dto.NormalizaEntrada
'    dto.SilabasAuto = SilabearTexto(dto.TextoNormalizado)
'
'End Sub


' ================================================================
'   SILABEAR TEXTO COMPLETO
' ================================================================
''Public Function SilabearTexto(ByVal Texto As String) As String
'Public Function Silabear() As String
'
'    Dim palabras() As String
'    Dim i As Long
'    Dim res As String
'
'    palabras = Split(Trim$(ObjDTO.TextoNormalizado), " ")
'
'    For i = LBound(palabras) To UBound(palabras)
'        If palabras(i) <> "" Then
'            If res <> "" Then res = res & " "
'            res = res & SilabearPalabra(palabras(i))
'        End If
'    Next i
'
'    SilabearTexto = res
'End Function

Public Function Silabear() As String

    Dim palabras() As String
    Dim i As Long
    Dim res As String
    
    palabras = Split(Trim$(ObjDTO.TextoNormalizado), " ")
    
    For i = LBound(palabras) To UBound(palabras)
        If palabras(i) <> "" Then
            If res <> "" Then res = res & " "
            res = res & SilabearPalabra(palabras(i))
        End If
    Next i

    ' Guardamos en el DTO
    ObjDTO.SilabasAuto = res

    ' Devolvemos el resultado
    Silabear = res

End Function


' ================================================================
'   SILABEAR UNA PALABRA
' ================================================================
Private Function SilabearPalabra(ByVal Texto As String) As String
    Dim i As Long
    Dim silaba As String
    Dim res As String
    Dim c As String

    Texto = LCase$(Texto)

    For i = 1 To Len(Texto)

        c = Mid$(Texto, i, 1)

        ' *** Evaluar la frontera ANTES de añadir c ***
        If DebeCortar(Texto, i - 1) And silaba <> "" Then
            If res <> "" Then res = res & " | "
            res = res & silaba
            silaba = ""
        End If

        silaba = silaba & c

    Next i

    If silaba <> "" Then
        If res <> "" Then res = res & " | "
        res = res & silaba
    End If

    SilabearPalabra = res
End Function

'Public Function SilabearPalabra(ByVal Texto As String) As String
'    Dim i As Long
'    Dim silaba As String
'    Dim res As String
'    Dim c As String
'
'    Texto = LCase$(Texto)
'
'    For i = 1 To Len(Texto)
'
'        c = Mid$(Texto, i, 1)
'        silaba = silaba & c
'
'        ' La frontera está DESPUÉS de i
'        If DebeCortar(Texto, i) Then
'            If res <> "" Then res = res & " | "
'            res = res & silaba
'            silaba = ""
'        End If
'
'    Next i
'
'    If silaba <> "" Then
'        If res <> "" Then res = res & " | "
'        res = res & silaba
'    End If
'
'    SilabearPalabra = res
'End Function

'Public Function SilabearPalabra(ByVal Texto As String) As String
'    Dim i As Long
'    Dim silaba As String
'    Dim res As String
'    Dim c As String
'
'    Texto = LCase$(Texto)
'
'    For i = 1 To Len(Texto)
'        c = Mid$(Texto, i, 1)
'        silaba = silaba & c
'
'        If DebeCortar(Texto, i) Then
'            If res <> "" Then res = res & " | "
'            res = res & silaba
'            silaba = ""
'        End If
'    Next i
'
'    If silaba <> "" Then
'        If res <> "" Then res = res & " | "
'        res = res & silaba
'    End If
'
'    SilabearPalabra = res
'End Function

'Public Function SilabearPalabra(ByVal Texto As String) As String
'    Dim i As Long
'    Dim silaba As String
'    Dim res As String
'
'    Dim c As String
'    Dim prev As String
'    Dim nextC As String
'    Dim next2 As String
'    Dim next3 As String
'    Dim ult As String
'
'    Texto = LCase$(Texto)
'
'    For i = 1 To Len(Texto)
'
'        c = Mid$(Texto, i, 1)
'
''        prev = IIf(i > 1, Mid$(Texto, i - 1, 1), "")
''        nextC = IIf(i < Len(Texto), Mid$(Texto, i + 1, 1), "")
''        next2 = IIf(i < Len(Texto) - 1, Mid$(Texto, i + 2, 1), "")
''        next3 = IIf(i < Len(Texto) - 2, Mid$(Texto, i + 3, 1), "")
'
'        ' Proteger prev
'        prev = ""
'        If i > 1 Then
'            prev = Mid$(Texto, i - 1, 1)
'        End If
'
'        ' Proteger nextC
'        nextC = ""
'        If i < Len(Texto) Then
'            nextC = Mid$(Texto, i + 1, 1)
'        End If
'
'        ' Proteger next2
'        next2 = ""
'        If i < Len(Texto) - 1 Then
'            next2 = Mid$(Texto, i + 2, 1)
'        End If
'
'        silaba = silaba & c
'
'        ' *** AQUÍ ESTABA EL ERROR ***
'        ' ult NO es prev ? es el último carácter REAL de la sílaba
'        ult = Right$(silaba, 1)
'
'        If DebeCortar(ult, c, nextC, next2, next3, silaba) Then
'            If res <> "" Then res = res & " | "
'            res = res & silaba
'            silaba = ""
'        End If
'
'    Next i
'
'    If silaba <> "" Then
'        If res <> "" Then res = res & " | "
'        res = res & silaba
'    End If
'
'    SilabearPalabra = res
'End Function

'Public Function SilabearPalabra(ByVal Texto As String) As String
'    Dim i As Long
'    Dim silaba As String
'    Dim res As String
'
'    Dim c As String
'    Dim prev As String
'    Dim nextC As String
'    Dim next2 As String
'    Dim next3 As String
'
'    Texto = LCase$(Texto)
'
'    For i = 1 To Len(Texto)
'
'        c = Mid$(Texto, i, 1)
''
''        prev = IIf(i > 1, Mid$(Texto, i - 1, 1), "")
''        nextC = IIf(i < Len(Texto), Mid$(Texto, i + 1, 1), "")
''        next2 = IIf(i < Len(Texto) - 1, Mid$(Texto, i + 2, 1), "")
''        next3 = IIf(i < Len(Texto) - 2, Mid$(Texto, i + 3, 1), "")
'
''        c = Mid$(Texto, i, 1)
'
'        ' Proteger prev
'        prev = ""
'        If i > 1 Then
'            prev = Mid$(Texto, i - 1, 1)
'        End If
'
'        ' Proteger nextC
'        nextC = ""
'        If i < Len(Texto) Then
'            nextC = Mid$(Texto, i + 1, 1)
'        End If
'
'        ' Proteger next2
'        next2 = ""
'        If i < Len(Texto) - 1 Then
'            next2 = Mid$(Texto, i + 2, 1)
'        End If
'
'        ' Proteger next3
'        next3 = ""
'        If i < Len(Texto) - 2 Then
'            next3 = Mid$(Texto, i + 3, 1)
'        End If
'
'
'
'        silaba = silaba & c
'
'        If DebeCortar(prev, c, nextC, next2, next3, silaba) Then
'            If res <> "" Then res = res & " | "
'            res = res & silaba
'            silaba = ""
'        End If
'
'    Next i
'
'    If silaba <> "" Then
'        If res <> "" Then res = res & " | "
'        res = res & silaba
'    End If
'
'    SilabearPalabra = res
'End Function


' ================================================================
'   DECISIÓN DE CORTE SILÁBICO
' ================================================================
Private Function DebeCortar(ByVal Texto As String, ByVal i As Long) As Boolean
    Dim L As Long
    Dim a1 As String, a2 As String, a3 As String
    Dim d1 As String, d2 As String, d3 As String

    L = Len(Texto)

    ' Frontera DESPUÉS de i
    If i >= 1 Then a1 = Mid$(Texto, i, 1)
    If i >= 2 Then a2 = Mid$(Texto, i - 1, 1)
    If i >= 3 Then a3 = Mid$(Texto, i - 2, 1)

    If i + 1 <= L Then d1 = Mid$(Texto, i + 1, 1)
    If i + 2 <= L Then d2 = Mid$(Texto, i + 2, 1)
    If i + 3 <= L Then d3 = Mid$(Texto, i + 3, 1)

    ' ----------------------------------------------------
    ' 0) Final de palabra
    ' ----------------------------------------------------
    If i = L Then
        DebeCortar = True
        Exit Function
    End If

    ' ----------------------------------------------------
    ' 1) V C V ? V - C V
    ' *** CORRECCIÓN CLAVE ***
    ' a2 = vocal, a1 = consonante, d1 = vocal
    ' ----------------------------------------------------
    If EsVocal(a2) And EsConsonante(a1) And EsVocal(d1) Then
        If Not EsAtaqueComplejo(a1, d1) Then
            DebeCortar = True
            Exit Function
        End If
    End If

    ' ----------------------------------------------------
    ' 2) V C C V ? V - CCV (ataque complejo)
    ' ----------------------------------------------------
    If EsVocal(a3) And EsConsonante(a2) And EsConsonante(a1) And EsVocal(d1) Then
        If EsAtaqueComplejo(a1, d1) Then
            DebeCortar = True
            Exit Function
        End If
    End If

    ' ----------------------------------------------------
    ' 3) C s V ? C - sV
    ' ----------------------------------------------------
    If a1 = "s" And EsConsonante(a2) And EsVocal(d1) Then
        DebeCortar = True
        Exit Function
    End If

    ' ----------------------------------------------------
    ' 4) C C ? C - C (salvo ataque complejo o coda compleja)
    ' ----------------------------------------------------
    If EsConsonante(a1) And EsConsonante(d1) Then

        If EsAtaqueComplejo(a1, d1) Then Exit Function

        If EsCodaCompleja(a1, d1) And EsConsonante(d2) Then Exit Function

        DebeCortar = True
        Exit Function
    End If

    ' ----------------------------------------------------
    ' 5) Hiato obligatorio V[-alta] V[-alta]
    ' ----------------------------------------------------
    If EsVocal(a1) And EsVocal(d1) Then
        If EsNoAlta(a1) And EsNoAlta(d1) Then
            DebeCortar = True
            Exit Function
        End If
    End If

    DebeCortar = False
End Function

'Private Function DebeCortar(ByVal Texto As String, ByVal i As Long) As Boolean
'    Dim L As Long
'    Dim a1 As String, a2 As String, a3 As String
'    Dim d1 As String, d2 As String, d3 As String
'
'    L = Len(Texto)
'
'    ' Frontera DESPUÉS de i
'    If i >= 1 Then a1 = Mid$(Texto, i, 1)
'    If i >= 2 Then a2 = Mid$(Texto, i - 1, 1)
'    If i >= 3 Then a3 = Mid$(Texto, i - 2, 1)
'
'    If i + 1 <= L Then d1 = Mid$(Texto, i + 1, 1)
'    If i + 2 <= L Then d2 = Mid$(Texto, i + 2, 1)
'    If i + 3 <= L Then d3 = Mid$(Texto, i + 3, 1)
'
'    ' ----------------------------------------------------
'    ' 0) Final de palabra
'    ' ----------------------------------------------------
'    If i = L Then
'        DebeCortar = True
'        Exit Function
'    End If
'
'    ' ----------------------------------------------------
'    ' 1) V C V ? V - C V
'    ' *** CORRECCIÓN CLAVE ***
'    ' a2 = vocal, a1 = consonante, d1 = vocal
'    ' ----------------------------------------------------
'    If EsVocal(a2) And EsConsonante(a1) And EsVocal(d1) Then
'        If Not EsAtaqueComplejo(a1, d1) Then
'            DebeCortar = True
'            Exit Function
'        End If
'    End If
'
'    ' ----------------------------------------------------
'    ' 2) V C C V ? V - CCV (ataque complejo)
'    ' ----------------------------------------------------
'    If EsVocal(a3) And EsConsonante(a2) And EsConsonante(a1) And EsVocal(d1) Then
'        If EsAtaqueComplejo(a1, d1) Then
'            DebeCortar = True
'            Exit Function
'        End If
'    End If
'
'    ' ----------------------------------------------------
'    ' 3) C s V ? C - sV
'    ' ----------------------------------------------------
'    If a1 = "s" And EsConsonante(a2) And EsVocal(d1) Then
'        DebeCortar = True
'        Exit Function
'    End If
'
'    ' ----------------------------------------------------
'    ' 4) C C ? C - C (salvo ataque complejo o coda compleja)
'    ' ----------------------------------------------------
'    If EsConsonante(a1) And EsConsonante(d1) Then
'
'        If EsAtaqueComplejo(a1, d1) Then Exit Function
'
'        If EsCodaCompleja(a1, d1) And EsConsonante(d2) Then Exit Function
'
'        DebeCortar = True
'        Exit Function
'    End If
'
'    ' ----------------------------------------------------
'    ' 5) Hiato obligatorio V[-alta] V[-alta]
'    ' ----------------------------------------------------
'    If EsVocal(a1) And EsVocal(d1) Then
'        If EsNoAlta(a1) And EsNoAlta(d1) Then
'            DebeCortar = True
'            Exit Function
'        End If
'    End If
'
'    DebeCortar = False
'End Function

'Private Function DebeCortar(ByVal Texto As String, ByVal i As Long) As Boolean
'    Dim L As Long
'    Dim a1 As String, a2 As String, a3 As String
'    Dim d1 As String, d2 As String, d3 As String
'
'    L = Len(Texto)
'
'    ' Frontera DESPUÉS de i
'    ' a1 = Texto(i)
'    ' a2 = Texto(i-1)
'    ' a3 = Texto(i-2)
'    ' d1 = Texto(i+1)
'    ' d2 = Texto(i+2)
'    ' d3 = Texto(i+3)
'
'    If i >= 1 Then a1 = Mid$(Texto, i, 1)
'    If i >= 2 Then a2 = Mid$(Texto, i - 1, 1)
'    If i >= 3 Then a3 = Mid$(Texto, i - 2, 1)
'
'    If i + 1 <= L Then d1 = Mid$(Texto, i + 1, 1)
'    If i + 2 <= L Then d2 = Mid$(Texto, i + 2, 1)
'    If i + 3 <= L Then d3 = Mid$(Texto, i + 3, 1)
'
'    ' ----------------------------------------------------
'    ' 0) Final de palabra
'    ' ----------------------------------------------------
'    If i = L Then
'        DebeCortar = True
'        Exit Function
'    End If
'
'    ' ----------------------------------------------------
'    ' 1) V C V ? V - C V
'    ' frontera entre a1 y d1
'    ' ----------------------------------------------------
''    If EsVocal(a1) And EsConsonante(d1) And EsVocal(d2) Then
''    If EsVocal(a2) And EsConsonante(a1) And EsVocal(d1) Then
'    If EsVocal(a2) And EsConsonante(a1) And EsVocal(d1) Then
'
'        ' salvo ataque complejo
'        If Not EsAtaqueComplejo(d1, d2) Then
'            DebeCortar = True
'            Exit Function
'        End If
'    End If
'
'    ' ----------------------------------------------------
'    ' 2) V C C V ? V - CCV (ataque complejo)
'    ' frontera entre a1 y d1
'    ' a1 = C, a2 = C, a3 = V, d1 = V
'    ' ----------------------------------------------------
'    If EsVocal(a3) And EsConsonante(a2) And EsConsonante(a1) And EsVocal(d1) Then
'        If EsAtaqueComplejo(a1, d1) Then
'            DebeCortar = True
'            Exit Function
'        End If
'    End If
'
'    ' ----------------------------------------------------
'    ' 3) C s V ? C - sV
'    ' frontera entre a1 y d1
'    ' ----------------------------------------------------
'    If a1 = "s" And EsConsonante(a2) And EsVocal(d1) Then
'        DebeCortar = True
'        Exit Function
'    End If
'
'    ' ----------------------------------------------------
'    ' 4) C C ? C - C (salvo ataque complejo o coda compleja)
'    ' frontera entre a1 y d1
'    ' ----------------------------------------------------
'    If EsConsonante(a1) And EsConsonante(d1) Then
'
'        ' No cortar si ataque complejo
'        If EsAtaqueComplejo(a1, d1) Then Exit Function
'
'        ' No cortar si coda compleja (C+s) y después consonante
'        If EsCodaCompleja(a1, d1) And EsConsonante(d2) Then Exit Function
'
'        DebeCortar = True
'        Exit Function
'    End If
'
'    ' ----------------------------------------------------
'    ' 5) Hiato obligatorio V[-alta] V[-alta]
'    ' ----------------------------------------------------
'    If EsVocal(a1) And EsVocal(d1) Then
'        If EsNoAlta(a1) And EsNoAlta(d1) Then
'            DebeCortar = True
'            Exit Function
'        End If
'    End If
'
'    DebeCortar = False
'End Function

'Private Function DebeCortar(ByVal Texto As String, ByVal i As Long) As Boolean
'    Dim L As Long
'    Dim a1 As String, a2 As String, a3 As String, a4 As String  ' izquierda
'    Dim d1 As String, d2 As String, d3 As String                ' derecha
'
'    L = Len(Texto)
'
'    ' Si estamos al final de la palabra, siempre cortamos
'    If i = L Then
'        DebeCortar = True
'        Exit Function
'    End If
'
'    ' Caracteres a la izquierda (a1 = justo antes de la frontera)
'    If i >= 1 Then a1 = Mid$(Texto, i, 1)
'    If i >= 2 Then a2 = Mid$(Texto, i - 1, 1)
'    If i >= 3 Then a3 = Mid$(Texto, i - 2, 1)
'    If i >= 4 Then a4 = Mid$(Texto, i - 3, 1)
'
'    ' Caracteres a la derecha (d1 = justo después de la frontera)
'    If i + 1 <= L Then d1 = Mid$(Texto, i + 1, 1)
'    If i + 2 <= L Then d2 = Mid$(Texto, i + 2, 1)
'    If i + 3 <= L Then d3 = Mid$(Texto, i + 3, 1)
'
'    ' ----------------------------------------------------
'    ' 1) V C V  ? V - C V  (a1 = C, a2 = V, d1 = V)
'    ' frontera entre a1 y d1
'    ' ... V C | V ...
'    ' ----------------------------------------------------
'    If EsVocal(a2) And EsConsonante(a1) And EsVocal(d1) Then
'        ' salvo que C+V formen ataque complejo (obstruyente+líquida)
'        If Not EsAtaqueComplejo(a1, d1) Then
'            DebeCortar = True
'            Exit Function
'        End If
'    End If
'
'    ' ----------------------------------------------------
'    ' 2) V C C V ? V - CCV  (ataque complejo)
'    ' patrón: a3 = V, a2 = C, a1 = C, d1 = V
'    ' frontera entre a1 y d1
'    ' ... V C C | V ...
'    ' ----------------------------------------------------
'    If EsVocal(a3) And EsConsonante(a2) And EsConsonante(a1) And EsVocal(d1) Then
'        If EsAtaqueComplejo(a1, d1) Then
'            DebeCortar = True
'            Exit Function
'        End If
'    End If
'
'    ' ----------------------------------------------------
'    ' 3) C s V ? C - sV
'    ' patrón: a1 = s, a2 = C, d1 = V
'    ' frontera entre a1 y d1
'    ' ... C s | V ...
'    ' ----------------------------------------------------
'    If EsConsonante(a2) And a1 = "s" And EsVocal(d1) Then
'        DebeCortar = True
'        Exit Function
'    End If
'
'    ' ----------------------------------------------------
'    ' 4) C C ? C - C (salvo ataque complejo o coda compleja)
'    ' frontera entre a1 y d1
'    ' ... C | C ...
'    ' ----------------------------------------------------
'    If EsConsonante(a1) And EsConsonante(d1) Then
'        ' No cortar si a1+d1 es ataque complejo
'        If EsAtaqueComplejo(a1, d1) Then Exit Function
'
'        ' No cortar si a1+d1 es coda compleja (C+s) y después viene consonante
'        If EsCodaCompleja(a1, d1) And EsConsonante(d2) Then Exit Function
'
'        DebeCortar = True
'        Exit Function
'    End If
'
'    ' ----------------------------------------------------
'    ' 5) Hiato obligatorio V[-alta] V[-alta]
'    ' frontera entre a1 y d1
'    ' ... V | V ...
'    ' ----------------------------------------------------
'    If EsVocal(a1) And EsVocal(d1) Then
'        If EsNoAlta(a1) And EsNoAlta(d1) Then
'            DebeCortar = True
'            Exit Function
'        End If
'    End If
'
'    DebeCortar = False
'End Function

'Private Function DebeCortar( _
'    ByVal ult As String, _
'    ByVal c As String, _
'    ByVal nextC As String, _
'    ByVal next2 As String, _
'    ByVal next3 As String, _
'    ByVal silaba As String) As Boolean
'
'    DebeCortar = False
'
'    ' 1) Final de palabra ? cortar
'    If nextC = "" Then
'        DebeCortar = True
'        Exit Function
'    End If
'
'    ' ============================================================
'    ' 1) V C V ? V - C V
'    ' ============================================================
'    If EsVocal(ult) And EsConsonante(c) And EsVocal(nextC) Then
'        If Not EsAtaqueComplejo(c, nextC) Then
'            DebeCortar = True
'            Exit Function
'        End If
'    End If
'
'    ' ============================================================
'    ' 2) V C C V ? V - CCV (ataque complejo)
'    ' ============================================================
'    If EsVocal(ult) And EsConsonante(c) And EsConsonante(nextC) And EsVocal(next2) Then
'        If EsAtaqueComplejo(c, nextC) Then
'            DebeCortar = True
'            Exit Function
'        End If
'    End If
'
'    ' ============================================================
'    ' 3) C s V ? C - sV
'    ' ============================================================
'    If EsConsonante(ult) And c = "s" And EsVocal(nextC) Then
'        DebeCortar = True
'        Exit Function
'    End If
'
'    ' ============================================================
'    ' *** NUEVO: NO aplicar C-C si estamos en VCCV ***
'    ' ============================================================
'    ' Detectar V C C V correctamente
'If Len(silaba) >= 3 Then
'    Dim v As String, c1 As String, c2 As String
'
'    v = Mid$(silaba, Len(silaba) - 2, 1)   ' vocal previa
'    c1 = Mid$(silaba, Len(silaba) - 1, 1)  ' consonante 1
'    c2 = c                                 ' consonante 2
'
'    If EsVocal(v) _
'       And EsConsonante(c1) _
'       And EsConsonante(c2) _
'       And EsVocal(nextC) Then
'
'        Exit Function
'    End If
'End If

'    ' Detectar V C C V correctamente
'If Len(silaba) >= 2 Then
'    Dim c1 As String, c2 As String
'    c1 = Mid$(silaba, Len(silaba) - 1, 1)   ' consonante 1
'    c2 = c                                  ' consonante 2
'
'    If EsVocal(Mid$(silaba, Len(silaba) - 2, 1)) _
'       And EsConsonante(c1) _
'       And EsConsonante(c2) _
'       And EsVocal(nextC) Then
'
'        Exit Function
'    End If
'End If

'    If EsVocal(ult) And EsConsonante(c) And EsConsonante(nextC) And EsVocal(next2) Then
'        Exit Function
'    End If

'    ' ============================================================
'    ' *** NUEVO: NO aplicar C-C si estamos en CCC ***
'    ' ============================================================
'    If EsConsonante(ult) And EsConsonante(c) And EsConsonante(nextC) Then
'        Exit Function
'    End If
'
'    ' ============================================================
'    ' 4) C C ? C - C (salvo ataque complejo o coda compleja)
'    ' ============================================================
'    If EsConsonante(ult) And EsConsonante(c) Then
'
'        ' No cortar si ult+c es ataque complejo
'        If EsAtaqueComplejo(ult, c) Then Exit Function
'
'        ' No cortar si ult+c es coda compleja (C+s) y nextC es consonante
'        If EsCodaCompleja(ult, c) And EsConsonante(nextC) Then Exit Function
'
'        DebeCortar = True
'        Exit Function
'    End If
'
'    ' ============================================================
'    ' 5) Hiato obligatorio V[-alta] V[-alta]
'    ' ============================================================
'    If EsVocal(ult) And EsVocal(c) Then
'        If EsNoAlta(ult) And EsNoAlta(c) Then
'            DebeCortar = True
'            Exit Function
'        End If
'    End If
'
'End Function

'Private Function DebeCortar( _
'    ByVal ult As String, _
'    ByVal c As String, _
'    ByVal nextC As String, _
'    ByVal next2 As String, _
'    ByVal next3 As String, _
'    ByVal silaba As String) As Boolean
'
'    DebeCortar = False
'
'    ' 1) Final de palabra ? cortar
'    If nextC = "" Then
'        DebeCortar = True
'        Exit Function
'    End If
'
'    ' ============================================================
'    ' 1) V C V ? V - C V
'    ' ============================================================
'    If EsVocal(ult) And EsConsonante(c) And EsVocal(nextC) Then
'        If Not EsAtaqueComplejo(c, nextC) Then
'            DebeCortar = True
'            Exit Function
'        End If
'    End If
'
'    ' ============================================================
'    ' 2) V C C V ? V - CCV (ataque complejo)
'    ' ============================================================
'    If EsVocal(ult) And EsConsonante(c) And EsConsonante(nextC) And EsVocal(next2) Then
'        If EsAtaqueComplejo(c, nextC) Then
'            DebeCortar = True
'            Exit Function
'        End If
'    End If
'
'    ' ============================================================
'    ' 3) C s V ? C - sV
'    ' ============================================================
'    If EsConsonante(ult) And c = "s" And EsVocal(nextC) Then
'        DebeCortar = True
'        Exit Function
'    End If
'
'    ' ============================================================
'    ' 4) C C ? C - C (salvo ataque complejo o coda compleja)
'    ' ============================================================
'    If EsConsonante(ult) And EsConsonante(c) Then
'
'        ' No cortar si ult+c es ataque complejo
'        If EsAtaqueComplejo(ult, c) Then
'            Exit Function
'        End If
'
'        ' No cortar si ult+c es coda compleja (C+s) y next2 es consonante
'        If EsCodaCompleja(ult, c) And EsConsonante(nextC) Then
'            Exit Function
'        End If
'
'        DebeCortar = True
'        Exit Function
'    End If
'
'    ' ============================================================
'    ' 5) Hiato obligatorio V[-alta] V[-alta]
'    ' ============================================================
'    If EsVocal(ult) And EsVocal(c) Then
'        If EsNoAlta(ult) And EsNoAlta(c) Then
'            DebeCortar = True
'            Exit Function
'        End If
'    End If
'
'End Function


' ================================================================
'   FUNCIONES AUXILIARES (comparación binaria)
' ================================================================

Private Function EsVocal(ByVal ch As String) As Boolean
    EsVocal = (InStr(1, "aeiouáéíóúü", ch, vbBinaryCompare) > 0)
End Function

Private Function EsConsonante(ByVal ch As String) As Boolean
    If ch = "" Then
        EsConsonante = False
    Else
        EsConsonante = Not EsVocal(ch)
    End If
End Function

Private Function EsAlta(ByVal ch As String) As Boolean
    EsAlta = (InStr(1, "iuíúü", ch, vbBinaryCompare) > 0)
End Function

Private Function EsNoAlta(ByVal ch As String) As Boolean
    EsNoAlta = (EsVocal(ch) And Not EsAlta(ch))
End Function

Private Function EsObstruyente(ByVal ch As String) As Boolean
    EsObstruyente = (InStr(1, "ptkbdgfzsxc", ch, vbBinaryCompare) > 0)
End Function

Private Function EsLiquida(ByVal ch As String) As Boolean
    EsLiquida = (ch = "l" Or ch = "r")
End Function

Private Function EsAtaqueComplejo(ByVal C1 As String, ByVal C2 As String) As Boolean
    If Not EsObstruyente(C1) Then Exit Function
    If Not EsLiquida(C2) Then Exit Function

    ' Excepciones: tl, dl no son ataque complejo
    If C1 = "t" And C2 = "l" Then Exit Function
    If C1 = "d" And C2 = "l" Then Exit Function

    EsAtaqueComplejo = True
End Function

Private Function EsCodaCompleja(ByVal C1 As String, ByVal C2 As String) As Boolean
    If C2 <> "s" Then Exit Function
    If InStr(1, "bdknlr", C1, vbBinaryCompare) > 0 Then
        EsCodaCompleja = True
    End If
End Function


