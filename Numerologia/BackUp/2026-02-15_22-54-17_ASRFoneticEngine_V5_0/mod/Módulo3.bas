Attribute VB_Name = "Módulo3"
'' ============================================================
'' Nombre:    bas-Motor_ES_Main
'' Tipo:      Módulo
'' Propósito: Módulo del motor fonético aplicado al Español (Castellano)
'' Autor:     Alba Salvá
'' Fecha:     11/02/2026
'' ============================================================
'
'' ================================
''   Módulo: MotorSilabico20
''   Versión: 2.0 (módulo independiente)
'' ================================
'' Expone:
''   - SilabearPalabra(ByVal Texto As String) As String
''   - SilabearTexto(ByVal Texto As String) As String
'' ================================
'
'Option Compare Database
'Option Explicit
'
'Public ObjDTO As clsDTO_Motor
'
'' ============================================================
''   ENTRADA PRINCIPAL DEL MOTOR (ESPAÑOL)
'' ============================================================
'' ----------------------------------------------------------------
'' Procedimiento: Entrada_Motor_ES
'' Propósito:     Punto de entrada al motor fonético del español general
'' Tipo proc.:    Function
'' Acceso proc.:  Public
'
'' Parameter Texto
''    (String): Texto que se recibe (nombre o apellido español general)
'
'' Tipo retorno: String -> Texto que contiene la lista de fonemas
''   resultado de la conversión
'
'' Autor:        Alba Salvá
'' Fecha:        11/02/2026
'' ----------------------------------------------------------------
'Public Function Entrada_Motor_ES(Texto As String) As String
'
'    Set ObjDTO = New clsDTO_Motor
'
'
'    strDebug = ""
'
'    ' 1) Asignamos el texto recibido y
'    '    Normalización (dentro del DTO)
'    ObjDTO.TextoOriginal = Texto
'    ObjDTO.NormalizaEntrada
'
'    ' 2) Silabeo automático
'    Call Silabear
'
'    Call MF_DebugDTO("Silabear")
'
'    ' 3) Devolver resultado
'    Entrada_Motor_ES = ObjDTO.SilabasAuto
'
'End Function
'
'
'' ============================================================
''  1.- SILABEO AUTOMÁTICO
'' ============================================================
'' ----------------------------------------------------------------
'' Procedimiento: Silabear (MÓDULO ÚNICO)
'' Propósito:     Convierte las palabras ya normalizadas en sílabas.
'' Tipo proc.:    Sub
'' Acceso proc.:  Private
'
'' Tipo retorno: Ninguno
'
'' Autor:        Alba Salvá
'' Fecha:        11/02/2026
'' ----------------------------------------------------------------
'Public Function Silabear(ByVal Texto As String) As String
'    Dim partes() As String
'    Dim i As Long
'    Dim res As String
'
'    partes = Split(Trim$(Texto), " ")
'    For i = LBound(partes) To UBound(partes)
'        If partes(i) <> "" Then
'            If res <> "" Then res = res & " "
'            res = res & SilabearPalabra(partes(i))
'        End If
'    Next i
'
'    SilabearTexto = res
'End Function
'
'Public Function SilabearPalabra(ByVal Texto As String) As String
'    Dim i As Long
'    Dim silabaActual As String
'    Dim resultado As String
'    Dim c As String
'    Dim prev As String
'    Dim nextC As String
'    Dim next2 As String
'    Dim next3 As String
'    Dim ult As String
'
'    Texto = LCase$(Texto)
'    Texto = Trim$(Texto)
'    If Texto = "" Then
'        SilabearPalabra = ""
'        Exit Function
'    End If
'
'    silabaActual = ""
'    resultado = ""
'
'    For i = 1 To Len(Texto)
'        c = Mid$(Texto, i, 1)
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
'        silabaActual = silabaActual & c
'
'        ult = ""
'        If Len(silabaActual) > 0 Then
'            ult = Right$(silabaActual, 1)
'        End If
'
'        If DebeCortar(ult, c, nextC, next2, next3, silabaActual) Then
'            If resultado <> "" Then
'                resultado = resultado & " | "
'            End If
'            resultado = resultado & silabaActual
'            silabaActual = ""
'        End If
'    Next i
'
'    If silabaActual <> "" Then
'        If resultado <> "" Then
'            resultado = resultado & " | "
'        End If
'        resultado = resultado & silabaActual
'    End If
'
'    SilabearPalabra = resultado
'End Function
'
'Private Function DebeCortar( _
'    ByVal ult As String, _
'    ByVal c As String, _
'    ByVal nextC As String, _
'    ByVal next2 As String, _
'    ByVal next3 As String, _
'    ByVal silaba As String) As Boolean
'
'    ' No cortamos si estamos al final de palabra
'    If nextC = "" Then
'        DebeCortar = True
'        Exit Function
'    End If
'
'    ' =========================
'    ' 1) V C V  -> V - C V
'    ' =========================
'    If EsVocal(ult) And EsConsonante(nextC) And EsVocal(next2) Then
'        ' Pero si nextC + next2 forman ataque complejo (obstruyente + líquida),
'        ' no cortamos aquí, lo hará la regla de VCCV.
'        If Not EsAtaqueComplejo(nextC, next2) Then
'            DebeCortar = True
'            Exit Function
'        End If
'    End If
'
'    ' =========================
'    ' 2) V C C V  (ataque complejo)
'    '    V C C V -> V - C C V
'    ' =========================
'    If EsVocal(ult) And EsConsonante(nextC) And EsConsonante(next2) And EsVocal(next3) Then
'        If EsAtaqueComplejo(nextC, next2) Then
'            DebeCortar = True
'            Exit Function
'        End If
'    End If
'
'    ' =========================
'    ' 3) C s V -> C - s V
'    ' =========================
'    If EsConsonante(ult) And c = "s" And EsVocal(nextC) Then
'        DebeCortar = True
'        Exit Function
'    End If
'
'    ' =========================
'    ' 4) C C (coda - ataque)
'    '    Entre dos consonantes siempre hay frontera
'    '    salvo que formen ataque complejo o coda compleja con s
'    ' =========================
'    If EsConsonante(ult) And EsConsonante(nextC) Then
'        ' No cortar si ult + nextC es ataque complejo (obstruyente + líquida)
'        If EsAtaqueComplejo(ult, nextC) Then
'            DebeCortar = False
'            Exit Function
'        End If
'
'        ' No cortar si ult + nextC es coda compleja con s (C + s) y después viene consonante
'        If EsCodaCompleja(ult, nextC) And EsConsonante(next2) Then
'            DebeCortar = False
'            Exit Function
'        End If
'
'        ' En el resto de casos, cortamos: C - C
'        DebeCortar = True
'        Exit Function
'    End If
'
'    ' =========================
'    ' 5) Hiatos obligatorios V[-alta] V[-alta]
'    ' =========================
'    If EsVocal(ult) And EsVocal(nextC) Then
'        If EsNoAlta(ult) And EsNoAlta(nextC) Then
'            DebeCortar = True
'            Exit Function
'        End If
'    End If
'
'    ' =========================
'    ' 6) Por defecto, no cortar
'    ' =========================
'    DebeCortar = False
'End Function
'
'' =========================
''   Funciones auxiliares
'' =========================
'
'Private Function EsVocal(ByVal ch As String) As Boolean
'    EsVocal = (InStr(1, "aeiouáéíóúü", ch, vbBinaryCompare) > 0)
'End Function
'
'Private Function EsConsonante(ByVal ch As String) As Boolean
'    If ch = "" Then
'        EsConsonante = False
'    Else
'        EsConsonante = Not EsVocal(ch)
'    End If
'End Function
'
'Private Function EsAlta(ByVal ch As String) As Boolean
'    EsAlta = (InStr(1, "iuíúü", ch, vbBinaryCompare) > 0)
'End Function
'
'Private Function EsNoAlta(ByVal ch As String) As Boolean
'    EsNoAlta = (EsVocal(ch) And Not EsAlta(ch))
'End Function
'
'Private Function EsObstruyente(ByVal ch As String) As Boolean
'    ' p, t, k, b, d, g, f, s, z, x, c (aprox. sistema)
'    EsObstruyente = (InStr(1, "ptkbdgfzsxc", ch, vbBinaryCompare) > 0)
'End Function
'
'Private Function EsLiquida(ByVal ch As String) As Boolean
'    EsLiquida = (ch = "l" Or ch = "r")
'End Function
'
'Private Function EsAtaqueComplejo(ByVal c1 As String, ByVal c2 As String) As Boolean
'    ' Grupos de obstruyente + líquida permitidos en ataque
'    If Not EsObstruyente(c1) Then Exit Function
'    If Not EsLiquida(c2) Then Exit Function
'
'    ' Excluimos combinaciones no aceptadas (tl, dl)
'    If c1 = "t" And c2 = "l" Then Exit Function
'    If c1 = "d" And c2 = "l" Then Exit Function
'
'    EsAtaqueComplejo = True
'End Function
'
'Private Function EsCodaCompleja(ByVal c1 As String, ByVal c2 As String) As Boolean
'    ' Coda compleja típica: C + s, con C en {b, d, k, n, l, r}
'    If c2 <> "s" Then Exit Function
'    If InStr(1, "bdknlr", c1, vbBinaryCompare) > 0 Then
'        EsCodaCompleja = True
'    End If
'End Function
'

