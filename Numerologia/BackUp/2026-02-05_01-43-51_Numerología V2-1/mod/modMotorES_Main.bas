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

    Call NormalizarConReglas
    Call SilabearAuto
    Call DetectarTonicaGeneral
    Call MarcarSilabasTonicas
    Call RevisionSilabeo
    Call ReconstruirSilabasFinales
    Call ConvertirSilabasAFonemas

    EntradaMotor_ES = objDTO.TextoFinal

End Function

' ============================================================
'   NORMALIZACIÓN
' ============================================================
Private Sub NormalizarConReglas()

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

    ' Normalizar vocales (sin tocar ü)
    's = MF_NormalizarVocales_ES(s)

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

''        If c1 = " " Then
''            If i - 2 >= ini Then col.Add Array(ini, i - 2)
''            ini = i
''            GoTo siguiente
''        End If
'
'        If c1 = " " Then
'            If i - 2 >= ini Then col.Add Array(ini, i - 2)
'            col.Add Array(i - 1, i - 1)   ' --> añade la sílaba vacía
'            ini = i
'            GoTo siguiente
'        End If
'
''        If c2 = " " Then
''            col.Add Array(ini, i - 1)
''            ini = i + 1
''            GoTo siguiente
''        End If
'        If c2 = " " Then
'            col.Add Array(ini, i - 1)
'            col.Add Array(i, i)   ' --> añade la sílaba vacía
'            ini = i + 1
'            GoTo siguiente
'        End If

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

'        If (c1 = "í" Or c1 = "ú") Or (c2 = "í" Or c2 = "ú") Then
'            col.Add Array(ini, i - 1)
'            ini = i
'            GoTo siguiente
'        End If

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
Private Sub DetectarTonicaGeneral()

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

    Dim i As Long, j As Long
    Dim strFinal As String
    Dim esTonica As Boolean
    Dim sil As String
    Dim arrFon() As Long
    Dim f As Variant 'Long

    Dim arrSilabas() As String
    Dim arrTonica() As String
    
    
    ' Convertir strings del DTO a arrays
    arrSilabas = Split(objDTO.SilabasFinal, "-")
    arrTonica = Split(objDTO.SilabaTonica, ",")

    If UBound(arrSilabas) < 0 Then Exit Sub

    strFinal = ""

    For i = 0 To UBound(arrSilabas)

        sil = arrSilabas(i)

        'sil = objDTO.SilabasFinal(i)
        sil = arrSilabas(i)


'        esTonica = False
'        If UBound(objDTO.SilabaTonica) >= 1 Then
'            For j = 1 To UBound(objDTO.SilabaTonica)
'                'If objDTO.SilabaTonica(j) = i Then
'                If CByte(arrTonica(j)) = i + 1 Then
'
'                    esTonica = True
'                    Exit For
'                End If
'            Next j
'        End If

        esTonica = False
        arrTonica = Split(objDTO.SilabaTonica, ",")
        
        If UBound(arrTonica) >= 0 Then
            For j = 0 To UBound(arrTonica)
                If CByte(arrTonica(j)) = i + 1 Then
                    esTonica = True
                    Exit For
                End If
            Next j
        End If

        IndiceSilabaActual = i
        EsTonicaActual = esTonica

        If esTonica Then
            strFinal = strFinal & "61, "
        End If

        If sil = " " Then
            strFinal = strFinal & "0 - "
            GoTo siguiente
        End If

        arrFon = ConvertirGrafemasDeSilabaAIdFonemas()
        
        For Each f In arrFon
            strFinal = strFinal & CStr(f) & ", "
        Next f

        strFinal = strFinal & "- "

siguiente:
    Next i

    objDTO.TextoFinal = Trim$(strFinal)

End Sub


' ============================================================
'   CONVERSIÓN GRAFEMAS ? IDFONEMAS
' ============================================================
Private Function ConvertirGrafemasDeSilabaAIdFonemas() As Long()

    Dim sil As String
    Dim s As String
    Dim i As Long
    Dim graf As String
    Dim fon As Byte
    Dim arr() As Long
    Dim idx As Long
    
    Dim arrSilabas() As String

'    sil = objDTO.SilabasFinal(IndiceSilabaActual)
    arrSilabas = Split(objDTO.SilabasFinal, "-")
    sil = arrSilabas(IndiceSilabaActual - 1)

    s = LCase$(sil)
    ' Normalizar vocales (sin tocar ü)
    s = MF_NormalizarVocales_ES(s)

    ReDim arr(1 To 1)
    idx = 1
    i = 1

    Do While i <= Len(s)

        GrafAnterior = ""
        GrafActual = ""
        GrafSiguiente = ""

        ' TRIGRAFEMAS
        If i <= Len(s) - 2 Then
            graf = Mid$(s, i, 3)
            If graf = "güe" Or graf = "güi" Or _
               graf = "gue" Or graf = "gui" Or _
               graf = "que" Or graf = "qui" Then

                If i > 1 Then GrafAnterior = Mid$(s, i - 1, 1)
                If i < Len(s) - 2 Then GrafSiguiente = Mid$(s, i + 3, 1)
                GrafActual = graf

                fon = ReglasCastellano(graf, GrafAnterior, GrafSiguiente, EsTonicaActual)

                If fon > 0 Then
                    arr(idx) = fon
                    idx = idx + 1
                    ReDim Preserve arr(1 To idx)
                    i = i + 3
                    GoTo siguiente
                End If
            End If
        End If

        ' DÍGRAFOS
        If i <= Len(s) - 1 Then
            graf = Mid$(s, i, 2)
            If graf = "ch" Or graf = "ll" Or graf = "rr" Or _
               graf = "gu" Or graf = "qu" Or _
               graf = "ai" Or graf = "ei" Or graf = "oi" Or graf = "ou" Or graf = "au" Then

                If i > 1 Then GrafAnterior = Mid$(s, i - 1, 1)
                If i < Len(s) - 1 Then GrafSiguiente = Mid$(s, i + 2, 1)
                GrafActual = graf

                fon = ReglasCastellano(graf, GrafAnterior, GrafSiguiente, EsTonicaActual)

                If fon > 0 Then
                    arr(idx) = fon
                    idx = idx + 1
                    ReDim Preserve arr(1 To idx)
                    i = i + 2
                    GoTo siguiente
                End If
            End If
        End If

        ' MONÓGRAFOS
        graf = Mid$(s, i, 1)
        GrafActual = graf
        If i > 1 Then GrafAnterior = Mid$(s, i - 1, 1)
        If i < Len(s) Then GrafSiguiente = Mid$(s, i + 1, 1)

        fon = ReglasCastellano(graf, GrafAnterior, GrafSiguiente, EsTonicaActual)

        If fon > 0 Then
            arr(idx) = fon
            idx = idx + 1
            ReDim Preserve arr(1 To idx)
        End If

        i = i + 1

siguiente:
    Loop

    If idx > 1 Then
        ReDim Preserve arr(1 To idx - 1)
    Else
        ReDim arr(1 To 0)
    End If

    ConvertirGrafemasDeSilabaAIdFonemas = arr

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




'Private Function ContarSilabasDePalabra( _
'                        ByVal palabra As String, _
'                        ByRef sils() As String, _
'                        ByVal offset As Long) As Long
'
'    Dim total As Long
'    Dim i As Long
'
'    total = 0
'
'    For i = offset To UBound(sils)
'        If InStr(1, palabra, sils(i), vbTextCompare) > 0 Then
'            total = total + 1
'        Else
'            Exit For
'        End If
'    Next i
'
'    ContarSilabasDePalabra = total
'End Function


'Private Sub RevisionSilabeo()
'
'    Dim s As String
'    Dim partes() As String
'    Dim resultado() As String
'    Dim tFinal() As Byte
'    Dim texto As String
'    Dim i As Long
'    Dim idxTonica As Collection
'    Dim raw As String
'    Dim limpio As String
'    Dim p1 As Long, p2 As Long
'
'    texto = objDTO.TextoNormalizado
'
'    s = ""
'
'    For i = 1 To UBound(objDTO.SilabasAuto)
'        s = s & objDTO.SilabasAuto(i)
'        If i < UBound(objDTO.SilabasAuto) Then s = s & "-"
'    Next i
'
'    s = RevisarSilabas_EnFormulario(texto, s)
'
'    If s = "" Then
'        objDTO.SilabasFinal = objDTO.SilabasAuto
'        objDTO.SilabaTonica = objDTO.SilabaTonica
'        Exit Sub
'    End If
'
'    partes = Split(s, "-")
'
'    Set idxTonica = New Collection
'
'    For i = LBound(partes) To UBound(partes)
'
'        raw = partes(i)
'
'        If raw = " " Then
'            partes(i) = " "
'            GoTo siguiente
'        End If
'
'        If InStr(raw, "*") > 0 Then
'
'            p1 = InStr(1, raw, "*")
'            p2 = InStrRev(raw, "*")
'
'            idxTonica.Add i + 1
'
'            limpio = Mid$(raw, p1 + 1, p2 - p1 - 1)
'
'            partes(i) = limpio
'
'        Else
'            partes(i) = raw
'        End If
'
'siguiente:
'    Next i
'
'    ReDim resultado(1 To UBound(partes) + 1)
'
'    For i = 1 To UBound(resultado)
'        resultado(i) = partes(i - 1)
'    Next i
'
'   objDTO.SilabasFinal = resultado
'
'    If idxTonica.Count > 0 Then
'        ReDim tFinal(1 To idxTonica.Count)
'        For i = 1 To idxTonica.Count
'            tFinal(i) = idxTonica(i)
'        Next i
'        objDTO.SilabaTonica = tFinal
'    End If
'
'End Sub

'Private Sub RevisionSilabeo()
'
'    Dim texto As String
'    Dim s As String
'    Dim partes() As String
'    Dim resultado As String
'    Dim idxTonica As New Collection
'    Dim raw As String
'    Dim limpio As String
'    Dim p1 As Long, p2 As Long
'    Dim i As Long
'
'    'texto = objDTO.TextoNormalizado
'    texto = objDTO.TextoOriginal
'
'    ' SilabasAuto ya es un string con guiones
'    s = objDTO.SilabasAuto
'
'    ' Abrir formulario de revisión
'    s = RevisarSilabas_EnFormulario(texto, s)
'
'    ' Si el usuario cancela, conservar lo anterior
'    If s = "" Then
'        objDTO.SilabasFinal = objDTO.SilabasAuto
'        objDTO.SilabaTonica = objDTO.SilabaTonica
'        Exit Sub
'    End If
'
'    ' Dividir sílabas revisadas
'    partes = Split(s, "-")
'
'    ' Procesar cada sílaba
'    For i = LBound(partes) To UBound(partes)
'
'        raw = partes(i)
'
'        ' Si es espacio literal
'        If raw = " " Then
'            partes(i) = " "
'            GoTo siguiente
'        End If
'
'        ' Si contiene marca de tónica
'        If EsSilabaMarcada(raw) Then
'
'            idxTonica.Add i + 1
'            partes(i) = LimpiarSilabaMarcada(raw)
'
'        Else
'            partes(i) = raw
'        End If
'
''        If InStr(raw, "*") > 0 Then
''
''            p1 = InStr(1, raw, "*")
''            p2 = InStrRev(raw, "*")
''
''            ' Guardar índice (1-based)
''            idxTonica.Add i + 1
''
''            ' Limpiar la sílaba
''            limpio = Mid$(raw, p1 + 1, p2 - p1 - 1)
''            partes(i) = limpio
''
''        Else
''            partes(i) = raw
''        End If
'
'siguiente:
'    Next i
'
'    ' Reconstruir SilabasFinal como string
'    resultado = Join(partes, "-")
'    objDTO.SilabasFinal = resultado
'
'    ' Reconstruir SilabaTonica como string
'    If idxTonica.Count > 0 Then
'        Dim arr() As String
'        ReDim arr(0 To idxTonica.Count - 1)
'
'        For i = 1 To idxTonica.Count
'            arr(i - 1) = CStr(idxTonica(i))
'        Next i
'
'        objDTO.SilabaTonica = Join(arr, ",")
'    End If
'
'End Sub

'Private Function EsSilabaMarcada(ByVal s As String) As Boolean
'    Dim t As String
'    t = Trim$(s)
'    EsSilabaMarcada = (Left$(t, 1) = "*" And Right$(t, 1) = "*")
'End Function

'Private Function LimpiarSilabaMarcada(ByVal s As String) As String
'    Dim t As String
'    t = Trim$(s)
'    t = Mid$(t, 2, Len(t) - 2) ' quitar los dos *
'    LimpiarSilabaMarcada = Trim$(t)
'End Function

'Private Sub DetectarTonicaGeneral()
'
'    Dim sils() As String
'    Dim i As Long
'    Dim idx As Long
'    Dim arrTonica() As String
'    Dim countTonica As Long
'
'    Dim palabra As String
'    Dim palabras() As String
'    Dim offset As Long
'    Dim silsPalabra() As String
'    Dim j As Long
'
'    ' Si no hay sílabas, no hay nada que hacer
'    If Len(objDTO.SilabasAuto) = 0 Then
'        objDTO.SilabaTonica = ""
'        Exit Sub
'    End If
'
'    ' Dividir en palabras (por espacios)
'    palabras = Split(objDTO.TextoNormalizado, " ")
'
'    ' Dividir sílabas globales
'    sils = Split(objDTO.SilabasAuto, "-")
'
'    ' Preparamos array dinámico para tónicas
'    ReDim arrTonica(0 To 0)
'    countTonica = 0
'
'    offset = 0   ' desplazamiento de sílabas acumuladas
'
'    ' Procesar palabra por palabra
'    For i = LBound(palabras) To UBound(palabras)
'
'        palabra = palabras(i)
'
'        ' Extraer sílabas de esta palabra
'        ' (todas las sílabas cuyo texto pertenece a esta palabra)
'        ' Para simplificar: contamos sílabas por longitud acumulada
'
'        Dim numSilabasPalabra As Long
'        numSilabasPalabra = ContarSilabasDePalabra(palabra, sils, offset)
'
'        If numSilabasPalabra > 0 Then
'
'            ' Detectar tónica de esta palabra
'            idx = DetectarTonicaDeUnaPalabra(palabra, numSilabasPalabra)
'
'            If idx > 0 Then
'                ' Guardar índice global
'                countTonica = countTonica + 1
'                ReDim Preserve arrTonica(0 To countTonica - 1)
'                arrTonica(countTonica - 1) = CStr(offset + idx)
'            End If
'
'            ' Avanzar offset global
'            offset = offset + numSilabasPalabra
'        End If
'
'    Next i
'
'    ' Guardar resultado final
'    If countTonica = 0 Then
'        objDTO.SilabaTonica = ""
'    Else
'        objDTO.SilabaTonica = Join(arrTonica, ",")
'    End If
'
'End Sub

