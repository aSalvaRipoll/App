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

    texto = objDTO.TextoNormalizado
    texto = Trim$(texto)

    If Len(texto) = 0 Then
        ReDim objDTO.SilabasAuto(0)
        Exit Sub
    End If

    ini = 1

    For i = 2 To Len(texto)

        c1 = Mid$(texto, i - 1, 1)
        c2 = Mid$(texto, i, 1)
        par = c1 & c2

        ' 0. ESPACIOS
        If c1 = " " Then
            If i - 2 >= ini Then col.Add Array(ini, i - 2)
            ini = i
            GoTo siguiente
        End If

        If c2 = " " Then
            col.Add Array(ini, i - 1)
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
        If (c1 = "í" Or c1 = "ú") Or (c2 = "í" Or c2 = "ú") Then
            col.Add Array(ini, i - 1)
            ini = i
            GoTo siguiente
        End If

        ' HIATO FUERTE + FUERTE
        If EsHiatoFuerteFuerte(c1, c2) Then
            col.Add Array(ini, i - 1)
            ini = i
            GoTo siguiente
        End If

        ' DIPTONGO
        If EsDiptongo(c1, c2) Then GoTo siguiente

        ' VV ? separar
        If EsVocal_ES(c1) And EsVocal_ES(c2) Then
            col.Add Array(ini, i - 1)
            ini = i
            GoTo siguiente
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

    ReDim objDTO.SilabasAuto(1 To col.Count)

    For i = 1 To col.Count
        objDTO.SilabasAuto(i) = Mid$(texto, col(i)(0), col(i)(1) - col(i)(0) + 1)
    Next i

End Sub



' ============================================================
'   DETECCIÓN DE TÓNICA
' ============================================================
Private Sub DetectarTonicaGeneral()

    Dim i As Long, j As Long
    Dim texto As String
    Dim sils() As String
    Dim vocalesTilde As String
    Dim idx As Long

    texto = objDTO.TextoNormalizado
    sils = objDTO.SilabasAuto
    vocalesTilde = "áéíóú"

    For i = 1 To UBound(sils)
        For j = 1 To Len(sils(i))
            If InStr(vocalesTilde, Mid$(sils(i), j, 1)) > 0 Then
                idx = i
                Exit For
            End If
        Next j
        If idx > 0 Then Exit For
    Next i

    If idx = 0 Then
        Dim ultima As String
        ultima = Right$(Trim$(texto), 1)

        If ultima = "n" Or ultima = "s" Or InStr("aeiou", ultima) > 0 Then
            If UBound(sils) >= 2 Then
                idx = UBound(sils) - 1
            Else
                idx = UBound(sils)
            End If
        Else
            idx = UBound(sils)
        End If
    End If

    ReDim objDTO.SilabaTonica(1 To 1)
    objDTO.SilabaTonica(1) = idx

End Sub


' ============================================================
'   REVISIÓN MANUAL
' ============================================================
Private Sub RevisionSilabeo()

    Dim s As String
    Dim partes() As String
    Dim resultado() As String
    Dim tFinal() As Byte
    Dim texto As String
    Dim i As Long
    Dim idxTonica As Collection
    Dim raw As String
    Dim limpio As String
    Dim p1 As Long, p2 As Long

    texto = objDTO.TextoNormalizado

    s = ""
    For i = 1 To UBound(objDTO.SilabasAuto)
        s = s & objDTO.SilabasAuto(i)
        If i < UBound(objDTO.SilabasAuto) Then s = s & "-"
    Next i

    s = RevisarSilabas_EnFormulario(texto, s)

    If s = "" Then
        objDTO.SilabasFinal = objDTO.SilabasAuto
        objDTO.SilabaTonica = objDTO.SilabaTonica
        Exit Sub
    End If

    partes = Split(s, "-")

    Set idxTonica = New Collection

    For i = LBound(partes) To UBound(partes)

        raw = partes(i)

        If raw = " " Then
            partes(i) = " "
            GoTo siguiente
        End If

        If InStr(raw, "*") > 0 Then

            p1 = InStr(1, raw, "*")
            p2 = InStrRev(raw, "*")

            idxTonica.Add i + 1

            limpio = Mid$(raw, p1 + 1, p2 - p1 - 1)

            partes(i) = limpio

        Else
            partes(i) = raw
        End If

siguiente:
    Next i

    ReDim resultado(1 To UBound(partes) + 1)

    For i = 1 To UBound(resultado)
        resultado(i) = partes(i - 1)
    Next i

    objDTO.SilabasFinal = resultado

    If idxTonica.Count > 0 Then
        ReDim tFinal(1 To idxTonica.Count)
        For i = 1 To idxTonica.Count
            tFinal(i) = idxTonica(i)
        Next i
        objDTO.SilabaTonica = tFinal
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
    Dim f As Long

    If UBound(objDTO.SilabasFinal) < 1 Then Exit Sub

    strFinal = ""

    For i = 1 To UBound(objDTO.SilabasFinal)

        sil = objDTO.SilabasFinal(i)

        esTonica = False
        If UBound(objDTO.SilabaTonica) >= 1 Then
            For j = 1 To UBound(objDTO.SilabaTonica)
                If objDTO.SilabaTonica(j) = i Then
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

    sil = objDTO.SilabasFinal(IndiceSilabaActual)
    s = LCase$(sil)

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




