Attribute VB_Name = "modMotor_Idioma_EU"

Option Compare Database
Option Explicit

'=================
'==   Euskara   ==
'=================

Public Sub MF_MarcarTonica_EU( _
        ByVal texto As String, _
        ByRef esTonica() As Boolean _
    )

    Dim Silabas As Collection
    Dim idxTonica As Long
    Dim inicio As Long, fin As Long
    Dim i As Long

    ' 1. Silabear palabra (motor con revisión)
    Set Silabas = Silabear_EU_ConRevision(texto)

    If Silabas Is Nothing Then Exit Sub
    If Silabas.Count = 0 Then Exit Sub

    ' 2. Euskera --> acento fijo en la penúltima sílaba
    If Silabas.Count = 1 Then
        idxTonica = 1
    Else
        idxTonica = Silabas.Count - 1
    End If

    ' 3. Marcar índices tónicos
    inicio = Silabas(idxTonica)(1)
    fin = Silabas(idxTonica)(2)

    For i = inicio To fin
        esTonica(i) = True
    Next i

End Sub

Public Function Silabear_EU_ConRevision(ByVal texto As String) As Collection

    Dim col As Collection
    Dim resultado As New Collection
    Dim item As Variant
    Dim s As String
    Dim partes As Variant
    Dim p As Variant
    Dim inicio As Long, fin As Long
    Dim valido As Boolean
    Dim msg As String

    ' 1. Silabear automáticamente (motor puro euskera)
    Set col = Silabear_EU(texto)

    ' 2. Convertir a string con "-"
    For Each item In col
        s = s & Mid$(texto, item(0), item(1) - item(0) + 1) & "-"
    Next item
    If Len(s) > 0 Then s = Left$(s, Len(s) - 1)

    ' 3. Bucle de validación con formulario
    Do
        valido = True
        msg = ""

        s = RevisarSilabas_EnFormulario(texto, s)

        If s = "" Then
            Set Silabear_EU_ConRevision = col
            Exit Function
        End If

        If Left$(s, 1) = "-" Or Right$(s, 1) = "-" Then
            valido = False
            msg = "No puede empezar ni terminar con '-'."
        End If

        If InStr(s, "--") > 0 Then
            valido = False
            msg = "No puede haber sílabas vacías ('--')."
        End If

        Dim reconstruido As String
        Dim textoSinEspacios As String

        reconstruido = Replace(s, "-", "")
        reconstruido = Replace(reconstruido, " ", "")
        textoSinEspacios = Replace(texto, " ", "")

        If UCase$(reconstruido) <> UCase$(textoSinEspacios) Then
            valido = False
            msg = "Las sílabas no coinciden con el texto original."
        End If

        If Not valido Then
            MsgBox msg, vbExclamation, "Error en las sílabas"
        End If

    Loop Until valido

    ' 4. Reconstruir colección final
    partes = Split(s, "-")
    inicio = 1

    For Each p In partes
        fin = inicio + Len(p) - 1
        resultado.Add Array(inicio, fin)
        inicio = fin + 1
    Next p

    Set Silabear_EU_ConRevision = resultado

End Function

Public Function Silabear_EU(ByVal texto As String) As Collection

    Dim col As New Collection
    Dim i As Long, ini As Long
    Dim c1 As String, c2 As String, c3 As String

    texto = Trim$(texto)
    If Len(texto) = 0 Then
        Set Silabear_EU = col
        Exit Function
    End If

    ini = 1

    For i = 2 To Len(texto)

        c1 = Mid$(texto, i - 1, 1)
        c2 = Mid$(texto, i, 1)

        ' ---------------------------------------------------------
        ' 0. Espacios --> separan palabras
        ' ---------------------------------------------------------
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

        ' ---------------------------------------------------------
        ' 1. VV --> hiato (euskera no forma diptongos complejos)
        ' ---------------------------------------------------------
        If EsVocal_EU(c1) And EsVocal_EU(c2) Then
            col.Add Array(ini, i - 1)
            ini = i
            GoTo siguiente
        End If

        ' ---------------------------------------------------------
        ' 2. VCV --> V | CV
        ' ---------------------------------------------------------
        If EsVocal_EU(c1) And EsConsonant_EU(c2) Then
            If i < Len(texto) Then
                c3 = Mid$(texto, i + 1, 1)
                If EsVocal_EU(c3) Then
                    col.Add Array(ini, i - 1)
                    ini = i
                    GoTo siguiente
                End If
            End If
        End If

        ' ---------------------------------------------------------
        ' 3. CCV --> C | CV
        ' ---------------------------------------------------------
        If EsConsonant_EU(c1) And EsConsonant_EU(c2) Then
            If i < Len(texto) Then
                c3 = Mid$(texto, i + 1, 1)
                If EsVocal_EU(c3) Then
                    col.Add Array(ini, i - 1)
                    ini = i
                    GoTo siguiente
                End If
            End If
        End If

siguiente:
    Next i

    ' Última sílaba
    If ini <= Len(texto) Then
        col.Add Array(ini, Len(texto))
    End If

    Set Silabear_EU = col

End Function

Public Function EsVocal_EU(c As String) As Boolean
    EsVocal_EU = InStr("AEIOUaeiouÁÉÍÓÚáéíóú", c) > 0
End Function

Public Function EsConsonant_EU(c As String) As Boolean
    EsConsonant_EU = Not EsVocal_EU(c) And c <> " "
End Function


' ============================================================
'   ReglasEuskera (EUS)
'   Devuelve idFonema según la fonética del euskera.
'   Si no aplica, devuelve 0 para que el motor siga probando.
' ============================================================

Public Function ReglasEuskera( _
        ByVal graf As String, _
        ByVal ant As String, _
        ByVal sig As String, _
        ByVal esTonica As Boolean _
    ) As Byte

    Dim g As String
    g = UCase$(graf)

    ' ============================================================
    '   TRIGRAFEMAS
    ' ============================================================

    ' GÜE / GÜI --> /gw/ --> id 57 (préstamos)
    If g = "GÜE" Or g = "GÜI" Then
        ReglasEuskera = 57
        Exit Function
    End If

    ' GUE / GUI --> /g/ (U muda) --> id 31
    If g = "GUE" Or g = "GUI" Then
        ReglasEuskera = 31
        Exit Function
    End If

    ' QUE / QUI --> /k/ --> id 30
    If g = "QUE" Or g = "QUI" Then
        ReglasEuskera = 30
        Exit Function
    End If


    ' ============================================================
    '   DÍGRAFOS Y CASOS ESPECIALES
    ' ============================================================

    ' TX --> /t?/ --> id 50
    If g = "TX" Then
        ReglasEuskera = 50
        Exit Function
    End If

    ' TS / TZ --> /ts/ --> id 52
    If g = "TS" Or g = "TZ" Then
        ReglasEuskera = 52
        Exit Function
    End If

    ' LL --> /?/ --> id 44
    If g = "LL" Then
        ReglasEuskera = 44
        Exit Function
    End If

    ' RR --> /r/ múltiple --> id 46
    If g = "RR" Then
        ReglasEuskera = 46
        Exit Function
    End If

    ' Ñ --> /?/ --> id 41
    If g = "Ñ" Then
        ReglasEuskera = 41
        Exit Function
    End If


    ' ============================================================
    '   DÍGRAFOS VOCÁLICOS (diptongos euskera)
    ' ============================================================

    If g = "AI" Then ReglasEuskera = 12: Exit Function
    If g = "EI" Then ReglasEuskera = 13: Exit Function
    If g = "OI" Then ReglasEuskera = 14: Exit Function
    If g = "AU" Then ReglasEuskera = 16: Exit Function


    ' ============================================================
    '   MONÓGRAFOS — VOCALES (5 vocales)
    ' ============================================================

    If g = "A" Then ReglasEuskera = 1: Exit Function
    If g = "E" Then ReglasEuskera = 5: Exit Function
    If g = "I" Then ReglasEuskera = 9: Exit Function
    If g = "O" Then ReglasEuskera = 7: Exit Function
    If g = "U" Then ReglasEuskera = 10: Exit Function


    ' ============================================================
    '   MONÓGRAFOS — CONSONANTES
    ' ============================================================

    If g = "P" Then ReglasEuskera = 26: Exit Function
    If g = "B" Then ReglasEuskera = 27: Exit Function
    If g = "T" Then ReglasEuskera = 28: Exit Function
    If g = "D" Then ReglasEuskera = 29: Exit Function
    If g = "K" Then ReglasEuskera = 30: Exit Function
    If g = "G" Then ReglasEuskera = 31: Exit Function

    If g = "F" Then ReglasEuskera = 32: Exit Function

    ' S / Z --> /s/ (no existe /?/)
    If g = "S" Then ReglasEuskera = 34: Exit Function
    If g = "Z" Then ReglasEuskera = 34: Exit Function

    ' X --> /?/ --> id 36
    If g = "X" Then ReglasEuskera = 36: Exit Function

    ' J --> /j/ --> id 48
    If g = "J" Then ReglasEuskera = 48: Exit Function

    If g = "M" Then ReglasEuskera = 39: Exit Function
    If g = "N" Then ReglasEuskera = 40: Exit Function

    If g = "L" Then ReglasEuskera = 43: Exit Function
    If g = "R" Then ReglasEuskera = 45: Exit Function

    ' H --> aspiración suave --> id 38
    If g = "H" Then ReglasEuskera = 38: Exit Function


    ' ============================================================
    '   SI NO APLICA, DEVOLVER 0
    ' ============================================================
    ReglasEuskera = 0

End Function

'============================================================================================

Public Function MF_NormalizarVocales_EU(ByVal texto As String) As String

    ' A
    texto = Replace(texto, "Á", "A")
    texto = Replace(texto, "À", "A")

    ' E
    texto = Replace(texto, "É", "E")
    texto = Replace(texto, "È", "E")

    ' I
    texto = Replace(texto, "Í", "I")
    texto = Replace(texto, "Ì", "I")

    ' O
    texto = Replace(texto, "Ó", "O")
    texto = Replace(texto, "Ò", "O")

    ' U
    texto = Replace(texto, "Ú", "U")
    texto = Replace(texto, "Ù", "U")

    MF_NormalizarVocales_EU = texto

End Function

