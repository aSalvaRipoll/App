Attribute VB_Name = "modMotor_Idioma_CA_VA"

Option Compare Database
Option Explicit

'================
'== Valenciano ==
'================

Public Sub MF_MarcarTonica_CA_VA( _
        ByVal texto As String, _
        ByRef esTonica() As Boolean _
    )

    Dim silabas As Collection
    Dim idxTonica As Long
    Dim inicio As Long, fin As Long
    Dim i As Long
    Dim vocalesTilde As String
    Dim ultima As String

    ' Vocales catalanas/valencianas con acento (incluye diéresis)
    vocalesTilde = "ÀÈÉÍÏÒÓÚÜ"

    ' --------------------------------------------------------
    ' 1. Silabear palabra (motor con revisión)
    ' --------------------------------------------------------
    Set silabas = Silabear_CA_VA_ConRevision(texto)

    If silabas Is Nothing Then Exit Sub
    If silabas.Count = 0 Then Exit Sub

    ' --------------------------------------------------------
    ' 2. Buscar vocal con acento
    ' --------------------------------------------------------
    For i = 1 To Len(texto)
        If InStr(vocalesTilde, Mid$(texto, i, 1)) > 0 Then
'            idxTonica = MF_SilabaDeIndice(i, Silabas)
            Exit For
        End If
    Next i

    ' --------------------------------------------------------
    ' 3. Si no hay tilde --> reglas valencianas
    ' --------------------------------------------------------
    If idxTonica = 0 Then

        ultima = Right$(texto, 1)

        ' 3.1. Infinitivos (AR, ER, IR) --> oxítonos
        If Len(texto) >= 2 Then
            Dim ult2 As String
            ult2 = Right$(texto, 2)

            If ult2 = "AR" Or ult2 = "ER" Or ult2 = "IR" Then
                idxTonica = silabas.Count
                GoTo Marcar
            End If
        End If

        ' 3.2. Palabras acabadas en -IG --> oxítonas
        If Len(texto) >= 2 Then
            If Right$(texto, 2) = "IG" Then
                idxTonica = silabas.Count
                GoTo Marcar
            End If
        End If

        ' 3.3. Regla general valenciana
        '     Oxítonas si terminan en:
        '     - vocal
        '     - vocal + S
        '     - EN, IN
        If InStr("AEIOU", ultima) > 0 Or _
           Right$(texto, 2) = "AS" Or _
           Right$(texto, 2) = "ES" Or _
           Right$(texto, 2) = "IS" Or _
           Right$(texto, 2) = "OS" Or _
           Right$(texto, 2) = "US" Or _
           Right$(texto, 2) = "EN" Or _
           Right$(texto, 2) = "IN" Then

            idxTonica = silabas.Count

        Else
            ' Paroxítona
            If silabas.Count = 1 Then
                idxTonica = 1
            Else
                idxTonica = silabas.Count - 1
            End If

        End If

    End If

Marcar:
    ' --------------------------------------------------------
    ' 4. Marcar índices tónicos
    ' --------------------------------------------------------
    inicio = silabas(idxTonica)(1)
    fin = silabas(idxTonica)(2)

    For i = inicio To fin
        esTonica(i) = True
    Next i

End Sub

Public Function Silabear_CA_VA_ConRevision(ByVal texto As String) As Collection

    Dim col As Collection
    Dim resultado As New Collection
    Dim item As Variant
    Dim s As String
    Dim partes As Variant
    Dim p As Variant
    Dim inicio As Long, fin As Long
    Dim valido As Boolean
    Dim msg As String

    ' ============================================================
    ' 1. Silabear automáticamente (motor puro valenciano)
    ' ============================================================
    Set col = Silabear_CA_VA(texto)

    ' 2. Convertir a string con separador "-"
    For Each item In col
        s = s & Mid$(texto, item(0), item(1) - item(0) + 1) & "-"
    Next item
    If Len(s) > 0 Then s = Left$(s, Len(s) - 1)

    ' ============================================================
    ' 3. Bucle de validación con formulario
    ' ============================================================
    Do
        valido = True
        msg = ""

        ' Abrir formulario de revisión
        s = RevisarSilabas_EnFormulario(texto, s)

        ' Si el usuario cancela --> devolver silabeo automático
        If s = "" Then
            Set Silabear_CA_VA_ConRevision = col
            Exit Function
        End If

        ' Validación 1: no puede empezar ni acabar con "-"
        If Left$(s, 1) = "-" Or Right$(s, 1) = "-" Then
            valido = False
            msg = "No puede empezar ni terminar con '-'."
        End If

        ' Validación 2: no puede contener "--"
        If InStr(s, "--") > 0 Then
            valido = False
            msg = "No puede haber sílabas vacías ('--')."
        End If

        ' Validación 3: reconstrucción del texto original
        Dim reconstruido As String
        Dim textoSinEspacios As String

        reconstruido = Replace(s, "-", "")
        reconstruido = Replace(reconstruido, " ", "")

        textoSinEspacios = Replace(texto, " ", "")

        If UCase$(reconstruido) <> UCase$(textoSinEspacios) Then
            valido = False
            msg = "Las sílabas no coinciden con el texto original (ignorando espacios)."
        End If

        If Not valido Then
            MsgBox msg, vbExclamation, "Error en las sílabas"
        End If

    Loop Until valido

    ' ============================================================
    ' 4. Reconstruir colección válida
    ' ============================================================
    partes = Split(s, "-")
    inicio = 1

    For Each p In partes
        fin = inicio + Len(p) - 1
        resultado.Add Array(inicio, fin)
        inicio = fin + 1
    Next p

    Set Silabear_CA_VA_ConRevision = resultado

End Function


Public Function Silabear_CA_VA(ByVal texto As String) As Collection

    Dim col As New Collection
    Dim i As Long, ini As Long
    Dim c1 As String, c2 As String, c3 As String
    Dim par As String

    texto = Trim$(texto)
    If Len(texto) = 0 Then
        Set Silabear_CA_VA = col
        Exit Function
    End If

    ini = 1

    For i = 2 To Len(texto)

        c1 = Mid$(texto, i - 1, 1)
        c2 = Mid$(texto, i, 1)
        par = c1 & c2

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
        ' 1. Ela geminada (L·L)
        ' ---------------------------------------------------------
        If par = "L·" Or par = "·L" Then
            col.Add Array(ini, i - 1)
            ini = i
            GoTo siguiente
        End If

        ' ---------------------------------------------------------
        ' 2. Grupos consonánticos inseparables
        ' ---------------------------------------------------------
        If EsConsonant_CA(c1) And EsConsonant_CA(c2) Then
            If EsGrupInseparable_CA(par) Then
                GoTo siguiente
            End If
        End If

        ' ---------------------------------------------------------
        ' 3. CCV --> C | CV
        ' ---------------------------------------------------------
        If EsConsonant_CA(c1) And EsConsonant_CA(c2) Then
            If i < Len(texto) Then
                c3 = Mid$(texto, i + 1, 1)
                If EsVocal_CA(c3) Then
                    If Not EsGrupInseparable_CA(par) Then
                        col.Add Array(ini, i - 1)
                        ini = i
                        GoTo siguiente
                    End If
                End If
            End If
        End If

        ' ---------------------------------------------------------
        ' 4. VCV --> V | CV
        ' ---------------------------------------------------------
        If EsVocal_CA(c1) And EsConsonant_CA(c2) Then
            If i < Len(texto) Then
                c3 = Mid$(texto, i + 1, 1)
                If EsVocal_CA(c3) Then
                    col.Add Array(ini, i - 1)
                    ini = i
                    GoTo siguiente
                End If
            End If
        End If

        ' ---------------------------------------------------------
        ' 5. VV --> hiato si hay vocal débil tónica (Í, Ú)
        ' ---------------------------------------------------------
        If EsVocal_CA(c1) And EsVocal_CA(c2) Then
            If c1 = "Í" Or c1 = "Ú" Or c2 = "Í" Or c2 = "Ú" Then
                col.Add Array(ini, i - 1)
                ini = i
                GoTo siguiente
            End If
        End If

siguiente:
    Next i

    ' Última sílaba
    If ini <= Len(texto) Then
        col.Add Array(ini, Len(texto))
    End If

    Set Silabear_CA_VA = col

End Function



' ============================================================
'   ReglasValenciano (VAL)
'   Devuelve idFonema según la fonética valenciana.
'   Si no aplica, devuelve 0 para que el motor siga probando.
' ============================================================

Public Function ReglasValenciano( _
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

    ' GÜE / GÜI --> /gw/ --> id 57
    If g = "GÜE" Or g = "GÜI" Then
        ReglasValenciano = 57
        Exit Function
    End If

    ' GUE / GUI --> /g/ (U muda) --> id 31
    If g = "GUE" Or g = "GUI" Then
        ReglasValenciano = 31
        Exit Function
    End If

    ' QUE / QUI --> /k/ --> id 30
    If g = "QUE" Or g = "QUI" Then
        ReglasValenciano = 30
        Exit Function
    End If


    ' ============================================================
    '   DÍGRAFOS Y CASOS ESPECIALES
    ' ============================================================

    ' TX --> /t?/ --> id 50 (en valenciano no existe /t?/)
    If g = "TX" Then
        ReglasValenciano = 50
        Exit Function
    End If

    ' CH --> /t?/ --> id 50
    If g = "CH" Then
        ReglasValenciano = 50
        Exit Function
    End If

    ' NY --> /?/ --> id 41
    If g = "NY" Then
        ReglasValenciano = 41
        Exit Function
    End If

    ' LL --> /?/ --> id 44
    If g = "LL" Then
        ReglasValenciano = 44
        Exit Function
    End If

    ' L·L --> /l?/ --> id 61
    If g = "L·L" Or g = "L.L" Then
        ReglasValenciano = 61
        Exit Function
    End If

    ' IX --> /?/ --> id 36
    If g = "IX" Then
        ReglasValenciano = 36
        Exit Function
    End If

    ' TJ / TG --> /d?/ --> id 51
    If g = "TJ" Or g = "TG" Then
        ReglasValenciano = 51
        Exit Function
    End If

    ' IG final --> /t?/ --> id 50
    If g = "IG" And sig = "" Then
        ReglasValenciano = 50
        Exit Function
    End If


    ' ============================================================
    '   DÍGRAFOS VOCÁLICOS (diptongos valencianos)
    ' ============================================================

    If g = "UA" Then ReglasValenciano = 23: Exit Function
    If g = "UE" Then ReglasValenciano = 24: Exit Function
    If g = "UO" Then ReglasValenciano = 25: Exit Function

    If g = "IA" Then ReglasValenciano = 20: Exit Function
    If g = "IE" Then ReglasValenciano = 21: Exit Function
    If g = "IO" Then ReglasValenciano = 22: Exit Function


    ' ============================================================
    '   MONÓGRAFOS — VOCALES (5 vocales)
    ' ============================================================

    If g = "A" Then ReglasValenciano = 1: Exit Function
    If g = "E" Then ReglasValenciano = 5: Exit Function
    If g = "I" Then ReglasValenciano = 9: Exit Function
    If g = "O" Then ReglasValenciano = 7: Exit Function
    If g = "U" Then ReglasValenciano = 10: Exit Function


    ' ============================================================
    '   MONÓGRAFOS — CONSONANTES
    ' ============================================================

    If g = "P" Then ReglasValenciano = 26: Exit Function
    If g = "B" Then ReglasValenciano = 27: Exit Function
    If g = "T" Then ReglasValenciano = 28: Exit Function
    If g = "D" Then ReglasValenciano = 29: Exit Function
    If g = "K" Or _
       g = "C" Then ReglasValenciano = 30: Exit Function
    If g = "G" Then ReglasValenciano = 31: Exit Function

    If g = "F" Then ReglasValenciano = 32: Exit Function
    If g = "V" Then ReglasValenciano = 33: Exit Function
    If g = "S" Then ReglasValenciano = 34: Exit Function
    If g = "Z" Then ReglasValenciano = 35: Exit Function
    If g = "J" Then ReglasValenciano = 37: Exit Function

    If g = "M" Then ReglasValenciano = 39: Exit Function
    If g = "N" Then ReglasValenciano = 40: Exit Function

    If g = "L" Then ReglasValenciano = 43: Exit Function
    If g = "R" Then ReglasValenciano = 45: Exit Function

    If g = "H" Then ReglasValenciano = 38: Exit Function


    ' ============================================================
    '   SI NO APLICA, DEVOLVER 0
    ' ============================================================
    ReglasValenciano = 0

End Function

'============================================================================================

Public Function MF_NormalizarVocales_CA_VA(ByVal texto As String) As String
    MF_NormalizarVocales_CA_VA = MF_NormalizarVocales_CA(texto)
End Function

