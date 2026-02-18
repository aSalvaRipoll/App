Attribute VB_Name = "modMotor_Idioma_CA_IB"

Option Compare Database
Option Explicit

'================
'== Mallorquín ==
'================

Public Sub MF_MarcarTonica_CA_IB( _
        ByVal Texto As String, _
        ByRef esTonica() As Boolean _
    )

    Dim Silabas As Collection
    Dim idxTonica As Long
    Dim inicio As Long, fin As Long
    Dim i As Long
    Dim vocalesTilde As String

    ' Vocales catalanas con acento (incluye diéresis)
    vocalesTilde = "ÀÈÉÍÏÒÓÚÜ"

    ' --------------------------------------------------------
    ' 1. Silabear palabra (motor con revisión)
    ' --------------------------------------------------------
    Set Silabas = Silabear_CA_IB_ConRevision(Texto)

    If Silabas Is Nothing Then Exit Sub
    If Silabas.Count = 0 Then Exit Sub

    ' --------------------------------------------------------
    ' 2. Buscar vocal con acento
    ' --------------------------------------------------------
    For i = 1 To Len(Texto)
        If InStr(vocalesTilde, Mid$(Texto, i, 1)) > 0 Then
            idxTonica = MF_SilabaDeIndice(i, Silabas)
            Exit For
        End If
    Next i

    ' --------------------------------------------------------
    ' 3. Si no hay tilde --> penúltima (mallorquín)
    ' --------------------------------------------------------
    If idxTonica = 0 Then
        If Silabas.Count = 1 Then
            idxTonica = 1
        Else
            idxTonica = Silabas.Count - 1
        End If
    End If

    ' --------------------------------------------------------
    ' 4. Marcar índices tónicos
    ' --------------------------------------------------------
    inicio = Silabas(idxTonica)(1)
    fin = Silabas(idxTonica)(2)

    For i = inicio To fin
        esTonica(i) = True
    Next i

End Sub

Public Function Silabear_CA_IB_ConRevision(ByVal Texto As String) As Collection

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
    ' 1. Silabear automáticamente (motor puro mallorquín)
    ' ============================================================
    Set col = Silabear_CA_IB(Texto)

    ' 2. Convertir a string con separador "-"
    For Each item In col
        s = s & Mid$(Texto, item(0), item(1) - item(0) + 1) & "-"
    Next item
    If Len(s) > 0 Then s = Left$(s, Len(s) - 1)

    ' ============================================================
    ' 3. Bucle de validación con formulario
    ' ============================================================
    Do
        valido = True
        msg = ""

        ' Abrir formulario de revisión
        s = RevisarSilabas_EnFormulario(Texto, s)

        ' Si el usuario cancela --> devolver silabeo automático
        If s = "" Then
            Set Silabear_CA_IB_ConRevision = col
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

        textoSinEspacios = Replace(Texto, " ", "")

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

    Set Silabear_CA_IB_ConRevision = resultado

End Function

Public Function Silabear_CA_IB(ByVal Texto As String) As Collection

    Dim col As New Collection
    Dim i As Long, ini As Long
    Dim c1 As String, c2 As String, c3 As String
    Dim par As String

    Texto = Trim$(Texto)
    If Len(Texto) = 0 Then
        Set Silabear_CA_IB = col
        Exit Function
    End If

    ini = 1

    For i = 2 To Len(Texto)

        c1 = Mid$(Texto, i - 1, 1)
        c2 = Mid$(Texto, i, 1)
        par = c1 & c2

        ' ---------------------------------------------------------
        ' 0. Espacios --> separan palabras
        ' ---------------------------------------------------------
        If c1 = " " Then
            If i - 2 >= ini Then col.Add Array(ini, i - 2)
            ini = i
            GoTo Siguiente
        End If

        If c2 = " " Then
            col.Add Array(ini, i - 1)
            ini = i + 1
            GoTo Siguiente
        End If

        ' ---------------------------------------------------------
        ' 1. Ela geminada (L·L)
        ' ---------------------------------------------------------
        If par = "L·" Or par = "·L" Then
            col.Add Array(ini, i - 1)
            ini = i
            GoTo Siguiente
        End If

        ' ---------------------------------------------------------
        ' 2. Grupos consonánticos inseparables
        ' ---------------------------------------------------------
        If EsConsonant_CA(c1) And EsConsonant_CA(c2) Then
            If EsGrupInseparable_CA(par) Then
                GoTo Siguiente
            End If
        End If

        ' ---------------------------------------------------------
        ' 3. CCV --> C | CV
        ' ---------------------------------------------------------
        If EsConsonant_CA(c1) And EsConsonant_CA(c2) Then
            If i < Len(Texto) Then
                c3 = Mid$(Texto, i + 1, 1)
                If EsVocal_CA(c3) Then
                    If Not EsGrupInseparable_CA(par) Then
                        col.Add Array(ini, i - 1)
                        ini = i
                        GoTo Siguiente
                    End If
                End If
            End If
        End If

        ' ---------------------------------------------------------
        ' 4. VCV --> V | CV
        ' ---------------------------------------------------------
        If EsVocal_CA(c1) And EsConsonant_CA(c2) Then
            If i < Len(Texto) Then
                c3 = Mid$(Texto, i + 1, 1)
                If EsVocal_CA(c3) Then
                    col.Add Array(ini, i - 1)
                    ini = i
                    GoTo Siguiente
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
                GoTo Siguiente
            End If
        End If

Siguiente:
    Next i

    ' Última sílaba
    If ini <= Len(Texto) Then
        col.Add Array(ini, Len(Texto))
    End If

    Set Silabear_CA_IB = col

End Function


' ============================================================
'   ReglasMallorquin (CA-IB)
'   Devuelve idFonema según la fonética mallorquina.
'   Si no aplica, devuelve 0 para que el motor siga probando.
' ============================================================

Public Function ReglasMallorquin( _
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
        ReglasMallorquin = 57
        Exit Function
    End If

    ' GUE / GUI --> /g/ (U muda) --> id 31
    If g = "GUE" Or g = "GUI" Then
        ReglasMallorquin = 31
        Exit Function
    End If

    ' QUE / QUI --> /k/ --> id 30
    If g = "QUE" Or g = "QUI" Then
        ReglasMallorquin = 30
        Exit Function
    End If


    ' ============================================================
    '   DÍGRAFOS Y CASOS ESPECIALES
    ' ============================================================

    ' TX --> /t?/ --> id 60 (mallorquín)
    If g = "TX" Then
        ReglasMallorquin = 60
        Exit Function
    End If

    ' CH --> /t?/ --> id 50 (préstamos)
    If g = "CH" Then
        ReglasMallorquin = 50
        Exit Function
    End If

    ' NY --> /?/ --> id 41
    If g = "NY" Then
        ReglasMallorquin = 41
        Exit Function
    End If

    ' LL --> /?/ --> id 44
    If g = "LL" Then
        ReglasMallorquin = 44
        Exit Function
    End If

    ' L·L --> /l?/ --> id 61 (ela geminada)
    If g = "L·L" Or g = "L.L" Then
        ReglasMallorquin = 61
        Exit Function
    End If

    ' IX --> /?/ --> id 36
    If g = "IX" Then
        ReglasMallorquin = 36
        Exit Function
    End If

    ' TJ / TG --> /d?/ --> id 51
    If g = "TJ" Or g = "TG" Then
        ReglasMallorquin = 51
        Exit Function
    End If

    ' IG final --> /t?/ --> id 50
    If g = "IG" And sig = "" Then
        ReglasMallorquin = 50
        Exit Function
    End If


    ' ============================================================
    '   DÍGRAFOS VOCÁLICOS (diptongos mallorquines)
    ' ============================================================

    If g = "UA" Then ReglasMallorquin = 23: Exit Function
    If g = "UE" Then ReglasMallorquin = 24: Exit Function
    If g = "UO" Then ReglasMallorquin = 25: Exit Function

    If g = "IA" Then ReglasMallorquin = 20: Exit Function
    If g = "IE" Then ReglasMallorquin = 21: Exit Function
    If g = "IO" Then ReglasMallorquin = 22: Exit Function


    ' ============================================================
    '   MONÓGRAFOS — VOCALES
    ' ============================================================

    ' Vocal neutra (schwa) en sílaba átona --> /?/ --> id 11
    If Not esTonica Then
        If g = "A" Or g = "E" Or g = "O" Then
            ReglasMallorquin = 11
            Exit Function
        End If
    End If

    ' Vocales tónicas básicas
    If g = "A" Then ReglasMallorquin = 1: Exit Function
    If g = "I" Then ReglasMallorquin = 9: Exit Function
    If g = "U" Then ReglasMallorquin = 10: Exit Function

    ' E tónica --> abierta /?/ (id 6), átona --> cerrada /e/ (id 5)
    If g = "E" Then
        If esTonica Then
            ReglasMallorquin = 6
        Else
            ReglasMallorquin = 5
        End If
        Exit Function
    End If

    ' O tónica --> abierta /?/ (id 8), átona --> cerrada /o/ (id 7)
    If g = "O" Then
        If esTonica Then
            ReglasMallorquin = 8
        Else
            ReglasMallorquin = 7
        End If
        Exit Function
    End If


    ' ============================================================
    '   MONÓGRAFOS — CONSONANTES
    ' ============================================================

    If g = "P" Then ReglasMallorquin = 26: Exit Function
    If g = "B" Then ReglasMallorquin = 27: Exit Function
    If g = "T" Then ReglasMallorquin = 28: Exit Function
    If g = "D" Then ReglasMallorquin = 29: Exit Function
    If g = "K" Or g = "C" Then ReglasMallorquin = 30: Exit Function
    If g = "G" Then ReglasMallorquin = 31: Exit Function

    If g = "F" Then ReglasMallorquin = 32: Exit Function
    If g = "V" Then ReglasMallorquin = 33: Exit Function
    If g = "S" Then ReglasMallorquin = 34: Exit Function
    If g = "Z" Then ReglasMallorquin = 35: Exit Function
    If g = "J" Then ReglasMallorquin = 37: Exit Function

    If g = "M" Then ReglasMallorquin = 39: Exit Function
    If g = "N" Then ReglasMallorquin = 40: Exit Function

    If g = "L" Then ReglasMallorquin = 43: Exit Function
    If g = "R" Then ReglasMallorquin = 45: Exit Function

    If g = "H" Then ReglasMallorquin = 38: Exit Function


    ' ============================================================
    '   SI NO APLICA, DEVOLVER 0
    ' ============================================================
    ReglasMallorquin = 0

End Function

'============================================================================================

Public Function MF_NormalizarVocales_CA_IB(ByVal Texto As String) As String
    MF_NormalizarVocales_CA_IB = MF_NormalizarVocales_CA(Texto)
End Function

