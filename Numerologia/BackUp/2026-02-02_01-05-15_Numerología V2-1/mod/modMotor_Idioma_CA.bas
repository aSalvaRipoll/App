Attribute VB_Name = "modMotor_Idioma_CA"

Option Compare Database
Option Explicit

'=============
'== Catalán ==
'=============

Public Sub MF_MarcarTonica_CA( _
        ByVal Texto As String, _
        ByRef esTonica() As Boolean _
    )

    Dim Silabas As Collection
    Dim idxTonica As Long
    Dim inicio As Long, fin As Long
    Dim i As Long
    Dim vocalesTilde As String
    Dim ultima As String

    ' Vocales catalanas con acento (incluye diéresis)
    vocalesTilde = "ÀÈÉÍÏÒÓÚÜ"

    ' --------------------------------------------------------
    ' 1. Silabear palabra (motor con revisión)
    ' --------------------------------------------------------
    Set Silabas = Silabear_CA_ConRevision(Texto)

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
    ' 3. Si no hay tilde --> reglas catalanas
    ' --------------------------------------------------------
    If idxTonica = 0 Then

        ultima = Right$(Texto, 1)

        ' 3.1. Infinitivos catalanes (AR, ER, IR) --> oxítonos
        If Len(Texto) >= 2 Then
            Dim ult2 As String
            ult2 = Right$(Texto, 2)

            If ult2 = "AR" Or ult2 = "ER" Or ult2 = "IR" Then
                idxTonica = Silabas.Count
                GoTo Marcar
            End If
        End If

        ' 3.2. Palabras acabadas en -IG --> oxítonas
        If Len(Texto) >= 2 Then
            If Right$(Texto, 2) = "IG" Then
                idxTonica = Silabas.Count
                GoTo Marcar
            End If
        End If

        ' 3.3. Regla general catalana
        '     Oxítonas si terminan en:
        '     - vocal
        '     - vocal + S
        '     - EN, IN
        If InStr("AEIOU", ultima) > 0 Or _
           Right$(Texto, 2) = "AS" Or _
           Right$(Texto, 2) = "ES" Or _
           Right$(Texto, 2) = "IS" Or _
           Right$(Texto, 2) = "OS" Or _
           Right$(Texto, 2) = "US" Or _
           Right$(Texto, 2) = "EN" Or _
           Right$(Texto, 2) = "IN" Then

            idxTonica = Silabas.Count

        Else
            ' Paroxítona
            If Silabas.Count = 1 Then
                idxTonica = 1
            Else
                idxTonica = Silabas.Count - 1
            End If

        End If

    End If

Marcar:
    ' --------------------------------------------------------
    ' 4. Marcar índices tónicos
    ' --------------------------------------------------------
    inicio = Silabas(idxTonica)(1)
    fin = Silabas(idxTonica)(2)

    For i = inicio To fin
        esTonica(i) = True
    Next i

End Sub

Public Function Silabear_CA_ConRevision(ByVal Texto As String) As Collection

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
    ' 1. Silabear automáticamente (motor puro catalán)
    ' ============================================================
    Set col = Silabear_CA(Texto)

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
            Set Silabear_CA_ConRevision = col
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

    Set Silabear_CA_ConRevision = resultado

End Function

Public Function Silabear_CA(ByVal Texto As String) As Collection

    Dim col As New Collection
    Dim i As Long, ini As Long
    Dim c1 As String, c2 As String, c3 As String
    Dim par As String

    Texto = Trim$(Texto)
    If Len(Texto) = 0 Then
        Set Silabear_CA = col
        Exit Function
    End If

    ini = 1

    For i = 2 To Len(Texto)

        c1 = Mid$(Texto, i - 1, 1)
        c2 = Mid$(Texto, i, 1)
        par = c1 & c2

        ' =====================================================
        ' 0. ESPACIOS
        ' =====================================================
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

        ' =====================================================
        ' 1. ELA GEMINADA (L·L) --> se separa L-L
        ' =====================================================
        If par = "L·" Or par = "·L" Then
            col.Add Array(ini, i - 1)
            ini = i
            GoTo Siguiente
        End If

        ' =====================================================
        ' 2. GRUPOS CONSONÀNTICS INSEPARABLES
        ' =====================================================
        If EsConsonant_CA(c1) And EsConsonant_CA(c2) Then
            If EsGrupInseparable_CA(par) Then
                GoTo Siguiente
            End If
        End If

        ' =====================================================
        ' 3. CCV --> C | CV
        ' =====================================================
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

        ' =====================================================
        ' 4. VCV --> V | CV
        ' =====================================================
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

        ' =====================================================
        ' 5. VV --> hiato si hay vocal débil tónica (Í, Ú)
        ' =====================================================
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

    Set Silabear_CA = col

End Function



' ============================================================
'   Silabear_CAT — Silabeador para nombres y apellidos en catalán
'   - Respeta espacios entre palabras
'   - No mezcla sílabas entre palabras
'   - Aplica reglas fonéticas del catalán
'   - Trata L·L como separador obligatorio
'   - Detecta dígrafos y trígrafos catalanes
'   - No separa IG final (PUIG)
'   - Elimina H final aislada
' ============================================================



'' ============================================================
''   Silabear_CAT — Silabeador para nombres y apellidos en catalán
''   - Respeta espacios entre palabras
''   - No mezcla sílabas entre palabras
''   - Aplica reglas fonéticas del catalán
''   - Trata L·L como separador obligatorio
''   - Detecta diptongos catalanes
'' ============================================================



' ============================================================
'   ReglasCatala (CAT)
'   Devuelve idFonema según la fonética del catalán central.
'   Si no aplica, devuelve 0 para que el motor siga probando.
' ============================================================

Public Function ReglasCatala( _
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
        ReglasCatala = 57
        Exit Function
    End If

    ' GUE / GUI --> /g/ (U muda) --> id 31
    If g = "GUE" Or g = "GUI" Then
        ReglasCatala = 31
        Exit Function
    End If

    ' QUE / QUI --> /k/ --> id 30
    If g = "QUE" Or g = "QUI" Then
        ReglasCatala = 30
        Exit Function
    End If


    ' ============================================================
    '   DÍGRAFOS Y CASOS ESPECIALES
    ' ============================================================

    ' TX --> /t?/ --> id 50 (en catalán central)
    If g = "TX" Then
        ReglasCatala = 50
        Exit Function
    End If

    ' CH --> /t?/ --> id 50 (préstecs)
    If g = "CH" Then
        ReglasCatala = 50
        Exit Function
    End If

    ' NY --> /?/ --> id 41
    If g = "NY" Then
        ReglasCatala = 41
        Exit Function
    End If

    ' LL --> /?/ --> id 44
    If g = "LL" Then
        ReglasCatala = 44
        Exit Function
    End If

    ' L·L --> /l?/ --> id 61 (ela geminada)
    If g = "L·L" Or g = "L.L" Then
        ReglasCatala = 61
        Exit Function
    End If

    ' IX --> /?/ --> id 36
    If g = "IX" Then
        ReglasCatala = 36
        Exit Function
    End If

    ' TJ / TG --> /d?/ --> id 51
    If g = "TJ" Or g = "TG" Then
        ReglasCatala = 51
        Exit Function
    End If

    ' IG final --> /t?/ --> id 50
    If g = "IG" And sig = "" Then
        ReglasCatala = 50
        Exit Function
    End If


    ' ============================================================
    '   DÍGRAFOS VOCÁLICOS (diftongs catalans)
    ' ============================================================

    If g = "UA" Then ReglasCatala = 23: Exit Function
    If g = "UE" Then ReglasCatala = 24: Exit Function
    If g = "UO" Then ReglasCatala = 25: Exit Function

    If g = "IA" Then ReglasCatala = 20: Exit Function
    If g = "IE" Then ReglasCatala = 21: Exit Function
    If g = "IO" Then ReglasCatala = 22: Exit Function


    ' ============================================================
    '   MONÒGRAFS — VOCALS (7 vocals)
    ' ============================================================

    ' /a/
    If g = "A" Then
        ReglasCatala = 1
        Exit Function
    End If

    ' /i/
    If g = "I" Then
        ReglasCatala = 9
        Exit Function
    End If

    ' /u/
    If g = "U" Then
        ReglasCatala = 10
        Exit Function
    End If

    ' E tònica --> /?/ (id 6), àtona --> /e/ (id 5)
    If g = "E" Then
        If esTonica Then
            ReglasCatala = 6   ' /?/
        Else
            ReglasCatala = 5   ' /e/
        End If
        Exit Function
    End If

    ' O tònica --> /?/ (id 8), àtona --> /o/ (id 7)
    If g = "O" Then
        If esTonica Then
            ReglasCatala = 8   ' /?/
        Else
            ReglasCatala = 7   ' /o/
        End If
        Exit Function
    End If


    ' ============================================================
    '   MONÒGRAFS — CONSONANTS
    ' ============================================================

    If g = "P" Then ReglasCatala = 26: Exit Function
    If g = "B" Then ReglasCatala = 27: Exit Function
    If g = "T" Then ReglasCatala = 28: Exit Function
    If g = "D" Then ReglasCatala = 29: Exit Function
    If g = "K" Or g = "C" Then ReglasCatala = 30: Exit Function
    If g = "G" Then ReglasCatala = 31: Exit Function

    If g = "F" Then ReglasCatala = 32: Exit Function
    If g = "V" Then ReglasCatala = 33: Exit Function
    If g = "S" Then ReglasCatala = 34: Exit Function
    If g = "Z" Then ReglasCatala = 35: Exit Function
    If g = "J" Then ReglasCatala = 37: Exit Function

    If g = "M" Then ReglasCatala = 39: Exit Function
    If g = "N" Then ReglasCatala = 40: Exit Function

    If g = "L" Then ReglasCatala = 43: Exit Function
    If g = "R" Then ReglasCatala = 45: Exit Function

    If g = "H" Then ReglasCatala = 38: Exit Function


    ' ============================================================
    '   SI NO APLICA, RETORNAR 0
    ' ============================================================
    ReglasCatala = 0

End Function

'============================================================================================

Public Function MF_NormalizarVocales_CA(ByVal Texto As String) As String

    ' A
    Texto = Replace(Texto, "À", "A")
    Texto = Replace(Texto, "Á", "A")

    ' E
    Texto = Replace(Texto, "È", "E")
    Texto = Replace(Texto, "É", "E")

    ' I  (NO tocar Ï)
    Texto = Replace(Texto, "Í", "I")

    ' O
    Texto = Replace(Texto, "Ò", "O")
    Texto = Replace(Texto, "Ó", "O")

    ' U  (NO tocar Ü)
    Texto = Replace(Texto, "Ú", "U")

    MF_NormalizarVocales_CA = Texto

End Function

