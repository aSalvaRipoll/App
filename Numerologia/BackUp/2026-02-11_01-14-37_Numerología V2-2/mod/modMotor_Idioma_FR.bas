Attribute VB_Name = "modMotor_Idioma_FR"

Option Compare Database
Option Explicit

'=============
'== Francés ==
'=============

Public Sub MF_MarcarTonica_FR( _
        ByVal texto As String, _
        ByRef esTonica() As Boolean _
    )

    Dim silabas As Collection
    Dim idxTonica As Long
    Dim inicio As Long, Fin As Long
    Dim i As Long

    ' 1. Silabear palabra (motor con revisión)
    Set silabas = Silabear_FR_ConRevision(texto)

    If silabas Is Nothing Then Exit Sub
    If silabas.Count = 0 Then Exit Sub

    ' 2. Acento francés: SIEMPRE en la última sílaba fonética
    idxTonica = silabas.Count

    ' 3. Marcar índices tónicos
    inicio = silabas(idxTonica)(1)
    Fin = silabas(idxTonica)(2)

    For i = inicio To Fin
        esTonica(i) = True
    Next i

End Sub

Public Function Silabear_FR_ConRevision(ByVal texto As String) As Collection

    Dim col As Collection
    Dim resultado As New Collection
    Dim item As Variant
    Dim s As String
    Dim partes As Variant
    Dim p As Variant
    Dim inicio As Long, Fin As Long
    Dim valido As Boolean
    Dim msg As String

    ' 1. Silabear automáticamente (motor puro francés)
    Set col = Silabear_FR(texto)

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
            Set Silabear_FR_ConRevision = col
            Exit Function
        End If

        If Left$(s, 1) = "-" Or Right$(s, 1) = "-" Then
            valido = False
            msg = "Ne peut pas commencer ou finir par '-'."
        End If

        If InStr(s, "--") > 0 Then
            valido = False
            msg = "Ne peut pas contenir des syllabes vides ('--')."
        End If

        Dim reconstruido As String
        Dim textoSinEspacios As String

        reconstruido = Replace(s, "-", "")
        reconstruido = Replace(reconstruido, " ", "")
        textoSinEspacios = Replace(texto, " ", "")

        If UCase$(reconstruido) <> UCase$(textoSinEspacios) Then
            valido = False
            msg = "Les syllabes ne correspondent pas au texte original."
        End If

        If Not valido Then
            MsgBox msg, vbExclamation, "Erreur de syllabes"
        End If

    Loop Until valido

    ' 4. Reconstruir colección final
    partes = Split(s, "-")
    inicio = 1

    For Each p In partes
        Fin = inicio + Len(p) - 1
        resultado.Add Array(inicio, Fin)
        inicio = Fin + 1
    Next p

    Set Silabear_FR_ConRevision = resultado

End Function

Public Function Silabear_FR(ByVal texto As String) As Collection

    Dim col As New Collection
    Dim i As Long, ini As Long
    Dim c1 As String, c2 As String, c3 As String
    Dim par As String

    texto = Trim$(texto)
    If Len(texto) = 0 Then
        Set Silabear_FR = col
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
        ' 1. VV --> hiato (francés no forma diptongos ortográficos)
        ' ---------------------------------------------------------
        If EsVocal_FR(c1) And EsVocal_FR(c2) Then
            col.Add Array(ini, i - 1)
            ini = i
            GoTo siguiente
        End If

        ' ---------------------------------------------------------
        ' 2. VCV --> V | CV
        ' ---------------------------------------------------------
        If EsVocal_FR(c1) And EsConsonant_FR(c2) Then
            If i < Len(texto) Then
                c3 = Mid$(texto, i + 1, 1)
                If EsVocal_FR(c3) Then
                    col.Add Array(ini, i - 1)
                    ini = i
                    GoTo siguiente
                End If
            End If
        End If

        ' ---------------------------------------------------------
        ' 3. CCV --> C | CV
        ' ---------------------------------------------------------
        If EsConsonant_FR(c1) And EsConsonant_FR(c2) Then
            If i < Len(texto) Then
                c3 = Mid$(texto, i + 1, 1)
                If EsVocal_FR(c3) Then
                    If Not EsGrupInseparable_FR(par) Then
                        col.Add Array(ini, i - 1)
                        ini = i
                        GoTo siguiente
                    End If
                End If
            End If
        End If

siguiente:
    Next i

    ' Última sílaba
    If ini <= Len(texto) Then
        col.Add Array(ini, Len(texto))
    End If

    Set Silabear_FR = col

End Function

Public Function EsVocal_FR(c As String) As Boolean
    EsVocal_FR = InStr("AEIOUYÀÂÄÉÈÊËÎÏÔÖÙÛÜaeiouyàâäéèêëîïôöùûü", c) > 0
End Function

Public Function EsConsonant_FR(c As String) As Boolean
    EsConsonant_FR = Not EsVocal_FR(c) And c <> " "
End Function

Public Function EsGrupInseparable_FR(par As String) As Boolean
    Select Case UCase$(par)
        Case "BR", "BL", "CR", "CL", "DR", "TR", "PR", "PL", "GR", "GL", "FR", "FL"
            EsGrupInseparable_FR = True
        Case Else
            EsGrupInseparable_FR = False
    End Select
End Function


' ============================================================
'   ReglasFrances (FR)
'   Devuelve idFonema según la fonética del francés.
'   Si no aplica, devuelve 0 para que el motor siga probando.
' ============================================================

Public Function ReglasFrances( _
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
        ReglasFrances = 57
        Exit Function
    End If

    ' GUE / GUI --> /g/ --> id 31
    If g = "GUE" Or g = "GUI" Then
        ReglasFrances = 31
        Exit Function
    End If

    ' QUE / QUI --> /k/ --> id 30
    If g = "QUE" Or g = "QUI" Then
        ReglasFrances = 30
        Exit Function
    End If


    ' ============================================================
    '   DÍGRAFOS Y CASOS ESPECIALES
    ' ============================================================

    ' CH --> /?/ --> id 36
    If g = "CH" Then
        ReglasFrances = 36
        Exit Function
    End If

    ' GN --> /?/ --> id 41
    If g = "GN" Then
        ReglasFrances = 41
        Exit Function
    End If

    ' J --> /?/ --> id 37
    If g = "J" Then
        ReglasFrances = 37
        Exit Function
    End If

    ' G + E/I/Y --> /?/ --> id 37
    If g = "G" And (sig = "E" Or sig = "I" Or sig = "Y") Then
        ReglasFrances = 37
        Exit Function
    End If

    ' S entre vocales --> /z/ --> id 35
    If g = "S" And (ant Like "[AEIOU]" And sig Like "[AEIOU]") Then
        ReglasFrances = 35
        Exit Function
    End If

    ' Ç --> /s/ --> id 34
    If g = "Ç" Then
        ReglasFrances = 34
        Exit Function
    End If


    ' ============================================================
    '   NASALIZACIONES
    ' ============================================================

    ' AN / AM / EN / EM --> /?~/ --> id 2
    If g = "AN" Or g = "AM" Or g = "EN" Or g = "EM" Then
        ReglasFrances = 2
        Exit Function
    End If

    ' IN / IM / AIN / EIN / EIM / YN / YM --> /?~/ --> id 3
    If g = "IN" Or g = "IM" Or g = "AIN" Or g = "EIN" Or g = "EIM" Or g = "YN" Or g = "YM" Then
        ReglasFrances = 3
        Exit Function
    End If

    ' ON / OM --> /?~/ --> id 4
    If g = "ON" Or g = "OM" Then
        ReglasFrances = 4
        Exit Function
    End If

    ' UN / UM --> /œ~/ --> id 3 (aproximación razonable)
    If g = "UN" Or g = "UM" Then
        ReglasFrances = 3
        Exit Function
    End If


    ' ============================================================
    '   DÍGRAFOS VOCÁLICOS (diptongos franceses)
    ' ============================================================

    ' OI --> /wa/ --> id 18
    If g = "OI" Then ReglasFrances = 18: Exit Function

    ' AI --> /?/ --> id 6
    If g = "AI" Then ReglasFrances = 6: Exit Function

    ' EI --> /e/ --> id 5
    If g = "EI" Then ReglasFrances = 5: Exit Function

    ' OU --> /u/ --> id 10
    If g = "OU" Then ReglasFrances = 10: Exit Function


    ' ============================================================
    '   MONÓGRAFOS — VOCALES
    ' ============================================================

    If g = "A" Then ReglasFrances = 1: Exit Function
    If g = "E" Then ReglasFrances = 5: Exit Function
    If g = "I" Then ReglasFrances = 9: Exit Function
    If g = "O" Then ReglasFrances = 7: Exit Function
    If g = "U" Then ReglasFrances = 10: Exit Function


    ' ============================================================
    '   MONÓGRAFOS — CONSONANTES
    ' ============================================================

    If g = "P" Then ReglasFrances = 26: Exit Function
    If g = "B" Then ReglasFrances = 27: Exit Function
    If g = "T" Then ReglasFrances = 28: Exit Function
    If g = "D" Then ReglasFrances = 29: Exit Function
    If g = "K" Then ReglasFrances = 30: Exit Function
    If g = "G" Then ReglasFrances = 31: Exit Function

    If g = "F" Then ReglasFrances = 32: Exit Function
    If g = "V" Then ReglasFrances = 33: Exit Function

    ' C + E/I/Y --> /s/
    If g = "C" And (sig = "E" Or sig = "I" Or sig = "Y") Then
        ReglasFrances = 34
        Exit Function
    End If

    ' S --> /s/
    If g = "S" Then ReglasFrances = 34: Exit Function

    ' X --> /ks/ o /gz/ --> simplificamos a /s/ (el motor segmenta la K aparte)
    If g = "X" Then
        ReglasFrances = 34
        Exit Function
    End If

    If g = "M" Then ReglasFrances = 39: Exit Function
    If g = "N" Then ReglasFrances = 40: Exit Function

    ' L --> /l/
    If g = "L" Then ReglasFrances = 43: Exit Function

    ' R --> /?/ --> id 47
    If g = "R" Then
        ReglasFrances = 47
        Exit Function
    End If

    ' H --> muda --> id 38
    If g = "H" Then
        ReglasFrances = 38
        Exit Function
    End If


    ' ============================================================
    '   SI NO APLICA, DEVOLVER 0
    ' ============================================================
    ReglasFrances = 0

End Function

'============================================================================================

Public Function MF_NormalizarVocales_FR(ByVal texto As String) As String

    ' A
    texto = Replace(texto, "À", "A")   ' abierta
    texto = Replace(texto, "Á", "A")   ' rara, pero robustez
    texto = Replace(texto, "Â", "Â")   ' cerrada
    texto = Replace(texto, "Ä", "A¨")  ' hiato

    ' E
    texto = Replace(texto, "È", "E")   ' abierta
    texto = Replace(texto, "É", "E´")  ' cerrada
    texto = Replace(texto, "Ê", "Ê")   ' cerrada tensa
    texto = Replace(texto, "Ë", "E¨")  ' hiato

    ' I
    texto = Replace(texto, "Ì", "I")   ' robustez
    texto = Replace(texto, "Í", "I")   ' robustez
    texto = Replace(texto, "Î", "Î")   ' cerrada
    texto = Replace(texto, "Ï", "I¨")  ' hiato

    ' O
    texto = Replace(texto, "Ò", "O")   ' robustez
    texto = Replace(texto, "Ó", "O")   ' robustez
    texto = Replace(texto, "Ô", "Ô")   ' cerrada
    texto = Replace(texto, "Ö", "O¨")  ' hiato

    ' U
    texto = Replace(texto, "Ù", "U")   ' abierta
    texto = Replace(texto, "Ú", "U")   ' robustez
    texto = Replace(texto, "Û", "Û")   ' cerrada
    texto = Replace(texto, "Ü", "U¨")  ' hiato

    MF_NormalizarVocales_FR = texto

End Function

