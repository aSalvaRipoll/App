Attribute VB_Name = "modMotor_Idioma_FR"

Option Compare Database
Option Explicit

'=============
'== Francés ==
'=============

Public Sub MF_MarcarTonicaFR( _
        ByVal Texto As String, _
        ByRef esTonica() As Boolean _
    )

    Dim Silabas As Collection
    Dim idxTonica As Long
    Dim inicio As Long, fin As Long
    Dim i As Long

    ' 1. Silabear palabra (motor con revisión)
    Set Silabas = Silabear_FR_ConRevision(Texto)

    If Silabas Is Nothing Then Exit Sub
    If Silabas.Count = 0 Then Exit Sub

    ' 2. Acento francés: SIEMPRE en la última sílaba fonética
    idxTonica = Silabas.Count

    ' 3. Marcar índices tónicos
    inicio = Silabas(idxTonica)(1)
    fin = Silabas(idxTonica)(2)

    For i = inicio To fin
        esTonica(i) = True
    Next i

End Sub

Public Function Silabear_FR_ConRevision(ByVal Texto As String) As Collection

    Dim col As Collection
    Dim resultado As New Collection
    Dim item As Variant
    Dim s As String
    Dim partes As Variant
    Dim p As Variant
    Dim inicio As Long, fin As Long
    Dim valido As Boolean
    Dim msg As String

    ' 1. Silabear automáticamente (motor puro francés)
    Set col = Silabear_FR(Texto)

    ' 2. Convertir a string con "-"
    For Each item In col
        s = s & Mid$(Texto, item(0), item(1) - item(0) + 1) & "-"
    Next item
    If Len(s) > 0 Then s = Left$(s, Len(s) - 1)

    ' 3. Bucle de validación con formulario
    Do
        valido = True
        msg = ""

        s = RevisarSilabas_EnFormulario(Texto, s)

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
        textoSinEspacios = Replace(Texto, " ", "")

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
        fin = inicio + Len(p) - 1
        resultado.Add Array(inicio, fin)
        inicio = fin + 1
    Next p

    Set Silabear_FR_ConRevision = resultado

End Function

Public Function Silabear_FR(ByVal Texto As String) As Collection

    Dim col As New Collection
    Dim i As Long, ini As Long
    Dim c1 As String, c2 As String, c3 As String
    Dim par As String

    Texto = Trim$(Texto)
    If Len(Texto) = 0 Then
        Set Silabear_FR = col
        Exit Function
    End If

    ini = 1

    For i = 2 To Len(Texto)

        c1 = Mid$(Texto, i - 1, 1)
        c2 = Mid$(Texto, i, 1)
        par = c1 & c2

        ' ---------------------------------------------------------
        ' 0. Espacios ? separan palabras
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
        ' 1. VV ? hiato (francés no forma diptongos ortográficos)
        ' ---------------------------------------------------------
        If EsVocal_FR(c1) And EsVocal_FR(c2) Then
            col.Add Array(ini, i - 1)
            ini = i
            GoTo Siguiente
        End If

        ' ---------------------------------------------------------
        ' 2. VCV ? V | CV
        ' ---------------------------------------------------------
        If EsVocal_FR(c1) And EsConsonant_FR(c2) Then
            If i < Len(Texto) Then
                c3 = Mid$(Texto, i + 1, 1)
                If EsVocal_FR(c3) Then
                    col.Add Array(ini, i - 1)
                    ini = i
                    GoTo Siguiente
                End If
            End If
        End If

        ' ---------------------------------------------------------
        ' 3. CCV ? C | CV
        ' ---------------------------------------------------------
        If EsConsonant_FR(c1) And EsConsonant_FR(c2) Then
            If i < Len(Texto) Then
                c3 = Mid$(Texto, i + 1, 1)
                If EsVocal_FR(c3) Then
                    If Not EsGrupInseparable_FR(par) Then
                        col.Add Array(ini, i - 1)
                        ini = i
                        GoTo Siguiente
                    End If
                End If
            End If
        End If

Siguiente:
    Next i

    ' Última sílaba
    If ini <= Len(Texto) Then
        col.Add Array(ini, Len(Texto))
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

    ' GÜE / GÜI ? /gw/ ? id 57
    If g = "GÜE" Or g = "GÜI" Then
        ReglasFrances = 57
        Exit Function
    End If

    ' GUE / GUI ? /g/ ? id 31
    If g = "GUE" Or g = "GUI" Then
        ReglasFrances = 31
        Exit Function
    End If

    ' QUE / QUI ? /k/ ? id 30
    If g = "QUE" Or g = "QUI" Then
        ReglasFrances = 30
        Exit Function
    End If


    ' ============================================================
    '   DÍGRAFOS Y CASOS ESPECIALES
    ' ============================================================

    ' CH ? /?/ ? id 36
    If g = "CH" Then
        ReglasFrances = 36
        Exit Function
    End If

    ' GN ? /?/ ? id 41
    If g = "GN" Then
        ReglasFrances = 41
        Exit Function
    End If

    ' J ? /?/ ? id 37
    If g = "J" Then
        ReglasFrances = 37
        Exit Function
    End If

    ' G + E/I/Y ? /?/ ? id 37
    If g = "G" And (sig = "E" Or sig = "I" Or sig = "Y") Then
        ReglasFrances = 37
        Exit Function
    End If

    ' S entre vocales ? /z/ ? id 35
    If g = "S" And (ant Like "[AEIOU]" And sig Like "[AEIOU]") Then
        ReglasFrances = 35
        Exit Function
    End If

    ' Ç ? /s/ ? id 34
    If g = "Ç" Then
        ReglasFrances = 34
        Exit Function
    End If


    ' ============================================================
    '   NASALIZACIONES
    ' ============================================================

    ' AN / AM / EN / EM ? /?~/ ? id 2
    If g = "AN" Or g = "AM" Or g = "EN" Or g = "EM" Then
        ReglasFrances = 2
        Exit Function
    End If

    ' IN / IM / AIN / EIN / EIM / YN / YM ? /?~/ ? id 3
    If g = "IN" Or g = "IM" Or g = "AIN" Or g = "EIN" Or g = "EIM" Or g = "YN" Or g = "YM" Then
        ReglasFrances = 3
        Exit Function
    End If

    ' ON / OM ? /?~/ ? id 4
    If g = "ON" Or g = "OM" Then
        ReglasFrances = 4
        Exit Function
    End If

    ' UN / UM ? /œ~/ ? id 3 (aproximación razonable)
    If g = "UN" Or g = "UM" Then
        ReglasFrances = 3
        Exit Function
    End If


    ' ============================================================
    '   DÍGRAFOS VOCÁLICOS (diptongos franceses)
    ' ============================================================

    ' OI ? /wa/ ? id 18
    If g = "OI" Then ReglasFrances = 18: Exit Function

    ' AI ? /?/ ? id 6
    If g = "AI" Then ReglasFrances = 6: Exit Function

    ' EI ? /e/ ? id 5
    If g = "EI" Then ReglasFrances = 5: Exit Function

    ' OU ? /u/ ? id 10
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

    ' C + E/I/Y ? /s/
    If g = "C" And (sig = "E" Or sig = "I" Or sig = "Y") Then
        ReglasFrances = 34
        Exit Function
    End If

    ' S ? /s/
    If g = "S" Then ReglasFrances = 34: Exit Function

    ' X ? /ks/ o /gz/ ? simplificamos a /s/ (el motor segmenta la K aparte)
    If g = "X" Then
        ReglasFrances = 34
        Exit Function
    End If

    If g = "M" Then ReglasFrances = 39: Exit Function
    If g = "N" Then ReglasFrances = 40: Exit Function

    ' L ? /l/
    If g = "L" Then ReglasFrances = 43: Exit Function

    ' R ? /?/ ? id 47
    If g = "R" Then
        ReglasFrances = 47
        Exit Function
    End If

    ' H ? muda ? id 38
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

Public Function MF_NormalizarVocales_FR(ByVal Texto As String) As String

    ' A
    Texto = Replace(Texto, "À", "A")   ' abierta
    Texto = Replace(Texto, "Á", "A")   ' rara, pero robustez
    Texto = Replace(Texto, "Â", "Â")   ' cerrada
    Texto = Replace(Texto, "Ä", "A¨")  ' hiato

    ' E
    Texto = Replace(Texto, "È", "E")   ' abierta
    Texto = Replace(Texto, "É", "E´")  ' cerrada
    Texto = Replace(Texto, "Ê", "Ê")   ' cerrada tensa
    Texto = Replace(Texto, "Ë", "E¨")  ' hiato

    ' I
    Texto = Replace(Texto, "Ì", "I")   ' robustez
    Texto = Replace(Texto, "Í", "I")   ' robustez
    Texto = Replace(Texto, "Î", "Î")   ' cerrada
    Texto = Replace(Texto, "Ï", "I¨")  ' hiato

    ' O
    Texto = Replace(Texto, "Ò", "O")   ' robustez
    Texto = Replace(Texto, "Ó", "O")   ' robustez
    Texto = Replace(Texto, "Ô", "Ô")   ' cerrada
    Texto = Replace(Texto, "Ö", "O¨")  ' hiato

    ' U
    Texto = Replace(Texto, "Ù", "U")   ' abierta
    Texto = Replace(Texto, "Ú", "U")   ' robustez
    Texto = Replace(Texto, "Û", "Û")   ' cerrada
    Texto = Replace(Texto, "Ü", "U¨")  ' hiato

    MF_NormalizarVocales_FR = Texto

End Function

'Public Sub MF_SilabearAjustesFR( _
'        ByVal Texto As String, _
'        ByRef Silabas As Collection _
'    )
'
'    Dim i As Long
'
'    ' ============================================================
'    ' 1. ELIMINAR LA E MUDA FINAL
'    ' ============================================================
'    If Right$(Texto, 1) = "E" Then
'        Call MF_EliminarVocalFinal(Silabas, Len(Texto))
'    End If
'
'    ' ============================================================
'    ' 2. AGRUPAR NASALES (an, am, en, em, in, im, ain, ein, un, um, on, om)
'    ' ============================================================
'    Dim nasales As Variant
'    nasales = Array("AN", "AM", "EN", "EM", "IN", "IM", "AIN", "EIN", "UN", "UM", "ON", "OM")
'
'    For i = 2 To Len(Texto)
'        Dim seq2 As String, seq3 As String
'        seq2 = Mid$(Texto, i - 1, 2)
'        seq3 = Mid$(Texto, i - 2, 3)
'
'        If EsMiembro(seq3, nasales) Then
'            Call MF_UnirVocalesEnDiptongo(Silabas, i - 1)
'        ElseIf EsMiembro(seq2, nasales) Then
'            Call MF_UnirVocalesEnDiptongo(Silabas, i)
'        End If
'    Next i
'
'    ' ============================================================
'    ' 3. ATAQUES CONSONÁNTICOS PERMITIDOS EN FRANCÉS
'    ' ============================================================
'    Dim ataques As Variant
'    ataques = Array("TR", "DR", "PR", "BR", "CR", "GR", "FR", "FL", "CL", "GL", "PL")
'
'    For i = 2 To Len(Texto) - 1
'        Dim par As String
'        par = Mid$(Texto, i, 2)
'
'        If EsMiembro(par, ataques) Then
'            Call MF_UnirConsonantesEnAtaque(Silabas, i)
'        End If
'    Next i
'
'    ' ============================================================
'    ' 4. HIATOS POR DIÉRESIS (ï, ë, ü)
'    ' ============================================================
'    Dim dieresis As String
'    dieresis = "ÏËÜ"
'
'    For i = 2 To Len(Texto)
'        If InStr(dieresis, Mid$(Texto, i, 1)) > 0 Then
'            Call MF_ForzarDivisionSilabica(Silabas, i)
'        End If
'    Next i
'
'End Sub

'Public Sub MF_MarcarTonicaFrances( _
'        ByVal Texto As String, _
'        ByRef esTonica() As Boolean _
'    )
'
'    Dim Silabas As Collection
'    Dim idxTonica As Long
'    Dim inicio As Long, fin As Long
'    Dim i As Long
'
'    ' --------------------------------------------------------
'    ' 1. Silabear palabra
'    ' --------------------------------------------------------
'    'Set silabas = MF_SilabearCastellano(Texto)
'    Set Silabas = MF_Silabear(Texto, "fr")
'
'
'    If Silabas Is Nothing Then Exit Sub
'    If Silabas.Count = 0 Then Exit Sub
'
'    ' --------------------------------------------------------
'    ' 2. Acento francés: SIEMPRE en la última sílaba fonética
'    ' --------------------------------------------------------
'    idxTonica = Silabas.Count
'
'    ' --------------------------------------------------------
'    ' 3. Marcar índices tónicos
'    ' --------------------------------------------------------
'    inicio = Silabas(idxTonica)(1)
'    fin = Silabas(idxTonica)(2)
'
'    For i = inicio To fin
'        esTonica(i) = True
'    Next i
'
'End Sub


