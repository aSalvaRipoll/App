Attribute VB_Name = "modMotor_Idioma_ES"

Option Compare Database
Option Explicit

'=================
'==  Castellano ==
'=================

Public Sub MF_SilabearAjustesES( _
        ByVal Texto As String, _
        ByRef silabas As Collection _
    )

    Dim i As Long
    Dim vocalesTilde As String
    vocalesTilde = "ÁÉÍÓÚ"

    ' ============================================================
    ' 1. HIATOS CON TILDE (rompen diptongo)
    ' ============================================================
    For i = 2 To Len(Texto)
        If InStr(vocalesTilde, Mid$(Texto, i, 1)) > 0 Then

            ' Si la vocal anterior es vocal ? romper sílaba
            If InStr("AEIOUÁÉÍÓÚ", Mid$(Texto, i - 1, 1)) > 0 Then
                Call MF_ForzarDivisionSilabica(silabas, i)
            End If

        End If
    Next i

    ' ============================================================
    ' 2. GRUPOS CONSONÁNTICOS INSEPARABLES
    ' ============================================================
    Dim grupos As Variant
    grupos = Array("BR", "BL", "CR", "CL", "DR", "FR", "FL", _
                   "GR", "GL", "PR", "PL", "TR", "TL")

    For i = 2 To Len(Texto) - 1
        Dim par As String
        par = Mid$(Texto, i, 2)

        If EsMiembro(par, grupos) Then
            ' Si el universal ha dividido entre estas dos consonantes ? corregir
            Call MF_UnirConsonantesEnAtaque(silabas, i)
        End If
    Next i

    ' ============================================================
    ' 3. LL y RR nunca se separan
    ' ============================================================
    For i = 2 To Len(Texto) - 1
        If Mid$(Texto, i, 2) = "LL" Or Mid$(Texto, i, 2) = "RR" Then
            Call MF_UnirConsonantesEnAtaque(silabas, i)
        End If
    Next i

End Sub


Public Sub MF_MarcarTonicaCastellano( _
        ByVal Texto As String, _
        ByRef esTonica() As Boolean _
    )

    Dim i As Long
    Dim silabas As Collection
    Dim idxTonica As Long
    Dim inicio As Long, fin As Long

    ' --------------------------------------------------------
    ' 1. Silabear palabra
    ' --------------------------------------------------------
    'Set silabas = MF_SilabearCastellano(Texto)
    Set silabas = MF_Silabear(Texto, "es")


    If silabas.Count = 0 Then Exit Sub

    ' --------------------------------------------------------
    ' 2. Determinar sílaba tónica
    ' --------------------------------------------------------
    idxTonica = MF_DetectarTonicaCastellano(Texto, silabas)

    If idxTonica = 0 Then Exit Sub

    ' --------------------------------------------------------
    ' 3. Marcar índices de la sílaba tónica
    ' --------------------------------------------------------
    inicio = silabas(idxTonica)(1)
    fin = silabas(idxTonica)(2)

    For i = inicio To fin
        esTonica(i) = True
    Next i

End Sub

'Public Function MF_SilabearCastellano(ByVal Texto As String) As Collection
'
'    Dim col As New Collection
'    Dim i As Long, ini As Long
'    Dim vocales As String
'    vocales = "AEIOUÁÉÍÓÚÜ"
'
'    ini = 1
'
'    For i = 2 To Len(Texto)
'        If InStr(vocales, Mid$(Texto, i, 1)) > 0 And _
'           InStr(vocales, Mid$(Texto, i - 1, 1)) = 0 Then
'
'            col.Add Array(ini, i - 1)
'            ini = i
'        End If
'    Next i
'
'    col.Add Array(ini, Len(Texto))
'
'    Set MF_SilabearCastellano = col
'
'End Function

Public Function MF_DetectarTonicaCastellano( _
        ByVal Texto As String, _
        ByVal silabas As Collection _
    ) As Long

    Dim i As Long
    Dim vocalesTilde As String
    vocalesTilde = "ÁÉÍÓÚ"

    ' 1. Si hay tilde ? esa sílaba es tónica
    For i = 1 To Len(Texto)
        If InStr(vocalesTilde, Mid$(Texto, i, 1)) > 0 Then
            MF_DetectarTonicaCastellano = MF_SilabaDeIndice(i, silabas)
            Exit Function
        End If
    Next i

    ' 2. Si no hay tilde ? reglas generales
    Dim ultima As String
    ultima = Right$(Texto, 1)

    If ultima = "N" Or ultima = "S" Or _
       InStr("AEIOU", ultima) > 0 Then
        MF_DetectarTonicaCastellano = silabas.Count - 1
    Else
        MF_DetectarTonicaCastellano = silabas.Count
    End If

End Function

Public Function MF_SilabaDeIndice( _
        ByVal idx As Long, _
        ByVal silabas As Collection _
    ) As Long

    Dim i As Long
    For i = 1 To silabas.Count
        Debug.Print "i=" & i, "Tipo=" & TypeName(silabas(i)), "Len=" & IIf(IsArray(silabas(i)), UBound(silabas(i)), "NO ARRAY")
        If idx >= silabas(i)(1) And idx <= silabas(i)(2) Then
            MF_SilabaDeIndice = i
            Exit Function
        End If
    Next i

End Function


' ============================================================
'   ReglasCastellano (ESP)
'   Devuelve idFonema según la fonética del castellano.
'   Si no aplica, devuelve 0 para que el motor siga probando.
' ============================================================

Public Function ReglasCastellano( _
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
        ReglasCastellano = 57
        Exit Function
    End If

    ' GUE / GUI ? /g/ (U muda) ? id 31
    If g = "GUE" Or g = "GUI" Then
        ReglasCastellano = 31
        Exit Function
    End If

    ' QUE / QUI ? /k/ ? id 30
    If g = "QUE" Or g = "QUI" Then
        ReglasCastellano = 30
        Exit Function
    End If


    ' ============================================================
    '   DÍGRAFOS Y CASOS ESPECIALES
    ' ============================================================

    ' CH ? /t?/ ? id 50
    If g = "CH" Then
        ReglasCastellano = 50
        Exit Function
    End If

    ' LL ? /?/ (fonema histórico; hoy yeísmo ? /?/)
    ' Usamos /?/ ? id 44 para mantener coherencia fonética
    If g = "LL" Then
        ReglasCastellano = 44
        Exit Function
    End If

    ' RR ? /r/ múltiple ? id 46
    If g = "RR" Then
        ReglasCastellano = 46
        Exit Function
    End If

    ' Ñ ? /?/ ? id 41
    If g = "Ñ" Then
        ReglasCastellano = 41
        Exit Function
    End If

    ' GU + vocal ? /g/ ? id 31
    If g = "GU" And (sig = "A" Or sig = "O" Or sig = "U") Then
        ReglasCastellano = 31
        Exit Function
    End If

    ' QU + vocal ? /k/ ? id 30
    If g = "QU" And (sig = "A" Or sig = "O" Or sig = "U") Then
        ReglasCastellano = 30
        Exit Function
    End If


    ' ============================================================
    '   DÍGRAFOS VOCÁLICOS (diptongos castellanos)
    ' ============================================================

    If g = "AI" Then ReglasCastellano = 12: Exit Function
    If g = "EI" Then ReglasCastellano = 13: Exit Function
    If g = "OI" Then ReglasCastellano = 14: Exit Function
    If g = "OU" Then ReglasCastellano = 15: Exit Function
    If g = "AU" Then ReglasCastellano = 16: Exit Function


    ' ============================================================
    '   MONÓGRAFOS — VOCALES (5 vocales)
    ' ============================================================

    If g = "A" Then ReglasCastellano = 1: Exit Function
    If g = "E" Then ReglasCastellano = 5: Exit Function
    If g = "I" Then ReglasCastellano = 9: Exit Function
    If g = "O" Then ReglasCastellano = 7: Exit Function
    If g = "U" Then ReglasCastellano = 10: Exit Function


    ' ============================================================
    '   MONÓGRAFOS — CONSONANTES
    ' ============================================================

    If g = "P" Then ReglasCastellano = 26: Exit Function
    If g = "B" Then ReglasCastellano = 27: Exit Function
    If g = "T" Then ReglasCastellano = 28: Exit Function
    If g = "D" Then ReglasCastellano = 29: Exit Function
    If g = "K" Then ReglasCastellano = 30: Exit Function
    If g = "G" Then ReglasCastellano = 31: Exit Function

    If g = "F" Then ReglasCastellano = 32: Exit Function

    ' C/Z ? /?/ (castellano estándar)
    If g = "C" And (sig = "E" Or sig = "I") Then
        ReglasCastellano = 54   ' /?/
        Exit Function
    End If
    If g = "Z" Then
        ReglasCastellano = 54   ' /?/
        Exit Function
    End If

    ' S ? /s/
    If g = "S" Then ReglasCastellano = 34: Exit Function

    ' J / G + E/I ? /x/ ? id 58
    If g = "J" Then ReglasCastellano = 58: Exit Function
    If g = "G" And (sig = "E" Or sig = "I") Then
        ReglasCastellano = 58
        Exit Function
    End If

    ' M / N
    If g = "M" Then ReglasCastellano = 39: Exit Function
    If g = "N" Then ReglasCastellano = 40: Exit Function

    ' L / R simple
    If g = "L" Then ReglasCastellano = 43: Exit Function
    If g = "R" Then ReglasCastellano = 45: Exit Function

    ' H muda ? /h/ glotal suave ? id 38
    If g = "H" Then ReglasCastellano = 38: Exit Function


    ' ============================================================
    '   SI NO APLICA, DEVOLVER 0
    ' ============================================================
    ReglasCastellano = 0

End Function
'============================================================================================

Public Function MF_NormalizarVocales_ES(ByVal Texto As String) As String

    ' A
    Texto = Replace(Texto, "Á", "A")
    Texto = Replace(Texto, "À", "A")
    Texto = Replace(Texto, "Ä", "A")
    Texto = Replace(Texto, "Â", "A")

    ' E
    Texto = Replace(Texto, "É", "E")
    Texto = Replace(Texto, "È", "E")
    Texto = Replace(Texto, "Ë", "E")
    Texto = Replace(Texto, "Ê", "E")

    ' I
    Texto = Replace(Texto, "Í", "I")
    Texto = Replace(Texto, "Ì", "I")
    Texto = Replace(Texto, "Ï", "I")
    Texto = Replace(Texto, "Î", "I")

    ' O
    Texto = Replace(Texto, "Ó", "O")
    Texto = Replace(Texto, "Ò", "O")
    Texto = Replace(Texto, "Ö", "O")
    Texto = Replace(Texto, "Ô", "O")

    ' U (sin tocar Ü)
    Texto = Replace(Texto, "Ú", "U")
    Texto = Replace(Texto, "Ù", "U")
    Texto = Replace(Texto, "Û", "U")

    MF_NormalizarVocales_ES = Texto

End Function
