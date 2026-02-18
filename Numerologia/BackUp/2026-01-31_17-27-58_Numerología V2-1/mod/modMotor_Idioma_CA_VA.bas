Attribute VB_Name = "modMotor_Idioma_CA_VA"
Option Compare Database
Option Explicit

'============================================================================================

'================
'== Valenciano ==
'================
Public Sub MF_SilabearAjustesCA_VA( _
        ByVal Texto As String, _
        ByRef silabas As Collection _
    )

    Dim i As Long

    ' ============================================================
    ' 1. ATAQUES CONSONÁNTICOS PROPIOS DEL VALENCIANO
    ' ============================================================
    Dim ataques As Variant
    ataques = Array( _
        "BR", "BL", "CR", "CL", "DR", "FR", "FL", _
        "GR", "GL", "PR", "PL", "TR", "TL", "DL", _
        "SC", "SP", "ST", "SM", "SN" _
    )

    For i = 2 To Len(Texto) - 1
        Dim par As String
        par = Mid$(Texto, i, 2)

        If EsMiembro(par, ataques) Then
            Call MF_UnirConsonantesEnAtaque(silabas, i)
        End If
    Next i

    ' ============================================================
    ' 2. TS y TZ ? no dividir si van seguidas de vocal
    ' ============================================================
    For i = 2 To Len(Texto) - 1
        Dim seq As String
        seq = Mid$(Texto, i, 2)

        If seq = "TS" Or seq = "TZ" Then
            If esVocal(Mid$(Texto, i + 2, 1)) Then
                Call MF_UnirConsonantesEnAtaque(silabas, i)
            End If
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

    ' ============================================================
    ' 4. HIATOS VALENCIANOS (ea, eo, oa)
    ' ============================================================
    Dim hiatos As Variant
    hiatos = Array("EA", "EO", "OA")

    For i = 2 To Len(Texto)
        Dim hv As String
        hv = Mid$(Texto, i - 1, 2)

        If EsMiembro(hv, hiatos) Then
            Call MF_ForzarDivisionSilabica(silabas, i)
        End If
    Next i

    ' ============================================================
    ' 5. Diptongos valencianos (refuerzo)
    ' ============================================================
    Dim dipt As Variant
    dipt = Array( _
        "AI", "EI", "OI", _
        "AU", "EU", "OU", _
        "IA", "IE", "IO", "IU", _
        "UA", "UE", "UI" _
    )

    For i = 2 To Len(Texto)
        Dim dv As String
        dv = Mid$(Texto, i - 1, 2)

        If EsMiembro(dv, dipt) Then
            Call MF_UnirVocalesEnDiptongo(silabas, i)
        End If
    Next i

End Sub

Public Sub MF_MarcarTonicaValenciano( _
        ByVal Texto As String, _
        ByRef esTonica() As Boolean _
    )

    Dim silabas As Collection
    Dim idxTonica As Long
    Dim inicio As Long, fin As Long
    Dim i As Long
    Dim vocalesTilde As String
    Dim ultima As String

    vocalesTilde = "ÁÉÍÓÚÀÈÌÒÙ"

    ' --------------------------------------------------------
    ' 1. Silabear palabra
    ' --------------------------------------------------------
    'Set silabas = MF_SilabearCastellano(Texto)
    Set silabas = MF_Silabear(Texto, "ca-va")


    If silabas Is Nothing Then Exit Sub
    If silabas.Count = 0 Then Exit Sub

    ' --------------------------------------------------------
    ' 2. Buscar tilde
    ' --------------------------------------------------------
    For i = 1 To Len(Texto)
        If InStr(vocalesTilde, Mid$(Texto, i, 1)) > 0 Then
            idxTonica = MF_SilabaDeIndice(i, silabas)
            Exit For
        End If
    Next i

    ' --------------------------------------------------------
    ' 3. Si no hay tilde ? reglas valencianas
    ' --------------------------------------------------------
    If idxTonica = 0 Then

        ultima = Right$(Texto, 1)

        ' 3.1. Infinitivos (terminados en -AR, -ER, -IR)
        If Len(Texto) >= 2 Then
            Dim ult2 As String
            ult2 = Right$(Texto, 2)

            If ult2 = "AR" Or ult2 = "ER" Or ult2 = "IR" Then
                idxTonica = silabas.Count
                GoTo Marcar
            End If
        End If

        ' 3.2. Palabras acabadas en -IG ? oxítonas
        If Len(Texto) >= 2 Then
            If Right$(Texto, 2) = "IG" Then
                idxTonica = silabas.Count
                GoTo Marcar
            End If
        End If

        ' 3.3. Reglas generales valencianas
        If InStr("AEIOU", ultima) > 0 Or _
           Right$(Texto, 2) = "AS" Or _
           Right$(Texto, 2) = "ES" Or _
           Right$(Texto, 2) = "IS" Or _
           Right$(Texto, 2) = "OS" Or _
           Right$(Texto, 2) = "US" Then

            ' Penúltima
            If silabas.Count = 1 Then
                idxTonica = 1
            Else
                idxTonica = silabas.Count - 1
            End If

        Else
            ' Última
            idxTonica = silabas.Count
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

    ' GÜE / GÜI ? /gw/ ? id 57
    If g = "GÜE" Or g = "GÜI" Then
        ReglasValenciano = 57
        Exit Function
    End If

    ' GUE / GUI ? /g/ (U muda) ? id 31
    If g = "GUE" Or g = "GUI" Then
        ReglasValenciano = 31
        Exit Function
    End If

    ' QUE / QUI ? /k/ ? id 30
    If g = "QUE" Or g = "QUI" Then
        ReglasValenciano = 30
        Exit Function
    End If


    ' ============================================================
    '   DÍGRAFOS Y CASOS ESPECIALES
    ' ============================================================

    ' TX ? /t?/ ? id 50 (en valenciano no existe /t?/)
    If g = "TX" Then
        ReglasValenciano = 50
        Exit Function
    End If

    ' CH ? /t?/ ? id 50
    If g = "CH" Then
        ReglasValenciano = 50
        Exit Function
    End If

    ' NY ? /?/ ? id 41
    If g = "NY" Then
        ReglasValenciano = 41
        Exit Function
    End If

    ' LL ? /?/ ? id 44
    If g = "LL" Then
        ReglasValenciano = 44
        Exit Function
    End If

    ' L·L ? /l?/ ? id 61
    If g = "L·L" Or g = "L.L" Then
        ReglasValenciano = 61
        Exit Function
    End If

    ' IX ? /?/ ? id 36
    If g = "IX" Then
        ReglasValenciano = 36
        Exit Function
    End If

    ' TJ / TG ? /d?/ ? id 51
    If g = "TJ" Or g = "TG" Then
        ReglasValenciano = 51
        Exit Function
    End If

    ' IG final ? /t?/ ? id 50
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

Public Function MF_NormalizarVocales_CA_VA(ByVal Texto As String) As String
    MF_NormalizarVocales_CA_VA = MF_NormalizarVocales_CA(Texto)
End Function

