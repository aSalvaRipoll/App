Attribute VB_Name = "modMotor_Idioma_CA"

Option Compare Database
Option Explicit

'============================================================================================

'=============
'== Catalán ==
'=============
Public Sub MF_SilabearAjustesCatalan( _
        ByVal Texto As String, _
        ByRef silabas As Collection _
    )

    Dim i As Long

    ' ============================================================
    ' 1. ATAQUES CONSONÁNTICOS PROPIOS DEL CATALÁN CENTRAL
    ' ============================================================
    ' Nota: estos grupos deben permanecer juntos si van seguidos de vocal.
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
    ' 2. LL y RR nunca se separan
    ' ============================================================
    For i = 2 To Len(Texto) - 1
        If Mid$(Texto, i, 2) = "LL" Or Mid$(Texto, i, 2) = "RR" Then
            Call MF_UnirConsonantesEnAtaque(silabas, i)
        End If
    Next i

    ' ============================================================
    ' 3. Diptongos catalanes (refuerzo)
    ' ============================================================
    Dim dipt As Variant
    dipt = Array( _
        "IA", "IE", "IO", "IU", _
        "UA", "UE", "UI", "UO", _
        "AI", "EI", "OI", "AU", "EU", "OU" _
    )

    For i = 2 To Len(Texto)
        Dim seq As String
        seq = Mid$(Texto, i - 1, 2)

        If EsMiembro(seq, dipt) Then
            Call MF_UnirVocalesEnDiptongo(silabas, i)
        End If
    Next i

End Sub

Public Sub MF_MarcarTonicaCatalan( _
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
    'Set silabas = MF_SilabearUniversalBase(Texto)
    Set silabas = MF_Silabear(Texto, "ca")
    
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
    ' 3. Si no hay tilde ? reglas catalanas
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

        ' 3.3. Reglas generales catalanas
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

    ' GÜE / GÜI ? /gw/ ? id 57
    If g = "GÜE" Or g = "GÜI" Then
        ReglasCatala = 57
        Exit Function
    End If

    ' GUE / GUI ? /g/ (U muda) ? id 31
    If g = "GUE" Or g = "GUI" Then
        ReglasCatala = 31
        Exit Function
    End If

    ' QUE / QUI ? /k/ ? id 30
    If g = "QUE" Or g = "QUI" Then
        ReglasCatala = 30
        Exit Function
    End If


    ' ============================================================
    '   DÍGRAFOS Y CASOS ESPECIALES
    ' ============================================================

    ' TX ? /t?/ ? id 50 (en catalán central)
    If g = "TX" Then
        ReglasCatala = 50
        Exit Function
    End If

    ' CH ? /t?/ ? id 50 (préstecs)
    If g = "CH" Then
        ReglasCatala = 50
        Exit Function
    End If

    ' NY ? /?/ ? id 41
    If g = "NY" Then
        ReglasCatala = 41
        Exit Function
    End If

    ' LL ? /?/ ? id 44
    If g = "LL" Then
        ReglasCatala = 44
        Exit Function
    End If

    ' L·L ? /l?/ ? id 61 (ela geminada)
    If g = "L·L" Or g = "L.L" Then
        ReglasCatala = 61
        Exit Function
    End If

    ' IX ? /?/ ? id 36
    If g = "IX" Then
        ReglasCatala = 36
        Exit Function
    End If

    ' TJ / TG ? /d?/ ? id 51
    If g = "TJ" Or g = "TG" Then
        ReglasCatala = 51
        Exit Function
    End If

    ' IG final ? /t?/ ? id 50
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

    ' E tònica ? /?/ (id 6), àtona ? /e/ (id 5)
    If g = "E" Then
        If esTonica Then
            ReglasCatala = 6   ' /?/
        Else
            ReglasCatala = 5   ' /e/
        End If
        Exit Function
    End If

    ' O tònica ? /?/ (id 8), àtona ? /o/ (id 7)
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

