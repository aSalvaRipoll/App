Attribute VB_Name = "modMotor_Idioma_EU"

Option Compare Database
Option Explicit

'=================
'==   Euskara   ==
'=================
Public Sub MF_SilabearAjustesEU( _
        ByVal Texto As String, _
        ByRef silabas As Collection _
    )

    Dim i As Long

    ' ============================================================
    ' 1. ATAQUES CONSONÁNTICOS PERMITIDOS EN EUSKARA
    ' ============================================================
    Dim ataques As Variant
    ataques = Array( _
        "BR", "BL", "TR", "DR", "KR", "KL", _
        "GR", "GL", "PR", "PL" _
    )

    For i = 2 To Len(Texto) - 1
        Dim par As String
        par = Mid$(Texto, i, 2)

        If EsMiembro(par, ataques) Then
            Call MF_UnirConsonantesEnAtaque(silabas, i)
        End If
    Next i

    ' ============================================================
    ' 2. RR nunca se separa
    ' ============================================================
    For i = 2 To Len(Texto) - 1
        If Mid$(Texto, i, 2) = "RR" Then
            Call MF_UnirConsonantesEnAtaque(silabas, i)
        End If
    Next i

    ' ============================================================
    ' 3. DIPTONGOS EUSKÉRICOS (muy estables)
    ' ============================================================
    Dim dipt As Variant
    dipt = Array( _
        "AI", "EI", "OI", "UI", _
        "AU", "EU", "OU", _
        "IA", "IE", "IO", "IU", _
        "UA", "UE", "UI", "UO" _
    )

    For i = 2 To Len(Texto)
        Dim dv As String
        dv = Mid$(Texto, i - 1, 2)

        If EsMiembro(dv, dipt) Then
            Call MF_UnirVocalesEnDiptongo(silabas, i)
        End If
    Next i

    ' ============================================================
    ' 4. EVITAR ATAQUES NO PERMITIDOS (SC, SP, ST, SM, SN, TL, DL…)
    '    ? si el universal los ha unido, los separamos
    ' ============================================================
    Dim noAtaques As Variant
    noAtaques = Array("SC", "SP", "ST", "SM", "SN", "TL", "DL", "TS", "TX", "TZ")

    For i = 2 To Len(Texto) - 1
        Dim seq As String
        seq = Mid$(Texto, i, 2)

        If EsMiembro(seq, noAtaques) Then
            Call MF_ForzarDivisionSilabica(silabas, i)
        End If
    Next i

End Sub

Public Sub MF_MarcarTonicaEuskara( _
        ByVal Texto As String, _
        ByRef esTonica() As Boolean _
    )

    Dim silabas As Collection
    Dim idxTonica As Long
    Dim inicio As Long, fin As Long
    Dim i As Long

    ' Puedes reutilizar el silabeador del castellano
    ' o crear uno específico si más adelante lo necesitas.
    'Set silabas = MF_SilabearCastellano(Texto)
    Set silabas = MF_Silabear(Texto, "eu")


    If silabas Is Nothing Then Exit Sub
    If silabas.Count = 0 Then Exit Sub

    ' Acento fijo en la penúltima sílaba
    If silabas.Count = 1 Then
        idxTonica = 1
    Else
        idxTonica = silabas.Count - 1
    End If

    inicio = silabas(idxTonica)(1)
    fin = silabas(idxTonica)(2)

    For i = inicio To fin
        esTonica(i) = True
    Next i

End Sub


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

    ' GÜE / GÜI ? /gw/ ? id 57 (préstamos)
    If g = "GÜE" Or g = "GÜI" Then
        ReglasEuskera = 57
        Exit Function
    End If

    ' GUE / GUI ? /g/ (U muda) ? id 31
    If g = "GUE" Or g = "GUI" Then
        ReglasEuskera = 31
        Exit Function
    End If

    ' QUE / QUI ? /k/ ? id 30
    If g = "QUE" Or g = "QUI" Then
        ReglasEuskera = 30
        Exit Function
    End If


    ' ============================================================
    '   DÍGRAFOS Y CASOS ESPECIALES
    ' ============================================================

    ' TX ? /t?/ ? id 50
    If g = "TX" Then
        ReglasEuskera = 50
        Exit Function
    End If

    ' TS / TZ ? /ts/ ? id 52
    If g = "TS" Or g = "TZ" Then
        ReglasEuskera = 52
        Exit Function
    End If

    ' LL ? /?/ ? id 44
    If g = "LL" Then
        ReglasEuskera = 44
        Exit Function
    End If

    ' RR ? /r/ múltiple ? id 46
    If g = "RR" Then
        ReglasEuskera = 46
        Exit Function
    End If

    ' Ñ ? /?/ ? id 41
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

    ' S / Z ? /s/ (no existe /?/)
    If g = "S" Then ReglasEuskera = 34: Exit Function
    If g = "Z" Then ReglasEuskera = 34: Exit Function

    ' X ? /?/ ? id 36
    If g = "X" Then ReglasEuskera = 36: Exit Function

    ' J ? /j/ ? id 48
    If g = "J" Then ReglasEuskera = 48: Exit Function

    If g = "M" Then ReglasEuskera = 39: Exit Function
    If g = "N" Then ReglasEuskera = 40: Exit Function

    If g = "L" Then ReglasEuskera = 43: Exit Function
    If g = "R" Then ReglasEuskera = 45: Exit Function

    ' H ? aspiración suave ? id 38
    If g = "H" Then ReglasEuskera = 38: Exit Function


    ' ============================================================
    '   SI NO APLICA, DEVOLVER 0
    ' ============================================================
    ReglasEuskera = 0

End Function

'============================================================================================

Public Function MF_NormalizarVocales_EU(ByVal Texto As String) As String

    ' A
    Texto = Replace(Texto, "Á", "A")
    Texto = Replace(Texto, "À", "A")

    ' E
    Texto = Replace(Texto, "É", "E")
    Texto = Replace(Texto, "È", "E")

    ' I
    Texto = Replace(Texto, "Í", "I")
    Texto = Replace(Texto, "Ì", "I")

    ' O
    Texto = Replace(Texto, "Ó", "O")
    Texto = Replace(Texto, "Ò", "O")

    ' U
    Texto = Replace(Texto, "Ú", "U")
    Texto = Replace(Texto, "Ù", "U")

    MF_NormalizarVocales_EU = Texto

End Function
