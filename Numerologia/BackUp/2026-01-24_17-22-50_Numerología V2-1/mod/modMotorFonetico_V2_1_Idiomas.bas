Attribute VB_Name = "modMotorFonetico_V2_1_Idiomas"

Option Compare Database
Option Explicit

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

    ' GÜE / GÜI ? /gw/ ? id 57
    If g = "GÜE" Or g = "GÜI" Then
        ReglasMallorquin = 57
        Exit Function
    End If

    ' GUE / GUI ? /g/ (U muda) ? id 31
    If g = "GUE" Or g = "GUI" Then
        ReglasMallorquin = 31
        Exit Function
    End If

    ' QUE / QUI ? /k/ ? id 30
    If g = "QUE" Or g = "QUI" Then
        ReglasMallorquin = 30
        Exit Function
    End If


    ' ============================================================
    '   DÍGRAFOS Y CASOS ESPECIALES
    ' ============================================================

    ' TX ? /t?/ ? id 60 (mallorquín)
    If g = "TX" Then
        ReglasMallorquin = 60
        Exit Function
    End If

    ' CH ? /t?/ ? id 50 (préstamos)
    If g = "CH" Then
        ReglasMallorquin = 50
        Exit Function
    End If

    ' NY ? /?/ ? id 41
    If g = "NY" Then
        ReglasMallorquin = 41
        Exit Function
    End If

    ' LL ? /?/ ? id 44
    If g = "LL" Then
        ReglasMallorquin = 44
        Exit Function
    End If

    ' L·L ? /l?/ ? id 61 (ela geminada)
    If g = "L·L" Or g = "L.L" Then
        ReglasMallorquin = 61
        Exit Function
    End If

    ' IX ? /?/ ? id 36
    If g = "IX" Then
        ReglasMallorquin = 36
        Exit Function
    End If

    ' TJ / TG ? /d?/ ? id 51
    If g = "TJ" Or g = "TG" Then
        ReglasMallorquin = 51
        Exit Function
    End If

    ' IG final ? /t?/ ? id 50
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

    ' Vocal neutra (schwa) en sílaba átona ? /?/ ? id 11
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

    ' E tónica ? abierta /?/ (id 6), átona ? cerrada /e/ (id 5)
    If g = "E" Then
        If esTonica Then
            ReglasMallorquin = 6
        Else
            ReglasMallorquin = 5
        End If
        Exit Function
    End If

    ' O tónica ? abierta /?/ (id 8), átona ? cerrada /o/ (id 7)
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
    If g = "K" Or g = "C" Then ReglasValenciano = 30: Exit Function
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



' ============================================================
'   ReglasGalego (GAL)
'   Devuelve idFonema según la fonética del gallego.
'   Si no aplica, devuelve 0 para que el motor siga probando.
' ============================================================

Public Function ReglasGalego( _
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
        ReglasGalego = 57
        Exit Function
    End If

    ' GUE / GUI ? /g/ (U muda) ? id 31
    If g = "GUE" Or g = "GUI" Then
        ReglasGalego = 31
        Exit Function
    End If

    ' QUE / QUI ? /k/ ? id 30
    If g = "QUE" Or g = "QUI" Then
        ReglasGalego = 30
        Exit Function
    End If


    ' ============================================================
    '   DÍGRAFOS Y CASOS ESPECIALES
    ' ============================================================

    ' CH ? /t?/ ? id 50
    If g = "CH" Then
        ReglasGalego = 50
        Exit Function
    End If

    ' X ? /?/ ? id 36
    If g = "X" Then
        ReglasGalego = 36
        Exit Function
    End If

    ' J ? /?/ ? id 37
    If g = "J" Then
        ReglasGalego = 37
        Exit Function
    End If

    ' G + E/I ? /?/ ? id 37
    If g = "G" And (sig = "E" Or sig = "I") Then
        ReglasGalego = 37
        Exit Function
    End If

    ' LL ? /?/ ? id 44
    If g = "LL" Then
        ReglasGalego = 44
        Exit Function
    End If

    ' Ñ ? /?/ ? id 41
    If g = "Ñ" Then
        ReglasGalego = 41
        Exit Function
    End If

    ' RR ? /r/ múltiple ? id 46
    If g = "RR" Then
        ReglasGalego = 46
        Exit Function
    End If


    ' ============================================================
    '   DÍGRAFOS VOCÁLICOS (diptongos gallegos)
    ' ============================================================

    If g = "AI" Then ReglasGalego = 12: Exit Function
    If g = "EI" Then ReglasGalego = 13: Exit Function
    If g = "OI" Then ReglasGalego = 14: Exit Function
    If g = "AU" Then ReglasGalego = 16: Exit Function
    If g = "EU" Then ReglasGalego = 17: Exit Function
    If g = "OU" Then ReglasGalego = 15: Exit Function


    ' ============================================================
    '   MONÓGRAFOS — VOCALES (5 vocales)
    ' ============================================================

    If g = "A" Then ReglasGalego = 1: Exit Function
    If g = "E" Then ReglasGalego = 5: Exit Function
    If g = "I" Then ReglasGalego = 9: Exit Function
    If g = "O" Then ReglasGalego = 7: Exit Function
    If g = "U" Then ReglasGalego = 10: Exit Function


    ' ============================================================
    '   MONÓGRAFOS — CONSONANTES
    ' ============================================================

    If g = "P" Then ReglasGalego = 26: Exit Function
    If g = "B" Then ReglasGalego = 27: Exit Function
    If g = "T" Then ReglasGalego = 28: Exit Function
    If g = "D" Then ReglasGalego = 29: Exit Function
    If g = "K" Then ReglasGalego = 30: Exit Function
    If g = "G" Then ReglasGalego = 31: Exit Function

    If g = "F" Then ReglasGalego = 32: Exit Function

    ' S / Z / C+E/I ? /s/
    If g = "S" Then ReglasGalego = 34: Exit Function
    If g = "Z" Then ReglasGalego = 34: Exit Function
    If g = "C" And (sig = "E" Or sig = "I") Then
        ReglasGalego = 34
        Exit Function
    End If

    If g = "M" Then ReglasGalego = 39: Exit Function
    If g = "N" Then ReglasGalego = 40: Exit Function

    If g = "L" Then ReglasGalego = 43: Exit Function
    If g = "R" Then ReglasGalego = 45: Exit Function

    ' H ? aspiración suave ? id 38
    If g = "H" Then ReglasGalego = 38: Exit Function


    ' ============================================================
    '   SI NO APLICA, DEVOLVER 0
    ' ============================================================
    ReglasGalego = 0

End Function

Public Function ReglasPortugues( _
        ByVal graf As String, _
        ByVal ant As String, _
        ByVal sig As String, _
        ByVal esTonica As Boolean _
    ) As Byte
    
'Se mantiene por compatibilidad
ReglasPortugues = ReglasPortugues_PT_EU(graf, ant, sig, esTonica)

End Function

Public Function ReglasPortugues_PT_EU( _
        ByVal graf As String, _
        ByVal ant As String, _
        ByVal sig As String, _
        ByVal esTonica As Boolean _
    ) As Byte

' Versión KOSMOS

    Dim g As String
    g = UCase$(graf)

    ' ============================================================
    '   TRIGRAFEMAS
    ' ============================================================
    If g = "GÜE" Or g = "GÜI" Then ReglasPortugues_PT_EU = 57: Exit Function
    If g = "GUE" Or g = "GUI" Then ReglasPortugues_PT_EU = 31: Exit Function
    If g = "QUE" Or g = "QUI" Then ReglasPortugues_PT_EU = 30: Exit Function

    ' Nasales con vocal acentuada
    If g = "ÃO" Then ReglasPortugues_PT_EU = 2: Exit Function
    If g = "ÃE" Then ReglasPortugues_PT_EU = 2: Exit Function
    If g = "ÃI" Then ReglasPortugues_PT_EU = 2: Exit Function
    If g = "ÕE" Then ReglasPortugues_PT_EU = 4: Exit Function
    If g = "ÕI" Then ReglasPortugues_PT_EU = 4: Exit Function

    ' ============================================================
    '   DÍGRAFOS Y CASOS ESPECIALES
    ' ============================================================
    If g = "NH" Then ReglasPortugues_PT_EU = 41: Exit Function
    If g = "LH" Then ReglasPortugues_PT_EU = 44: Exit Function
    If g = "CH" Then ReglasPortugues_PT_EU = 36: Exit Function
    If g = "RR" Then ReglasPortugues_PT_EU = 47: Exit Function

    ' R inicial fuerte
    If g = "R" And ant = "" Then ReglasPortugues_PT_EU = 47: Exit Function

    ' SS ? /s/
    If g = "SS" Then ReglasPortugues_PT_EU = 34: Exit Function

    ' S entre vocales ? /z/
    If g = "S" And (ant Like "[AEIOUÃÕÁÉÍÓÚÂÊÔ]" And sig Like "[AEIOUÃÕÁÉÍÓÚÂÊÔ]") Then
        ReglasPortugues_PT_EU = 35: Exit Function
    End If

    ' S final ? /?/
    If g = "S" And sig = "" Then ReglasPortugues_PT_EU = 36: Exit Function

    ' X ? /?/ estándar
    If g = "X" Then ReglasPortugues_PT_EU = 36: Exit Function

    ' J ? /?/
    If g = "J" Then ReglasPortugues_PT_EU = 37: Exit Function

    ' G + E/I ? /?/
    If g = "G" And (sig = "E" Or sig = "I") Then ReglasPortugues_PT_EU = 37: Exit Function

    ' ============================================================
    '   NASALIZACIONES
    ' ============================================================

    ' Nasales internas (coda)
    If (g = "AN" Or g = "AM" Or g = "EN" Or g = "EM" _
     Or g = "IN" Or g = "IM" Or g = "ON" Or g = "OM" _
     Or g = "UN" Or g = "UM") _
     And Not (sig Like "[AEIOUÃÕÁÉÍÓÚÂÊÔ]") Then

        If g = "AN" Or g = "AM" Then ReglasPortugues_PT_EU = 2: Exit Function
        If g = "EN" Or g = "EM" Then ReglasPortugues_PT_EU = 3: Exit Function
        If g = "ON" Or g = "OM" Then ReglasPortugues_PT_EU = 4: Exit Function
        If g = "UN" Or g = "UM" Then ReglasPortugues_PT_EU = 11: Exit Function
    End If

    ' Nasales finales
    If (g = "AM" Or g = "AN") And sig = "" Then ReglasPortugues_PT_EU = 2: Exit Function
    If (g = "EM" Or g = "EN") And sig = "" Then ReglasPortugues_PT_EU = 3: Exit Function
    If (g = "OM" Or g = "ON") And sig = "" Then ReglasPortugues_PT_EU = 4: Exit Function

    ' ============================================================
    '   DÍGRAFOS VOCÁLICOS
    ' ============================================================
    If g = "AI" Then ReglasPortugues_PT_EU = 12: Exit Function
    If g = "EI" Then ReglasPortugues_PT_EU = 13: Exit Function
    If g = "OI" Then ReglasPortugues_PT_EU = 14: Exit Function
    If g = "OU" Then ReglasPortugues_PT_EU = 15: Exit Function
    If g = "AU" Then ReglasPortugues_PT_EU = 16: Exit Function
    If g = "EU" Then ReglasPortugues_PT_EU = 17: Exit Function
    If g = "UI" Then ReglasPortugues_PT_EU = 19: Exit Function

    ' ============================================================
    '   MONÓGRAFOS — VOCALES
    ' ============================================================
    If g = "A" Then ReglasPortugues_PT_EU = 1: Exit Function
    If g = "Á" Then ReglasPortugues_PT_EU = 1: Exit Function
    If g = "Â" Then ReglasPortugues_PT_EU = 1: Exit Function
    If g = "Ã" Then ReglasPortugues_PT_EU = 2: Exit Function

    If g = "E" Then ReglasPortugues_PT_EU = 5: Exit Function
    If g = "É" Then ReglasPortugues_PT_EU = 5: Exit Function
    If g = "Ê" Then ReglasPortugues_PT_EU = 5: Exit Function

    If g = "I" Then ReglasPortugues_PT_EU = 9: Exit Function
    If g = "Í" Then ReglasPortugues_PT_EU = 9: Exit Function

    If g = "O" Then ReglasPortugues_PT_EU = 7: Exit Function
    If g = "Ó" Then ReglasPortugues_PT_EU = 7: Exit Function
    If g = "Ô" Then ReglasPortugues_PT_EU = 7: Exit Function
    If g = "Õ" Then ReglasPortugues_PT_EU = 4: Exit Function

    If g = "U" Then ReglasPortugues_PT_EU = 10: Exit Function
    If g = "Ú" Then ReglasPortugues_PT_EU = 10: Exit Function

    ' ============================================================
    '   MONÓGRAFOS — CONSONANTES
    ' ============================================================
    If g = "P" Then ReglasPortugues_PT_EU = 26: Exit Function
    If g = "B" Then ReglasPortugues_PT_EU = 27: Exit Function
    If g = "T" Then ReglasPortugues_PT_EU = 28: Exit Function
    If g = "D" Then ReglasPortugues_PT_EU = 29: Exit Function
    If g = "K" Then ReglasPortugues_PT_EU = 30: Exit Function
    If g = "G" Then ReglasPortugues_PT_EU = 31: Exit Function
    If g = "F" Then ReglasPortugues_PT_EU = 32: Exit Function
    If g = "S" Then ReglasPortugues_PT_EU = 34: Exit Function
    If g = "M" Then ReglasPortugues_PT_EU = 39: Exit Function
    If g = "N" Then ReglasPortugues_PT_EU = 40: Exit Function
    If g = "L" Then ReglasPortugues_PT_EU = 43: Exit Function
    If g = "R" Then ReglasPortugues_PT_EU = 45: Exit Function
    If g = "H" Then ReglasPortugues_PT_EU = 38: Exit Function

    ReglasPortugues_PT_EU = 0

End Function


'Public Function ReglasPortugues_PT_EU( _
'        ByVal graf As String, _
'        ByVal ant As String, _
'        ByVal sig As String, _
'        ByVal esTonica As Boolean _
'    ) As Byte
'
'    Dim g As String
'    g = UCase$(graf)
'
'    ' TRIGRAFEMAS
'    If g = "GÜE" Or g = "GÜI" Then ReglasPortugues_PT_EU = 57: Exit Function
'    If g = "GUE" Or g = "GUI" Then ReglasPortugues_PT_EU = 31: Exit Function
'    If g = "QUE" Or g = "QUI" Then ReglasPortugues_PT_EU = 30: Exit Function
'
'    ' DÍGRAFOS Y CASOS ESPECIALES
'    If g = "NH" Then ReglasPortugues_PT_EU = 41: Exit Function
'    If g = "LH" Then ReglasPortugues_PT_EU = 44: Exit Function
'    If g = "CH" Then ReglasPortugues_PT_EU = 36: Exit Function
'    If g = "RR" Then ReglasPortugues_PT_EU = 47: Exit Function
'
'    ' R inicial fuerte
'    If g = "R" And ant = "" Then
'        ReglasPortugues_PT_EU = 47
'        Exit Function
'    End If
'
'    ' SS ? /s/
'    If g = "SS" Then ReglasPortugues_PT_EU = 34: Exit Function
'
'    ' S entre vocales ? /z/
'    If g = "S" And (ant Like "[AEIOU]" And sig Like "[AEIOU]") Then
'        ReglasPortugues_PT_EU = 35
'        Exit Function
'    End If
'
'    ' S final ? /?/ (norma europea)
'    If g = "S" And sig = "" Then
'        ReglasPortugues_PT_EU = 36
'        Exit Function
'    End If
'
'    ' X ? /?/ estándar
'    If g = "X" Then ReglasPortugues_PT_EU = 36: Exit Function
'
'    ' J ? /?/
'    If g = "J" Then ReglasPortugues_PT_EU = 37: Exit Function
'
'    ' G + E/I ? /?/
'    If g = "G" And (sig = "E" Or sig = "I") Then
'        ReglasPortugues_PT_EU = 37
'        Exit Function
'    End If
'
'    ' NASALIZACIONES
'    ' Nasales internas (coda)
'    If (g = "AN" Or g = "AM" Or g = "EN" Or g = "EM" _
'        Or g = "IN" Or g = "IM" Or g = "ON" Or g = "OM" _
'        Or g = "UN" Or g = "UM") _
'        And Not (sig Like "[AEIOU]") Then
'
'        ' AN/AM ? /ã/
'        If g = "AN" Or g = "AM" Then ReglasPortugues_PT_EU = 2: Exit Function
'
'        ' EN/EM ? /?/
'        If g = "EN" Or g = "EM" Then ReglasPortugues_PT_EU = 3: Exit Function
'
'        ' ON/OM ? /õ/
'        If g = "ON" Or g = "OM" Then ReglasPortugues_PT_EU = 4: Exit Function
'
'        ' UN/UM ? /u/
'        If g = "UN" Or g = "UM" Then ReglasPortugues_PT_EU = 11: Exit Function
'    End If
'
'
'    ' ÃO normalizado ? A~O
'    If g = "A~O" Then
'        ReglasPortugues_PT_EU = 2
'        Exit Function
'    End If
'
'    ' AM / AN final ? /ã/
'    If (g = "AM" Or g = "AN") And sig = "" Then
'        ReglasPortugues_PT_EU = 2
'        Exit Function
'    End If
'
'    ' EM / EN final
'    If (g = "EM" Or g = "EN") And sig = "" Then
'        ReglasPortugues_PT_EU = 3
'        Exit Function
'    End If
'
'    ' OM / ON final ? /õ/
'    If (g = "OM" Or g = "ON") And sig = "" Then
'        ReglasPortugues_PT_EU = 4
'        Exit Function
'    End If
'
'    ' DÍGRAFOS VOCÁLICOS
'    If g = "AI" Then ReglasPortugues_PT_EU = 12: Exit Function
'    If g = "EI" Then ReglasPortugues_PT_EU = 13: Exit Function
'    If g = "OI" Then ReglasPortugues_PT_EU = 14: Exit Function
'    If g = "OU" Then ReglasPortugues_PT_EU = 15: Exit Function
'    If g = "AU" Then ReglasPortugues_PT_EU = 16: Exit Function
'    If g = "EU" Then ReglasPortugues_PT_EU = 17: Exit Function
'    If g = "UI" Then ReglasPortugues_PT_EU = 19: Exit Function
'
'    ' MONÓGRAFOS — VOCALES (aquí luego podrás distinguir A, A´, Â, A~, etc.)
'    If g = "A" Then ReglasPortugues_PT_EU = 1: Exit Function
'    If g = "E" Then ReglasPortugues_PT_EU = 5: Exit Function
'    If g = "I" Then ReglasPortugues_PT_EU = 9: Exit Function
'    If g = "O" Then ReglasPortugues_PT_EU = 7: Exit Function
'    If g = "U" Then ReglasPortugues_PT_EU = 10: Exit Function
'
'    ' MONÓGRAFOS — CONSONANTES
'    If g = "P" Then ReglasPortugues_PT_EU = 26: Exit Function
'    If g = "B" Then ReglasPortugues_PT_EU = 27: Exit Function
'    If g = "T" Then ReglasPortugues_PT_EU = 28: Exit Function
'    If g = "D" Then ReglasPortugues_PT_EU = 29: Exit Function
'    If g = "K" Then ReglasPortugues_PT_EU = 30: Exit Function
'    If g = "G" Then ReglasPortugues_PT_EU = 31: Exit Function
'    If g = "F" Then ReglasPortugues_PT_EU = 32: Exit Function
'
'    ' S simple (no entre vocales, no final) ? /s/
'    If g = "S" Then ReglasPortugues_PT_EU = 34: Exit Function
'
'    If g = "M" Then ReglasPortugues_PT_EU = 39: Exit Function
'    If g = "N" Then ReglasPortugues_PT_EU = 40: Exit Function
'    If g = "L" Then ReglasPortugues_PT_EU = 43: Exit Function
'    If g = "R" Then ReglasPortugues_PT_EU = 45: Exit Function
'
'    ' H
'    If g = "H" Then ReglasPortugues_PT_EU = 38: Exit Function
'
'    ReglasPortugues_PT_EU = 0
'
'End Function

Public Function ReglasPortugues_PT_BR( _
        ByVal graf As String, _
        ByVal ant As String, _
        ByVal sig As String, _
        ByVal esTonica As Boolean _
    ) As Byte

' Versión KOSMOS

    Dim g As String
    g = UCase$(graf)

    ' ============================================================
    '   TRIGRAFEMAS
    ' ============================================================
    If g = "GÜE" Or g = "GÜI" Then ReglasPortugues_PT_BR = 57: Exit Function
    If g = "GUE" Or g = "GUI" Then ReglasPortugues_PT_BR = 31: Exit Function
    If g = "QUE" Or g = "QUI" Then ReglasPortugues_PT_BR = 30: Exit Function

    ' Nasales con vocal acentuada
    If g = "ÃO" Then ReglasPortugues_PT_BR = 2: Exit Function
    If g = "ÃE" Then ReglasPortugues_PT_BR = 2: Exit Function
    If g = "ÃI" Then ReglasPortugues_PT_BR = 2: Exit Function
    If g = "ÕE" Then ReglasPortugues_PT_BR = 4: Exit Function
    If g = "ÕI" Then ReglasPortugues_PT_BR = 4: Exit Function

    ' ============================================================
    '   DÍGRAFOS Y CASOS ESPECIALES
    ' ============================================================
    If g = "NH" Then ReglasPortugues_PT_BR = 41: Exit Function
    If g = "LH" Then ReglasPortugues_PT_BR = 44: Exit Function
    If g = "CH" Then ReglasPortugues_PT_BR = 36: Exit Function
    If g = "RR" Then ReglasPortugues_PT_BR = 47: Exit Function

    ' R inicial ? aspirado (lo mapeamos a H suave: 38)
    If g = "R" And ant = "" Then ReglasPortugues_PT_BR = 38: Exit Function

    ' SS ? /s/
    If g = "SS" Then ReglasPortugues_PT_BR = 34: Exit Function

    ' S entre vocales ? /z/
    If g = "S" And (ant Like "[AEIOUÃÕÁÉÍÓÚÂÊÔ]" And sig Like "[AEIOUÃÕÁÉÍÓÚÂÊÔ]") Then
        ReglasPortugues_PT_BR = 35: Exit Function
    End If

    ' S final ? /s/ (no /?/)
    If g = "S" And sig = "" Then ReglasPortugues_PT_BR = 34: Exit Function

    ' X ? /?/ estándar
    If g = "X" Then ReglasPortugues_PT_BR = 36: Exit Function

    ' J ? /?/
    If g = "J" Then ReglasPortugues_PT_BR = 37: Exit Function

    ' G + E/I ? /?/
    If g = "G" And (sig = "E" Or sig = "I") Then ReglasPortugues_PT_BR = 37: Exit Function

    ' ============================================================
    '   NASALIZACIONES
    ' ============================================================

    ' Nasales internas (coda)
    If (g = "AN" Or g = "AM" Or g = "EN" Or g = "EM" _
     Or g = "IN" Or g = "IM" Or g = "ON" Or g = "OM" _
     Or g = "UN" Or g = "UM") _
     And Not (sig Like "[AEIOUÃÕÁÉÍÓÚÂÊÔ]") Then

        If g = "AN" Or g = "AM" Then ReglasPortugues_PT_BR = 2: Exit Function
        If g = "EN" Or g = "EM" Then ReglasPortugues_PT_BR = 3: Exit Function
        If g = "ON" Or g = "OM" Then ReglasPortugues_PT_BR = 4: Exit Function
        If g = "UN" Or g = "UM" Then ReglasPortugues_PT_BR = 11: Exit Function
    End If

    ' Nasales finales
    If (g = "AM" Or g = "AN") And sig = "" Then ReglasPortugues_PT_BR = 2: Exit Function
    If (g = "EM" Or g = "EN") And sig = "" Then ReglasPortugues_PT_BR = 3: Exit Function
    If (g = "OM" Or g = "ON") And sig = "" Then ReglasPortugues_PT_BR = 4: Exit Function

    ' ============================================================
    '   DÍGRAFOS VOCÁLICOS
    ' ============================================================
    If g = "AI" Then ReglasPortugues_PT_BR = 12: Exit Function
    If g = "EI" Then ReglasPortugues_PT_BR = 13: Exit Function
    If g = "OI" Then ReglasPortugues_PT_BR = 14: Exit Function
    If g = "OU" Then ReglasPortugues_PT_BR = 15: Exit Function
    If g = "AU" Then ReglasPortugues_PT_BR = 16: Exit Function
    If g = "EU" Then ReglasPortugues_PT_BR = 17: Exit Function
    If g = "UI" Then ReglasPortugues_PT_BR = 19: Exit Function

    ' ============================================================
    '   MONÓGRAFOS — VOCALES
    ' ============================================================
    If g = "A" Then ReglasPortugues_PT_BR = 1: Exit Function
    If g = "Á" Then ReglasPortugues_PT_BR = 1: Exit Function
    If g = "Â" Then ReglasPortugues_PT_BR = 1: Exit Function
    If g = "Ã" Then ReglasPortugues_PT_BR = 2: Exit Function

    If g = "E" Then ReglasPortugues_PT_BR = 5: Exit Function
    If g = "É" Then ReglasPortugues_PT_BR = 5: Exit Function
    If g = "Ê" Then ReglasPortugues_PT_BR = 5: Exit Function

    If g = "I" Then ReglasPortugues_PT_BR = 9: Exit Function
    If g = "Í" Then ReglasPortugues_PT_BR = 9: Exit Function

    If g = "O" Then ReglasPortugues_PT_BR = 7: Exit Function
    If g = "Ó" Then ReglasPortugues_PT_BR = 7: Exit Function
    If g = "Ô" Then ReglasPortugues_PT_BR = 7: Exit Function
    If g = "Õ" Then ReglasPortugues_PT_BR = 4: Exit Function

    If g = "U" Then ReglasPortugues_PT_BR = 10: Exit Function
    If g = "Ú" Then ReglasPortugues_PT_BR = 10: Exit Function

    ' ============================================================
    '   MONÓGRAFOS — CONSONANTES
    ' ============================================================
    If g = "P" Then ReglasPortugues_PT_BR = 26: Exit Function
    If g = "B" Then ReglasPortugues_PT_BR = 27: Exit Function
    If g = "T" Then ReglasPortugues_PT_BR = 28: Exit Function
    If g = "D" Then ReglasPortugues_PT_BR = 29: Exit Function
    If g = "K" Then ReglasPortugues_PT_BR = 30: Exit Function
    If g = "G" Then ReglasPortugues_PT_BR = 31: Exit Function
    If g = "F" Then ReglasPortugues_PT_BR = 32: Exit Function
    If g = "S" Then ReglasPortugues_PT_BR = 34: Exit Function
    If g = "M" Then ReglasPortugues_PT_BR = 39: Exit Function
    If g = "N" Then ReglasPortugues_PT_BR = 40: Exit Function
    If g = "L" Then ReglasPortugues_PT_BR = 43: Exit Function
    If g = "R" Then ReglasPortugues_PT_BR = 45: Exit Function
    If g = "H" Then ReglasPortugues_PT_BR = 38: Exit Function

    ReglasPortugues_PT_BR = 0

End Function


'Public Function ReglasPortugues_PT_BR( _
'        ByVal graf As String, _
'        ByVal ant As String, _
'        ByVal sig As String, _
'        ByVal esTonica As Boolean _
'    ) As Byte
'
'    Dim g As String
'    g = UCase$(graf)
'
'    ' TRIGRAFEMAS
'    If g = "GÜE" Or g = "GÜI" Then ReglasPortugues_PT_BR = 57: Exit Function
'    If g = "GUE" Or g = "GUI" Then ReglasPortugues_PT_BR = 31: Exit Function
'    If g = "QUE" Or g = "QUI" Then ReglasPortugues_PT_BR = 30: Exit Function
'
'    ' DÍGRAFOS
'    If g = "NH" Then ReglasPortugues_PT_BR = 41: Exit Function
'    If g = "LH" Then ReglasPortugues_PT_BR = 44: Exit Function
'    If g = "CH" Then ReglasPortugues_PT_BR = 36: Exit Function
'    If g = "RR" Then ReglasPortugues_PT_BR = 47: Exit Function
'
'    ' R inicial ? más aspirado (lo mapeamos a H suave: 38)
'    If g = "R" And ant = "" Then
'        ReglasPortugues_PT_BR = 38
'        Exit Function
'    End If
'
'    ' SS ? /s/
'    If g = "SS" Then ReglasPortugues_PT_BR = 34: Exit Function
'
'    ' S entre vocales ? /z/
'    If g = "S" And (ant Like "[AEIOU]" And sig Like "[AEIOU]") Then
'        ReglasPortugues_PT_BR = 35
'        Exit Function
'    End If
'
'    ' S final ? /s/ (no /?/)
'    If g = "S" And sig = "" Then
'        ReglasPortugues_PT_BR = 34
'        Exit Function
'    End If
'
'    ' X (de momento igual que PT-EU)
'    If g = "X" Then ReglasPortugues_PT_BR = 36: Exit Function
'
'    ' J
'    If g = "J" Then ReglasPortugues_PT_BR = 37: Exit Function
'
'    ' G + E/I
'    If g = "G" And (sig = "E" Or sig = "I") Then
'        ReglasPortugues_PT_BR = 37
'        Exit Function
'    End If
'
'    ' NASALIZACIONES
'    ' Nasales internas (coda)
'    If (g = "AN" Or g = "AM" Or g = "EN" Or g = "EM" _
'        Or g = "IN" Or g = "IM" Or g = "ON" Or g = "OM" _
'        Or g = "UN" Or g = "UM") _
'        And Not (sig Like "[AEIOU]") Then
'
'        ' AN/AM ? /ã/ (más abierto)
'        If g = "AN" Or g = "AM" Then ReglasPortugues_PT_BR = 2: Exit Function
'
'        ' EN/EM ? /?/
'        If g = "EN" Or g = "EM" Then ReglasPortugues_PT_BR = 3: Exit Function
'
'        ' ON/OM ? /õ/
'        If g = "ON" Or g = "OM" Then ReglasPortugues_PT_BR = 4: Exit Function
'
'        ' UN/UM ? /u/
'        If g = "UN" Or g = "UM" Then ReglasPortugues_PT_BR = 11: Exit Function
'    End If
'
'
'    If g = "A~O" Then
'        ReglasPortugues_PT_BR = 2
'        Exit Function
'    End If
'
'    If (g = "AM" Or g = "AN") And sig = "" Then
'        ReglasPortugues_PT_BR = 2
'        Exit Function
'    End If
'
'    If (g = "EM" Or g = "EN") And sig = "" Then
'        ReglasPortugues_PT_BR = 3
'        Exit Function
'    End If
'
'    If (g = "OM" Or g = "ON") And sig = "" Then
'        ReglasPortugues_PT_BR = 4
'        Exit Function
'    End If
'
'    ' DÍGRAFOS VOCÁLICOS
'    If g = "AI" Then ReglasPortugues_PT_BR = 12: Exit Function
'    If g = "EI" Then ReglasPortugues_PT_BR = 13: Exit Function
'    If g = "OI" Then ReglasPortugues_PT_BR = 14: Exit Function
'    If g = "OU" Then ReglasPortugues_PT_BR = 15: Exit Function
'    If g = "AU" Then ReglasPortugues_PT_BR = 16: Exit Function
'    If g = "EU" Then ReglasPortugues_PT_BR = 17: Exit Function
'    If g = "UI" Then ReglasPortugues_PT_BR = 19: Exit Function
'
'    ' MONÓGRAFOS — VOCALES
'    If g = "A" Then ReglasPortugues_PT_BR = 1: Exit Function
'    If g = "E" Then ReglasPortugues_PT_BR = 5: Exit Function
'    If g = "I" Then ReglasPortugues_PT_BR = 9: Exit Function
'    If g = "O" Then ReglasPortugues_PT_BR = 7: Exit Function
'    If g = "U" Then ReglasPortugues_PT_BR = 10: Exit Function
'
'    ' MONÓGRAFOS — CONSONANTES
'    If g = "P" Then ReglasPortugues_PT_BR = 26: Exit Function
'    If g = "B" Then ReglasPortugues_PT_BR = 27: Exit Function
'    If g = "T" Then ReglasPortugues_PT_BR = 28: Exit Function
'    If g = "D" Then ReglasPortugues_PT_BR = 29: Exit Function
'    If g = "K" Then ReglasPortugues_PT_BR = 30: Exit Function
'    If g = "G" Then ReglasPortugues_PT_BR = 31: Exit Function
'    If g = "F" Then ReglasPortugues_PT_BR = 32: Exit Function
'
'    If g = "S" Then ReglasPortugues_PT_BR = 34: Exit Function
'    If g = "M" Then ReglasPortugues_PT_BR = 39: Exit Function
'    If g = "N" Then ReglasPortugues_PT_BR = 40: Exit Function
'    If g = "L" Then ReglasPortugues_PT_BR = 43: Exit Function
'    If g = "R" Then ReglasPortugues_PT_BR = 45: Exit Function
'
'    If g = "H" Then ReglasPortugues_PT_BR = 38: Exit Function
'
'    ReglasPortugues_PT_BR = 0
'
'End Function



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


' ============================================================
'   ReglasIngles (ENG)
'   Devuelve idFonema según la fonética del inglés.
'   Si no aplica, devuelve 0 para que el motor siga probando.
' ============================================================

Public Function ReglasIngles( _
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

    ' GUE / GUI ? /g/ (U muda)
    If g = "GUE" Or g = "GUI" Then
        ReglasIngles = 31
        Exit Function
    End If

    ' QUE / QUI ? /k/
    If g = "QUE" Or g = "QUI" Then
        ReglasIngles = 30
        Exit Function
    End If


    ' ============================================================
    '   DÍGRAFOS ESPECIALES
    ' ============================================================

    ' TH ? /?/
    If g = "TH" Then
        ReglasIngles = 54
        Exit Function
    End If

    ' DH ? /ð/
    If g = "DH" Then
        ReglasIngles = 55
        Exit Function
    End If

    ' SH ? /?/
    If g = "SH" Then
        ReglasIngles = 36
        Exit Function
    End If

    ' CH ? /t?/
    If g = "CH" Then
        ReglasIngles = 50
        Exit Function
    End If

    ' PH ? /f/
    If g = "PH" Then
        ReglasIngles = 32
        Exit Function
    End If

    ' NG ? /?/
    If g = "NG" Then
        ReglasIngles = 42
        Exit Function
    End If

    ' WH ? /w/
    If g = "WH" Then
        ReglasIngles = 49
        Exit Function
    End If


    ' ============================================================
    '   DÍGRAFOS VOCÁLICOS (diptongos ingleses)
    ' ============================================================

    If g = "AI" Or g = "AY" Then ReglasIngles = 13: Exit Function
    If g = "EI" Then ReglasIngles = 13: Exit Function
    If g = "OI" Or g = "OY" Then ReglasIngles = 14: Exit Function
    If g = "OU" Or g = "OW" Then ReglasIngles = 15: Exit Function
    If g = "AU" Or g = "AW" Then ReglasIngles = 16: Exit Function

    ' EA / EE ? /i/
    If g = "EA" Or g = "EE" Then ReglasIngles = 9: Exit Function

    ' IE ? /ai/
    If g = "IE" Then ReglasIngles = 12: Exit Function


    ' ============================================================
    '   MONÓGRAFOS — VOCALES
    ' ============================================================

    If g = "A" Then ReglasIngles = 1: Exit Function
    If g = "E" Then ReglasIngles = 5: Exit Function
    If g = "I" Then ReglasIngles = 9: Exit Function
    If g = "O" Then ReglasIngles = 7: Exit Function
    If g = "U" Then ReglasIngles = 10: Exit Function


    ' ============================================================
    '   MONÓGRAFOS — CONSONANTES
    ' ============================================================

    If g = "P" Then ReglasIngles = 26: Exit Function
    If g = "B" Then ReglasIngles = 27: Exit Function
    If g = "T" Then ReglasIngles = 28: Exit Function
    If g = "D" Then ReglasIngles = 29: Exit Function
    If g = "K" Or g = "C" Then ReglasIngles = 30: Exit Function
    If g = "G" Then ReglasIngles = 31: Exit Function

    If g = "F" Then ReglasIngles = 32: Exit Function
    If g = "V" Then ReglasIngles = 33: Exit Function
    If g = "S" Then ReglasIngles = 34: Exit Function
    If g = "Z" Then ReglasIngles = 35: Exit Function

    ' J ? /d?/
    If g = "J" Then ReglasIngles = 51: Exit Function

    ' Y ? /j/
    If g = "Y" Then ReglasIngles = 48: Exit Function

    ' W ? /w/
    If g = "W" Then ReglasIngles = 49: Exit Function

    ' X ? /ks/ ? devolvemos /s/
    If g = "X" Then ReglasIngles = 34: Exit Function

    If g = "M" Then ReglasIngles = 39: Exit Function
    If g = "N" Then ReglasIngles = 40: Exit Function

    If g = "L" Then ReglasIngles = 43: Exit Function
    If g = "R" Then ReglasIngles = 45: Exit Function

    ' H ? aspiración suave
    If g = "H" Then ReglasIngles = 38: Exit Function


    ' ============================================================
    '   SI NO APLICA, DEVOLVER 0
    ' ============================================================
    ReglasIngles = 0

End Function

