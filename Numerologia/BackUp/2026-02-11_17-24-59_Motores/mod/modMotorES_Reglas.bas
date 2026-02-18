Attribute VB_Name = "modMotorES_Reglas"
'
'' ============================
''  modMotorES_Reglas
''  Reglas fonéticas del español
''  (grafema ? IdFonema)
'' ============================
'
'Option Compare Database
'Option Explicit
'
'' ============================================================
''   ReglasCastellano
''   Devuelve idFonema según la fonética del castellano.
''   Si no aplica, devuelve 0 para que el motor siga probando.
'' ============================================================
'Public Function ReglasCastellano( _
'        ByVal graf As String, _
'        ByVal ant As String, _
'        ByVal sig As String, _
'        ByVal esTonica As Boolean _
'    ) As Byte
'
'    Dim g As String
'    g = LCase$(graf)
'
'    ' ============================================================
'    '   TRIGRAFEMAS
'    ' ============================================================
'
'    ' güe / güi ? id 57
'    If g = "güe" Or g = "güi" Then
'        ReglasCastellano = 57
'        Exit Function
'    End If
'
'    ' gue / gui ? id 31
'    If g = "gue" Or g = "gui" Then
'        ReglasCastellano = 31
'        Exit Function
'    End If
'
'    ' que / qui ? id 30
'    If g = "que" Or g = "qui" Then
'        ReglasCastellano = 30
'        Exit Function
'    End If
'
'
'    ' ============================================================
'    '   DÍGRAFOS Y CASOS ESPECIALES
'    ' ============================================================
'
'    If g = "ch" Then ReglasCastellano = 50: Exit Function
'    If g = "ll" Then ReglasCastellano = 44: Exit Function
'    If g = "rr" Then ReglasCastellano = 46: Exit Function
'    If g = "ñ" Then ReglasCastellano = 41: Exit Function
'
'    ' gu + vocal ? /g/
'    If g = "gu" And (sig = "a" Or sig = "o" Or sig = "u") Then
'        ReglasCastellano = 31
'        Exit Function
'    End If
'
'    ' qu + vocal ? /k/
'    If g = "qu" And (sig = "a" Or sig = "o" Or sig = "u") Then
'        ReglasCastellano = 30
'        Exit Function
'    End If
'
'
'    ' ============================================================
'    '   DÍGRAFOS VOCÁLICOS (diptongos)
'    ' ============================================================
'
'    If g = "ai" Then ReglasCastellano = 12: Exit Function
'    If g = "ei" Then ReglasCastellano = 13: Exit Function
'    If g = "oi" Then ReglasCastellano = 14: Exit Function
'    If g = "ou" Then ReglasCastellano = 15: Exit Function
'    If g = "au" Then ReglasCastellano = 16: Exit Function
'
'
'    ' ============================================================
'    '   MONÓGRAFOS — VOCALES
'    ' ============================================================
'
'    If g = "a" Then ReglasCastellano = 1: Exit Function
'    If g = "e" Then ReglasCastellano = 5: Exit Function
'    If g = "i" Then ReglasCastellano = 9: Exit Function
'    If g = "o" Then ReglasCastellano = 7: Exit Function
'    If g = "u" Then ReglasCastellano = 10: Exit Function
'
'
'    ' ============================================================
'    '   MONÓGRAFOS — CONSONANTES
'    ' ============================================================
'
'    If g = "p" Then ReglasCastellano = 26: Exit Function
'    If g = "b" Or g = "v" Then ReglasCastellano = 27: Exit Function
'    If g = "t" Then ReglasCastellano = 28: Exit Function
'    If g = "d" Then ReglasCastellano = 29: Exit Function
'    If g = "k" Then ReglasCastellano = 30: Exit Function
'    If g = "g" Then ReglasCastellano = 31: Exit Function
'
'    If g = "f" Then ReglasCastellano = 32: Exit Function
'
'    ' c/z ? /?/ (castellano estándar)
'    If g = "c" And (sig = "e" Or sig = "i") Then
'        ReglasCastellano = 54
'        Exit Function
'    End If
'    If g = "z" Then
'        ReglasCastellano = 54
'        Exit Function
'    End If
'
'    ' s ? /s/
'    If g = "s" Then ReglasCastellano = 34: Exit Function
'
'    ' j / g + e/i ? /x/
'    If g = "j" Then ReglasCastellano = 58: Exit Function
'    If g = "g" And (sig = "e" Or sig = "i") Then
'        ReglasCastellano = 58
'        Exit Function
'    End If
'
'    ' m / n
'    If g = "m" Then ReglasCastellano = 39: Exit Function
'    If g = "n" Then ReglasCastellano = 40: Exit Function
'
'    ' l / r simple
'    If g = "l" Then ReglasCastellano = 43: Exit Function
'    If g = "r" Then ReglasCastellano = 45: Exit Function
'
'    ' h muda
'    If g = "h" Then ReglasCastellano = 38: Exit Function
'
'
'    ' ============================================================
'    '   SI NO APLICA
'    ' ============================================================
'    ReglasCastellano = 0
'
'End Function
'
'
'
'
'
'Public Function MF_NormalizarVocales_ES(ByVal texto As String) As String
'
'    texto = Replace(texto, "á", "a")
'    texto = Replace(texto, "à", "a")
'    texto = Replace(texto, "ä", "a")
'    texto = Replace(texto, "â", "a")
'
'    texto = Replace(texto, "é", "e")
'    texto = Replace(texto, "è", "e")
'    texto = Replace(texto, "ë", "e")
'    texto = Replace(texto, "ê", "e")
'
'    texto = Replace(texto, "í", "i")
'    texto = Replace(texto, "ì", "i")
'    texto = Replace(texto, "ï", "i")
'    texto = Replace(texto, "î", "i")
'
'    texto = Replace(texto, "ó", "o")
'    texto = Replace(texto, "ò", "o")
'    texto = Replace(texto, "ö", "o")
'    texto = Replace(texto, "ô", "o")
'
'    texto = Replace(texto, "ú", "u")
'    texto = Replace(texto, "ù", "u")
'    texto = Replace(texto, "û", "u")
'
'    ' ü se mantiene
'    MF_NormalizarVocales_ES = texto
'
'End Function
'
'
