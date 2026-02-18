Attribute VB_Name = "modMotorFonetico_V2_1_Idiomas"

Option Compare Database
Option Explicit

'Public Function ReglasPortugues( _
'        ByVal graf As String, _
'        ByVal ant As String, _
'        ByVal sig As String, _
'        ByVal esTonica As Boolean _
'    ) As Byte
'
''Se mantiene por compatibilidad
'ReglasPortugues = ReglasPortugues_PT_EU(graf, ant, sig, esTonica)
'
'End Function





'' ============================================================
''   ReglasIngles_EN_US
''   Motor fonético para inglés americano (EN-US)
''   Extiende ReglasIngles (EN-GB) con:
''       - R-coloring (r siempre pronunciada)
''       - Flapping (t/d --> /-->/ entre vocales)
''       - Ajustes vocálicos típicos del inglés americano
'' ============================================================
'
'Public Function ReglasIngles_EN_US( _
'        ByVal graf As String, _
'        ByVal ant As String, _
'        ByVal sig As String, _
'        ByVal esTonica As Boolean _
'    ) As Byte
'
'    Dim g As String
'    g = UCase$(graf)
'
'    ' ============================================================
'    '   1. R-COLORING (r siempre pronunciada)
'    ' ============================================================
'
'    ' AR --> /?r/
'    If g = "AR" Then
'        ReglasIngles_EN_US = 60   ' tu fonema /?r/
'        Exit Function
'    End If
'
'    ' ER --> /?/ (schwa+r tónica)
'    If g = "ER" Then
'        ReglasIngles_EN_US = 61   ' tu fonema /?/
'        Exit Function
'    End If
'
'    ' OR --> /?r/ o /o?/
'    If g = "OR" Then
'        ReglasIngles_EN_US = 62   ' tu fonema /?r/
'        Exit Function
'    End If
'
'    ' IR / UR --> /?/ o /?/
'    If g = "IR" Or g = "UR" Then
'        ReglasIngles_EN_US = 61
'        Exit Function
'    End If
'
'
'    ' ============================================================
'    '   2. FLAPPING (T/D entre vocales --> /?/)
'    ' ============================================================
'
'    If (g = "T" Or g = "D") _
'       And EsVocal(ant) _
'       And EsVocal(sig) Then
'
'        ReglasIngles_EN_US = 63   ' tu fonema /?/
'        Exit Function
'    End If
'
'
'    ' ============================================================
'    '   3. AJUSTES VOCÁLICOS (cot–caught merger)
'    ' ============================================================
'
'    If g = "O" Then
'        ReglasIngles_EN_US = 7    ' /?/ en la mayoría de dialectos US
'        Exit Function
'    End If
'
'
'    ' ============================================================
'    '   4. SI NO APLICA, USAR MOTOR BASE (EN-GB)
'    ' ============================================================
'
'    ReglasIngles_EN_US = ReglasIngles(graf, ant, sig, esTonica)
'
'End Function
'
'' ============================================================
''   ReglasIngles_EN_US_AF
''   Motor fonético para nombres extranjeros en contexto EN-US
''   Extiende ReglasIngles_EN_US con:
''       - LL --> /j/
''       - Ñ --> /nj/
''       - Vocales acentuadas normalizadas
'' ============================================================
'
'Public Function ReglasIngles_EN_US_AF( _
'        ByVal graf As String, _
'        ByVal ant As String, _
'        ByVal sig As String, _
'        ByVal esTonica As Boolean _
'    ) As Byte
'
'    Dim g As String
'    g = UCase$(graf)
'
'    ' ============================================================
'    '   1. APLICAR MOTOR EN-US
'    ' ============================================================
'
'    Dim r As Byte
'    r = ReglasIngles_EN_US(graf, ant, sig, esTonica)
'
'    If r <> 0 Then
'        ReglasIngles_EN_US_AF = r
'        Exit Function
'    End If
'
'
'    ' ============================================================
'    '   2. REGLAS PARA NOMBRES EXTRANJEROS
'    ' ============================================================
'
'    ' LL --> /j/ (como en español)
'    If g = "LL" Then
'        ReglasIngles_EN_US_AF = 48   ' /j/
'        Exit Function
'    End If
'
'    ' Ñ --> /nj/
'    If g = "Ñ" Then
'        ReglasIngles_EN_US_AF = 40   ' /n/ (tu sistema añade /j/ después)
'        Exit Function
'    End If
'
'    ' Ç --> /s/
'    If g = "Ç" Then
'        ReglasIngles_EN_US_AF = 34   ' /s/
'        Exit Function
'    End If
'
'    ' Vocales acentuadas (por robustez)
'    If g = "Á" Then ReglasIngles_EN_US_AF = 1: Exit Function
'    If g = "É" Then ReglasIngles_EN_US_AF = 5: Exit Function
'    If g = "Í" Then ReglasIngles_EN_US_AF = 9: Exit Function
'    If g = "Ó" Then ReglasIngles_EN_US_AF = 7: Exit Function
'    If g = "Ú" Then ReglasIngles_EN_US_AF = 10: Exit Function
'
'
'    ' ============================================================
'    '   3. FALLBACK FINAL
'    ' ============================================================
'
'    ReglasIngles_EN_US_AF = 0
'
'End Function


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
'    ' SS --> /s/
'    If g = "SS" Then ReglasPortugues_PT_EU = 34: Exit Function
'
'    ' S entre vocales --> /z/
'    If g = "S" And (ant Like "[AEIOU]" And sig Like "[AEIOU]") Then
'        ReglasPortugues_PT_EU = 35
'        Exit Function
'    End If
'
'    ' S final --> /?/ (norma europea)
'    If g = "S" And sig = "" Then
'        ReglasPortugues_PT_EU = 36
'        Exit Function
'    End If
'
'    ' X --> /?/ estándar
'    If g = "X" Then ReglasPortugues_PT_EU = 36: Exit Function
'
'    ' J --> /?/
'    If g = "J" Then ReglasPortugues_PT_EU = 37: Exit Function
'
'    ' G + E/I --> /?/
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
'        ' AN/AM --> /ã/
'        If g = "AN" Or g = "AM" Then ReglasPortugues_PT_EU = 2: Exit Function
'
'        ' EN/EM --> /?/
'        If g = "EN" Or g = "EM" Then ReglasPortugues_PT_EU = 3: Exit Function
'
'        ' ON/OM --> /õ/
'        If g = "ON" Or g = "OM" Then ReglasPortugues_PT_EU = 4: Exit Function
'
'        ' UN/UM --> /u/
'        If g = "UN" Or g = "UM" Then ReglasPortugues_PT_EU = 11: Exit Function
'    End If
'
'
'    ' ÃO normalizado --> A~O
'    If g = "A~O" Then
'        ReglasPortugues_PT_EU = 2
'        Exit Function
'    End If
'
'    ' AM / AN final --> /ã/
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
'    ' OM / ON final --> /õ/
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
'    ' S simple (no entre vocales, no final) --> /s/
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
'    ' R inicial --> más aspirado (lo mapeamos a H suave: 38)
'    If g = "R" And ant = "" Then
'        ReglasPortugues_PT_BR = 38
'        Exit Function
'    End If
'
'    ' SS --> /s/
'    If g = "SS" Then ReglasPortugues_PT_BR = 34: Exit Function
'
'    ' S entre vocales --> /z/
'    If g = "S" And (ant Like "[AEIOU]" And sig Like "[AEIOU]") Then
'        ReglasPortugues_PT_BR = 35
'        Exit Function
'    End If
'
'    ' S final --> /s/ (no /?/)
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
'        ' AN/AM --> /ã/ (más abierto)
'        If g = "AN" Or g = "AM" Then ReglasPortugues_PT_BR = 2: Exit Function
'
'        ' EN/EM --> /?/
'        If g = "EN" Or g = "EM" Then ReglasPortugues_PT_BR = 3: Exit Function
'
'        ' ON/OM --> /õ/
'        If g = "ON" Or g = "OM" Then ReglasPortugues_PT_BR = 4: Exit Function
'
'        ' UN/UM --> /u/
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

