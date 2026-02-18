Attribute VB_Name = "modMotor_Idioma_PT_EU"

Option Compare Database
Option Explicit

'==================
'== PortuguÈs EU ==
'==================
Public Sub MF_SilabearAjustesPT_EU( _
        ByVal Texto As String, _
        ByRef silabas As Collection _
    )

    Dim i As Long

    ' ============================================================
    ' 1. ATAQUES CONSON¡NTICOS PERMITIDOS EN PT-PT
    ' ============================================================
    Dim ataques As Variant
    ataques = Array( _
        "BR", "BL", "CR", "CL", "DR", "FR", "FL", _
        "GR", "GL", "PR", "PL", "TR" _
    )

    For i = 2 To Len(Texto) - 1
        Dim par As String
        par = Mid$(Texto, i, 2)

        If EsMiembro(par, ataques) Then
            Call MF_UnirConsonantesEnAtaque(silabas, i)
        End If
    Next i

    ' ============================================================
    ' 2. DIPTONGOS NASAIS (PT-PT)
    ' ============================================================
    Dim diptNasal As Variant
    diptNasal = Array("√O", "√E", "’E", "’I")

    For i = 2 To Len(Texto)
        Dim dn As String
        dn = Mid$(Texto, i - 1, 2)

        If EsMiembro(dn, diptNasal) Then
            Call MF_UnirVocalesEnDiptongo(silabas, i)
        End If
    Next i

    ' ============================================================
    ' 3. VOCALES NASAIS FINALES (am, em, im, om, um)
    '    ? deben ser UNA sola sÌlaba
    ' ============================================================
    Dim nasalesFinales As Variant
    nasalesFinales = Array("AM", "EM", "IM", "OM", "UM")

    For i = 2 To Len(Texto)
        Dim nf As String
        nf = Mid$(Texto, i - 1, 2)

        If EsMiembro(nf, nasalesFinales) And i = Len(Texto) Then
            Call MF_UnirVocalesEnDiptongo(silabas, i)
        End If
    Next i

    ' ============================================================
    ' 4. HIATOS PT-PT (ea, eo, oa, oe, ui, iu)
    ' ============================================================
    Dim hiatos As Variant
    hiatos = Array("EA", "EO", "OA", "OE", "UI", "IU")

    For i = 2 To Len(Texto)
        Dim hv As String
        hv = Mid$(Texto, i - 1, 2)

        If EsMiembro(hv, hiatos) Then
            Call MF_ForzarDivisionSilabica(silabas, i)
        End If
    Next i

    ' ============================================================
    ' 5. DIPTONGOS ORALES PT-PT (estables)
    ' ============================================================
    Dim dipt As Variant
    dipt = Array( _
        "AI", "EI", "OI", "UI", _
        "AU", "EU", "OU" _
    )

    For i = 2 To Len(Texto)
        Dim dv As String
        dv = Mid$(Texto, i - 1, 2)

        If EsMiembro(dv, dipt) Then
            Call MF_UnirVocalesEnDiptongo(silabas, i)
        End If
    Next i

    ' ============================================================
    ' 6. RR y SS intervoc·licas ? no dividir
    ' ============================================================
    For i = 2 To Len(Texto) - 1
        Dim seq As String
        seq = Mid$(Texto, i, 2)

        If seq = "RR" Or seq = "SS" Then
            Call MF_UnirConsonantesEnAtaque(silabas, i)
        End If
    Next i

End Sub

Public Sub MF_MarcarTonicaPortuguesEU( _
        ByVal Texto As String, _
        ByRef esTonica() As Boolean _
    )

    Dim silabas As Collection
    Dim idxTonica As Long
    Dim inicio As Long, fin As Long
    Dim i As Long
    Dim ultima As String
    Dim ult2 As String
    Dim vocalesTilde As String

    vocalesTilde = "¡…Õ”⁄¬ ‘"

    ' --------------------------------------------------------
    ' 1. Silabear palabra
    ' --------------------------------------------------------
    'Set silabas = MF_SilabearCastellano(Texto)
    Set silabas = MF_Silabear(Texto, "pt-eu")


    If silabas Is Nothing Then Exit Sub
    If silabas.Count = 0 Then Exit Sub

    ultima = Right$(Texto, 1)
    If Len(Texto) >= 2 Then ult2 = Right$(Texto, 2)

    ' --------------------------------------------------------
    ' 2. Si hay tilde ? esa sÌlaba
    ' --------------------------------------------------------
    For i = 1 To Len(Texto)
        If InStr(vocalesTilde, Mid$(Texto, i, 1)) > 0 Then
            idxTonica = MF_SilabaDeIndice(i, silabas)
            GoTo Marcar
        End If
    Next i

    ' --------------------------------------------------------
    ' 3. Reglas generales PT-PT
    ' --------------------------------------------------------

    ' 3.1. Terminaciones oxÌtonas tÌpicas
    If ultima = "L" Or ultima = "R" Or ultima = "Z" Then
        idxTonica = silabas.Count
        GoTo Marcar
    End If

    If ult2 = "IM" Or ult2 = "UM" Then
        idxTonica = silabas.Count
        GoTo Marcar
    End If

    ' 3.2. Terminaciones paroxÌtonas
    If InStr("AEIOU", ultima) > 0 Or _
       ult2 = "AS" Or ult2 = "ES" Or ult2 = "OS" Or _
       ult2 = "AM" Or ult2 = "EM" Then

        If silabas.Count = 1 Then
            idxTonica = 1
        Else
            idxTonica = silabas.Count - 1
        End If

        GoTo Marcar
    End If

    ' 3.3. Resto ? oxÌtona
    idxTonica = silabas.Count

Marcar:
    inicio = silabas(idxTonica)(1)
    fin = silabas(idxTonica)(2)

    For i = inicio To fin
        esTonica(i) = True
    Next i

End Sub

' ============================================================
'   ReglasPortugues (PT_EU)
'   Devuelve idFonema seg˙n la fonÈtica del francÈs.
'   Si no aplica, devuelve 0 para que el motor siga probando.
' ============================================================
Public Function ReglasPortugues_PT_EU( _
        ByVal graf As String, _
        ByVal ant As String, _
        ByVal sig As String, _
        ByVal esTonica As Boolean _
    ) As Byte

' VersiÛn KOSMOS

    Dim g As String
    g = UCase$(graf)

    ' ============================================================
    '   TRIGRAFEMAS
    ' ============================================================
    If g = "G‹E" Or g = "G‹I" Then ReglasPortugues_PT_EU = 57: Exit Function
    If g = "GUE" Or g = "GUI" Then ReglasPortugues_PT_EU = 31: Exit Function
    If g = "QUE" Or g = "QUI" Then ReglasPortugues_PT_EU = 30: Exit Function

    ' Nasales con vocal acentuada
    If g = "√O" Then ReglasPortugues_PT_EU = 2: Exit Function
    If g = "√E" Then ReglasPortugues_PT_EU = 2: Exit Function
    If g = "√I" Then ReglasPortugues_PT_EU = 2: Exit Function
    If g = "’E" Then ReglasPortugues_PT_EU = 4: Exit Function
    If g = "’I" Then ReglasPortugues_PT_EU = 4: Exit Function

    ' ============================================================
    '   DÕGRAFOS Y CASOS ESPECIALES
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
    If g = "S" And (ant Like "[AEIOU√’¡…Õ”⁄¬ ‘]" And sig Like "[AEIOU√’¡…Õ”⁄¬ ‘]") Then
        ReglasPortugues_PT_EU = 35: Exit Function
    End If

    ' S final ? /?/
    If g = "S" And sig = "" Then ReglasPortugues_PT_EU = 36: Exit Function

    ' X ? /?/ est·ndar
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
     And Not (sig Like "[AEIOU√’¡…Õ”⁄¬ ‘]") Then

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
    '   DÕGRAFOS VOC¡LICOS
    ' ============================================================
    If g = "AI" Then ReglasPortugues_PT_EU = 12: Exit Function
    If g = "EI" Then ReglasPortugues_PT_EU = 13: Exit Function
    If g = "OI" Then ReglasPortugues_PT_EU = 14: Exit Function
    If g = "OU" Then ReglasPortugues_PT_EU = 15: Exit Function
    If g = "AU" Then ReglasPortugues_PT_EU = 16: Exit Function
    If g = "EU" Then ReglasPortugues_PT_EU = 17: Exit Function
    If g = "UI" Then ReglasPortugues_PT_EU = 19: Exit Function

    ' ============================================================
    '   MON”GRAFOS ó VOCALES
    ' ============================================================
    If g = "A" Then ReglasPortugues_PT_EU = 1: Exit Function
    If g = "¡" Then ReglasPortugues_PT_EU = 1: Exit Function
    If g = "¬" Then ReglasPortugues_PT_EU = 1: Exit Function
    If g = "√" Then ReglasPortugues_PT_EU = 2: Exit Function

    If g = "E" Then ReglasPortugues_PT_EU = 5: Exit Function
    If g = "…" Then ReglasPortugues_PT_EU = 5: Exit Function
    If g = " " Then ReglasPortugues_PT_EU = 5: Exit Function

    If g = "I" Then ReglasPortugues_PT_EU = 9: Exit Function
    If g = "Õ" Then ReglasPortugues_PT_EU = 9: Exit Function

    If g = "O" Then ReglasPortugues_PT_EU = 7: Exit Function
    If g = "”" Then ReglasPortugues_PT_EU = 7: Exit Function
    If g = "‘" Then ReglasPortugues_PT_EU = 7: Exit Function
    If g = "’" Then ReglasPortugues_PT_EU = 4: Exit Function

    If g = "U" Then ReglasPortugues_PT_EU = 10: Exit Function
    If g = "⁄" Then ReglasPortugues_PT_EU = 10: Exit Function

    ' ============================================================
    '   MON”GRAFOS ó CONSONANTES
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

'============================================================================================

Public Function MF_NormalizarVocales_PT_EU(ByVal Texto As String) As String

    ' Nasales
    Texto = Replace(Texto, "√", "A~")
    Texto = Replace(Texto, "’", "O~")

    ' Cerradas (circunflejo)
    Texto = Replace(Texto, "¬", "¬")
    Texto = Replace(Texto, " ", " ")
    Texto = Replace(Texto, "Œ", "I") ' no existe en PT, pero por robustez
    Texto = Replace(Texto, "‘", "‘")
    Texto = Replace(Texto, "€", "U") ' no existe en PT, robustez

    ' Abiertas (agudas)
    Texto = Replace(Texto, "¡", "A¥")
    Texto = Replace(Texto, "…", "E¥")
    Texto = Replace(Texto, "Õ", "I¥")
    Texto = Replace(Texto, "”", "O¥")
    Texto = Replace(Texto, "⁄", "U¥")

    ' Graves (no existen en PT, pero pueden aparecer en nombres importados)
    Texto = Replace(Texto, "¿", "A")
    Texto = Replace(Texto, "»", "E")
    Texto = Replace(Texto, "Ã", "I")
    Texto = Replace(Texto, "“", "O")
    Texto = Replace(Texto, "Ÿ", "U")

    MF_NormalizarVocales_PT_EU = Texto

End Function

