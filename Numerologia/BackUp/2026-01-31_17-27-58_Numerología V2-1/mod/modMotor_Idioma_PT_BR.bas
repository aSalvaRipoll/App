Attribute VB_Name = "modMotor_Idioma_PT_BR"

Option Compare Database
Option Explicit

'==================
'== Portugués BR ==
'==================
Public Sub MF_SilabearAjustesPT_BR( _
        ByVal Texto As String, _
        ByRef silabas As Collection _
    )

    Dim i As Long

    ' ============================================================
    ' 1. ATAQUES CONSONÁNTICOS PERMITIDOS EN PT-BR
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
    ' 2. DIPTONGOS NASAIS (muy estables)
    ' ============================================================
    Dim diptNasal As Variant
    diptNasal = Array("ÃO", "ÃE", "ÕE", "ÕI", "ÃI", "ÕU")

    For i = 2 To Len(Texto)
        Dim dn As String
        dn = Mid$(Texto, i - 1, 2)

        If EsMiembro(dn, diptNasal) Then
            Call MF_UnirVocalesEnDiptongo(silabas, i)
        End If
    Next i

    ' ============================================================
    ' 3. DIPTONGOS ORALES (muy estables)
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
    ' 4. HIATOS OBLIGATORIOS (aí, aí, oá, oé, oê, eí, eú…)
    ' ============================================================
    Dim hiatos As Variant
    hiatos = Array("AÍ", "AÌ", "AÚ", "AÓ", _
                   "EÍ", "EÚ", "OÁ", "OÉ", "OÊ", _
                   "IÁ", "IÉ", "IÓ", "IÚ")

    For i = 2 To Len(Texto)
        Dim hv As String
        hv = Mid$(Texto, i - 1, 2)

        If EsMiembro(hv, hiatos) Then
            Call MF_ForzarDivisionSilabica(silabas, i)
        End If
    Next i

    ' ============================================================
    ' 5. RR y SS intervocálicas ? no dividir
    ' ============================================================
    For i = 2 To Len(Texto) - 1
        Dim seq As String
        seq = Mid$(Texto, i, 2)

        If seq = "RR" Or seq = "SS" Then
            Call MF_UnirConsonantesEnAtaque(silabas, i)
        End If
    Next i

End Sub

Public Sub MF_MarcarTonicaPortuguesBR( _
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

    vocalesTilde = "ÁÉÍÓÚÂÊÔ"

    ' --------------------------------------------------------
    ' 1. Silabear palabra
    ' --------------------------------------------------------
    'Set silabas = MF_SilabearCastellano(Texto)
    Set silabas = MF_Silabear(Texto, "pt-br")

    If silabas Is Nothing Then Exit Sub
    If silabas.Count = 0 Then Exit Sub

    ultima = Right$(Texto, 1)
    If Len(Texto) >= 2 Then ult2 = Right$(Texto, 2)

    ' --------------------------------------------------------
    ' 2. Si hay tilde ? esa sílaba
    ' --------------------------------------------------------
    For i = 1 To Len(Texto)
        If InStr(vocalesTilde, Mid$(Texto, i, 1)) > 0 Then
            idxTonica = MF_SilabaDeIndice(i, silabas)
            GoTo Marcar
        End If
    Next i

    ' --------------------------------------------------------
    ' 3. Reglas generales PT-BR
    ' --------------------------------------------------------

    ' 3.1. Oxítonas típicas
    If ultima = "L" Or ultima = "R" Or ultima = "Z" Or ultima = "X" Then
        idxTonica = silabas.Count
        GoTo Marcar
    End If

    If ultima = "I" Or ultima = "U" Then
        idxTonica = silabas.Count
        GoTo Marcar
    End If

    If ult2 = "IM" Or ult2 = "UM" Then
        idxTonica = silabas.Count
        GoTo Marcar
    End If

    ' 3.2. Paroxítonas (la mayoría)
    If silabas.Count = 1 Then
        idxTonica = 1
    Else
        idxTonica = silabas.Count - 1
    End If

Marcar:
    inicio = silabas(idxTonica)(1)
    fin = silabas(idxTonica)(2)

    For i = inicio To fin
        esTonica(i) = True
    Next i

End Sub

' ============================================================
'   ReglasPortugues (PT_BR)
'   Devuelve idFonema según la fonética del francés.
'   Si no aplica, devuelve 0 para que el motor siga probando.
' ============================================================
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

'============================================================================================

Public Function MF_NormalizarVocales_PT_BR(ByVal Texto As String) As String

    ' Nasales (idénticas a PT-EU)
    Texto = Replace(Texto, "Ã", "A~")
    Texto = Replace(Texto, "Õ", "O~")

    ' Cerradas ? se suavizan (PT-BR no las mantiene tan tensas)
    Texto = Replace(Texto, "Â", "A")
    Texto = Replace(Texto, "Ê", "E")
    Texto = Replace(Texto, "Ô", "O")

    ' Abiertas (agudas)
    Texto = Replace(Texto, "Á", "A´")
    Texto = Replace(Texto, "É", "E´")
    Texto = Replace(Texto, "Í", "I´")
    Texto = Replace(Texto, "Ó", "O´")
    Texto = Replace(Texto, "Ú", "U´")

    ' Graves (robustez)
    Texto = Replace(Texto, "À", "A")
    Texto = Replace(Texto, "È", "E")
    Texto = Replace(Texto, "Ì", "I")
    Texto = Replace(Texto, "Ò", "O")
    Texto = Replace(Texto, "Ù", "U")

    MF_NormalizarVocales_PT_BR = Texto

End Function

