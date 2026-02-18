Attribute VB_Name = "modMotor_Idioma_CA_IB"

Option Compare Database
Option Explicit

'================
'== Mallorquín ==
'================
Public Sub MF_SilabearAjustesCA_IB( _
        ByVal Texto As String, _
        ByRef silabas As Collection _
    )

    Dim i As Long

    ' ============================================================
    ' 1. ATAQUES CONSONÁNTICOS PROPIOS DEL MALLORQUÍN
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
    ' 2. LL y RR nunca se separan
    ' ============================================================
    For i = 2 To Len(Texto) - 1
        If Mid$(Texto, i, 2) = "LL" Or Mid$(Texto, i, 2) = "RR" Then
            Call MF_UnirConsonantesEnAtaque(silabas, i)
        End If
    Next i

    ' ============================================================
    ' 3. HIATOS MALLORQUINES (ea, eo, oa)
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
    ' 4. Diptongos mallorquines (muy estables)
    ' ============================================================
    Dim dipt As Variant
    dipt = Array( _
        "IA", "IE", "IO", "IU", _
        "UA", "UE", "UI", "UO", _
        "AI", "EI", "OI", "AU", "EU", "OU" _
    )

    For i = 2 To Len(Texto)
        Dim dv As String
        dv = Mid$(Texto, i - 1, 2)

        If EsMiembro(dv, dipt) Then
            Call MF_UnirVocalesEnDiptongo(silabas, i)
        End If
    Next i

End Sub

Public Sub MF_MarcarTonicaMallorquin( _
        ByVal Texto As String, _
        ByRef esTonica() As Boolean _
    )

    Dim silabas As Collection
    Dim idxTonica As Long
    Dim inicio As Long, fin As Long
    Dim i As Long
    Dim vocalesTilde As String

    vocalesTilde = "ÁÉÍÓÚÀÈÌÒÙ"

    ' --------------------------------------------------------
    ' 1. Silabear palabra
    ' --------------------------------------------------------
    'Set silabas = MF_SilabearCastellano(Texto)
    Set silabas = MF_Silabear(Texto, "ca-ib")

    If silabas Is Nothing Then Exit Sub
    If silabas.Count = 0 Then Exit Sub

    ' --------------------------------------------------------
    ' 2. Buscar tilde (aguda, grave o esdrújula)
    ' --------------------------------------------------------
    For i = 1 To Len(Texto)
        If InStr(vocalesTilde, Mid$(Texto, i, 1)) > 0 Then
            idxTonica = MF_SilabaDeIndice(i, silabas)
            Exit For
        End If
    Next i

    ' --------------------------------------------------------
    ' 3. Si no hay tilde ? acento penúltimo (mallorquín)
    ' --------------------------------------------------------
    If idxTonica = 0 Then
        If silabas.Count = 1 Then
            idxTonica = 1
        Else
            idxTonica = silabas.Count - 1
        End If
    End If

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

'============================================================================================

Public Function MF_NormalizarVocales_CA_IB(ByVal Texto As String) As String
    MF_NormalizarVocales_CA_IB = MF_NormalizarVocales_CA(Texto)
End Function

