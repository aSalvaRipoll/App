Attribute VB_Name = "modMotor_Idioma_GL"

Option Compare Database
Option Explicit

'=================
'==   Galego    ==
'=================
Public Sub MF_SilabearAjustesGL( _
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

Public Sub MF_MarcarTonicaGallego( _
        ByVal Texto As String, _
        ByRef esTonica() As Boolean)

    Dim silabas As Collection
    Dim idxTonica As Long
    Dim inicio As Long, fin As Long
    Dim i As Long
    Dim vocalesTilde As String

    vocalesTilde = "ÁÉÍÓÚ"

    ' --------------------------------------------------------
    ' 1. Silabear palabra
    ' --------------------------------------------------------
    'Set silabas = MF_SilabearCastellano(Texto)
    Set silabas = MF_Silabear(Texto, "gl")


    If silabas Is Nothing Then Exit Sub
    If silabas.Count = 0 Then Exit Sub

    ' --------------------------------------------------------
    ' 2. Buscar tilde (igual que en castellano)
    ' --------------------------------------------------------
    For i = 1 To Len(Texto)
        If InStr(vocalesTilde, Mid$(Texto, i, 1)) > 0 Then
            idxTonica = MF_SilabaDeIndice(i, silabas)
            Exit For
        End If
    Next i

    ' --------------------------------------------------------
    ' 3. Si no hay tilde ? reglas generales del gallego
    ' --------------------------------------------------------
    If idxTonica = 0 Then
        Dim ultima As String
        ultima = Right$(Texto, 1)

        If InStr("AEIOUÁÉÍÓÚNS", ultima) > 0 Then
            ' Palabras rematadas en vogal, -n, -s ? penúltima
            If silabas.Count = 1 Then
                idxTonica = 1
            Else
                idxTonica = silabas.Count - 1
            End If
        Else
            ' Resto ? última
            idxTonica = silabas.Count
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

'============================================================================================

Public Function MF_NormalizarVocales_GL(ByVal Texto As String) As String

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

    MF_NormalizarVocales_GL = Texto

End Function

