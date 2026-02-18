Attribute VB_Name = "modMotor_Idioma_GL"

Option Compare Database
Option Explicit

'=================
'==   Galego    ==
'=================

Public Sub MF_MarcarTonica_GL( _
        ByVal texto As String, _
        ByRef esTonica() As Boolean)

    Dim silabas As Collection
    Dim idxTonica As Long
    Dim inicio As Long, Fin As Long
    Dim i As Long
    Dim vocalesTilde As String

    vocalesTilde = "ÁÉÍÓÚ"

    ' 1. Silabear palabra (motor con revisión)
    Set silabas = Silabear_GL_ConRevision(texto)

    If silabas Is Nothing Then Exit Sub
    If silabas.Count = 0 Then Exit Sub

    ' 2. Buscar tilde
    For i = 1 To Len(texto)
        If InStr(vocalesTilde, Mid$(texto, i, 1)) > 0 Then
'            idxTonica = MF_SilabaDeIndice(i, Silabas)
            Exit For
        End If
    Next i

    ' 3. Regras xerais do galego
    If idxTonica = 0 Then
        Dim ultima As String
        ultima = Right$(texto, 1)

        If InStr("AEIOUÁÉÍÓÚNS", ultima) > 0 Then
            ' Paroxítona
            If silabas.Count = 1 Then
                idxTonica = 1
            Else
                idxTonica = silabas.Count - 1
            End If
        Else
            ' Oxítona
            idxTonica = silabas.Count
        End If
    End If

    ' 4. Marcar índices tónicos
    inicio = silabas(idxTonica)(1)
    Fin = silabas(idxTonica)(2)

    For i = inicio To Fin
        esTonica(i) = True
    Next i

End Sub

Public Function Silabear_GL_ConRevision(ByVal texto As String) As Collection

    Dim col As Collection
    Dim resultado As New Collection
    Dim item As Variant
    Dim s As String
    Dim partes As Variant
    Dim p As Variant
    Dim inicio As Long, Fin As Long
    Dim valido As Boolean
    Dim msg As String

    ' 1. Silabear automáticamente (motor puro gallego)
    Set col = Silabear_GL(texto)

    ' 2. Convertir a string con "-"
    For Each item In col
        s = s & Mid$(texto, item(0), item(1) - item(0) + 1) & "-"
    Next item
    If Len(s) > 0 Then s = Left$(s, Len(s) - 1)

    ' 3. Bucle de validación con formulario
    Do
        valido = True
        msg = ""

        s = RevisarSilabas_EnFormulario(texto, s)

        If s = "" Then
            Set Silabear_GL_ConRevision = col
            Exit Function
        End If

        If Left$(s, 1) = "-" Or Right$(s, 1) = "-" Then
            valido = False
            msg = "Non pode comezar nin rematar con '-'."
        End If

        If InStr(s, "--") > 0 Then
            valido = False
            msg = "Non pode haber sílabas baleiras ('--')."
        End If

        Dim reconstruido As String
        Dim textoSenEspazos As String

        reconstruido = Replace(s, "-", "")
        reconstruido = Replace(reconstruido, " ", "")
        textoSenEspazos = Replace(texto, " ", "")

        If UCase$(reconstruido) <> UCase$(textoSenEspazos) Then
            valido = False
            msg = "As sílabas non coinciden co texto orixinal."
        End If

        If Not valido Then
            MsgBox msg, vbExclamation, "Erro nas sílabas"
        End If

    Loop Until valido

    ' 4. Reconstruír colección final
    partes = Split(s, "-")
    inicio = 1

    For Each p In partes
        Fin = inicio + Len(p) - 1
        resultado.Add Array(inicio, Fin)
        inicio = Fin + 1
    Next p

    Set Silabear_GL_ConRevision = resultado

End Function

Public Function Silabear_GL(ByVal texto As String) As Collection

    Dim col As New Collection
    Dim i As Long, ini As Long
    Dim c1 As String, c2 As String, c3 As String
    Dim par As String

    texto = Trim$(texto)
    If Len(texto) = 0 Then
        Set Silabear_GL = col
        Exit Function
    End If

    ini = 1

    For i = 2 To Len(texto)

        c1 = Mid$(texto, i - 1, 1)
        c2 = Mid$(texto, i, 1)
        par = c1 & c2

        ' ---------------------------------------------------------
        ' 0. Espacios --> separan palabras
        ' ---------------------------------------------------------
        If c1 = " " Then
            If i - 2 >= ini Then col.Add Array(ini, i - 2)
            ini = i
            GoTo siguiente
        End If

        If c2 = " " Then
            col.Add Array(ini, i - 1)
            ini = i + 1
            GoTo siguiente
        End If

        ' ---------------------------------------------------------
        ' 1. Diptongos/hiatos (igual que castellano)
        ' ---------------------------------------------------------
        If EsVocal_GL(c1) And EsVocal_GL(c2) Then
            If EsHiato_GL(c1, c2) Then
                col.Add Array(ini, i - 1)
                ini = i
                GoTo siguiente
            Else
                GoTo siguiente
            End If
        End If

        ' ---------------------------------------------------------
        ' 2. VCV --> V | CV
        ' ---------------------------------------------------------
        If EsVocal_GL(c1) And EsConsonant_GL(c2) Then
            If i < Len(texto) Then
                c3 = Mid$(texto, i + 1, 1)
                If EsVocal_GL(c3) Then
                    col.Add Array(ini, i - 1)
                    ini = i
                    GoTo siguiente
                End If
            End If
        End If

        ' ---------------------------------------------------------
        ' 3. CCV --> C | CV
        ' ---------------------------------------------------------
        If EsConsonant_GL(c1) And EsConsonant_GL(c2) Then
            If i < Len(texto) Then
                c3 = Mid$(texto, i + 1, 1)
                If EsVocal_GL(c3) Then
                    If Not EsGrupInseparable_GL(par) Then
                        col.Add Array(ini, i - 1)
                        ini = i
                        GoTo siguiente
                    End If
                End If
            End If
        End If

siguiente:
    Next i

    ' Última sílaba
    If ini <= Len(texto) Then
        col.Add Array(ini, Len(texto))
    End If

    Set Silabear_GL = col

End Function

Public Function EsVocal_GL(c As String) As Boolean
    EsVocal_GL = InStr("AEIOUÁÉÍÓÚaeiouáéíóú", c) > 0
End Function

Public Function EsConsonant_GL(c As String) As Boolean
    EsConsonant_GL = Not EsVocal_GL(c) And c <> " "
End Function

Public Function EsGrupInseparable_GL(par As String) As Boolean
    Select Case UCase$(par)
        Case "BR", "BL", "CR", "CL", "DR", "TR", "PR", "PL", "GR", "GL", "FR", "FL"
            EsGrupInseparable_GL = True
        Case Else
            EsGrupInseparable_GL = False
    End Select
End Function

Public Function EsHiato_GL(v1 As String, v2 As String) As Boolean
    ' Vocal fuerte + vocal fuerte --> hiato
    If InStr("AÁEÉOÓaáeéoó", v1) > 0 And InStr("AÁEÉOÓaáeéoó", v2) > 0 Then
        EsHiato_GL = True
        Exit Function
    End If

    ' Vocal débil tónica --> hiato
    If v1 = "Í" Or v1 = "Ú" Or v2 = "Í" Or v2 = "Ú" Then
        EsHiato_GL = True
        Exit Function
    End If

    EsHiato_GL = False
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

    ' GÜE / GÜI --> /gw/ --> id 57
    If g = "GÜE" Or g = "GÜI" Then
        ReglasGalego = 57
        Exit Function
    End If

    ' GUE / GUI --> /g/ (U muda) --> id 31
    If g = "GUE" Or g = "GUI" Then
        ReglasGalego = 31
        Exit Function
    End If

    ' QUE / QUI --> /k/ --> id 30
    If g = "QUE" Or g = "QUI" Then
        ReglasGalego = 30
        Exit Function
    End If


    ' ============================================================
    '   DÍGRAFOS Y CASOS ESPECIALES
    ' ============================================================

    ' CH --> /t?/ --> id 50
    If g = "CH" Then
        ReglasGalego = 50
        Exit Function
    End If

    ' X --> /?/ --> id 36
    If g = "X" Then
        ReglasGalego = 36
        Exit Function
    End If

    ' J --> /?/ --> id 37
    If g = "J" Then
        ReglasGalego = 37
        Exit Function
    End If

    ' G + E/I --> /?/ --> id 37
    If g = "G" And (sig = "E" Or sig = "I") Then
        ReglasGalego = 37
        Exit Function
    End If

    ' LL --> /?/ --> id 44
    If g = "LL" Then
        ReglasGalego = 44
        Exit Function
    End If

    ' Ñ --> /?/ --> id 41
    If g = "Ñ" Then
        ReglasGalego = 41
        Exit Function
    End If

    ' RR --> /r/ múltiple --> id 46
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

    ' S / Z / C+E/I --> /s/
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

    ' H --> aspiración suave --> id 38
    If g = "H" Then ReglasGalego = 38: Exit Function


    ' ============================================================
    '   SI NO APLICA, DEVOLVER 0
    ' ============================================================
    ReglasGalego = 0

End Function

'============================================================================================

Public Function MF_NormalizarVocales_GL(ByVal texto As String) As String

    ' A
    texto = Replace(texto, "Á", "A")
    texto = Replace(texto, "À", "A")

    ' E
    texto = Replace(texto, "É", "E")
    texto = Replace(texto, "È", "E")

    ' I
    texto = Replace(texto, "Í", "I")
    texto = Replace(texto, "Ì", "I")

    ' O
    texto = Replace(texto, "Ó", "O")
    texto = Replace(texto, "Ò", "O")

    ' U
    texto = Replace(texto, "Ú", "U")
    texto = Replace(texto, "Ù", "U")

    MF_NormalizarVocales_GL = texto

End Function
