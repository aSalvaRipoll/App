Attribute VB_Name = "modMotor_Idioma_PT_BR"

Option Compare Database
Option Explicit

'==================
'== Portugués BR ==
'==================

Public Sub MF_MarcarTonica_PT_BR( _
        ByVal texto As String, _
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

    ' 1. Silabear palavra (motor com revisão)
    Set silabas = Silabear_PT_BR_ConRevision(texto)

    If silabas Is Nothing Then Exit Sub
    If silabas.Count = 0 Then Exit Sub

    ultima = Right$(texto, 1)
    If Len(texto) >= 2 Then ult2 = Right$(texto, 2)

    ' 2. Se houver acento --> essa sílaba
    For i = 1 To Len(texto)
        If InStr(vocalesTilde, Mid$(texto, i, 1)) > 0 Then
'            idxTonica = MF_SilabaDeIndice(i, Silabas)
            GoTo Marcar
        End If
    Next i

    ' 3. Regras gerais PT-BR

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

    ' 3.2. Paroxítonas (a maioria)
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

Public Function Silabear_PT_BR_ConRevision(ByVal texto As String) As Collection

    Dim col As Collection
    Dim resultado As New Collection
    Dim item As Variant
    Dim s As String
    Dim partes As Variant
    Dim p As Variant
    Dim inicio As Long, fin As Long
    Dim valido As Boolean
    Dim msg As String

    ' 1. Silabear automaticamente (motor puro PT-BR)
    Set col = Silabear_PT_BR(texto)

    ' 2. Converter para string com "-"
    For Each item In col
        s = s & Mid$(texto, item(0), item(1) - item(0) + 1) & "-"
    Next item
    If Len(s) > 0 Then s = Left$(s, Len(s) - 1)

    ' 3. Loop de validação com formulário
    Do
        valido = True
        msg = ""

        s = RevisarSilabas_EnFormulario(texto, s)

        If s = "" Then
            Set Silabear_PT_BR_ConRevision = col
            Exit Function
        End If

        If Left$(s, 1) = "-" Or Right$(s, 1) = "-" Then
            valido = False
            msg = "Não pode começar nem terminar com '-'."
        End If

        If InStr(s, "--") > 0 Then
            valido = False
            msg = "Não pode haver sílabas vazias ('--')."
        End If

        Dim reconstruido As String
        Dim textoSemEspacos As String

        reconstruido = Replace(s, "-", "")
        reconstruido = Replace(reconstruido, " ", "")
        textoSemEspacos = Replace(texto, " ", "")

        If UCase$(reconstruido) <> UCase$(textoSemEspacos) Then
            valido = False
            msg = "As sílabas não coincidem com o texto original."
        End If

        If Not valido Then
            MsgBox msg, vbExclamation, "Erro nas sílabas"
        End If

    Loop Until valido

    ' 4. Reconstruir coleção final
    partes = Split(s, "-")
    inicio = 1

    For Each p In partes
        fin = inicio + Len(p) - 1
        resultado.Add Array(inicio, fin)
        inicio = fin + 1
    Next p

    Set Silabear_PT_BR_ConRevision = resultado

End Function

Public Function Silabear_PT_BR(ByVal texto As String) As Collection

    Dim col As New Collection
    Dim i As Long, ini As Long
    Dim c1 As String, c2 As String, c3 As String
    Dim par As String

    texto = Trim$(texto)
    If Len(texto) = 0 Then
        Set Silabear_PT_BR = col
        Exit Function
    End If

    ini = 1

    For i = 2 To Len(texto)

        c1 = Mid$(texto, i - 1, 1)
        c2 = Mid$(texto, i, 1)
        par = c1 & c2

        ' ---------------------------------------------------------
        ' 0. Espaços --> separam palavras
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
        ' 1. VV --> hiato se vocal fraca tônica (Í, Ú)
        ' ---------------------------------------------------------
        If EsVocal_PT(c1) And EsVocal_PT(c2) Then
            If c1 = "Í" Or c1 = "Ú" Or c2 = "Í" Or c2 = "Ú" Then
                col.Add Array(ini, i - 1)
                ini = i
                GoTo siguiente
            End If
        End If

        ' ---------------------------------------------------------
        ' 2. VCV --> V | CV
        ' ---------------------------------------------------------
        If EsVocal_PT(c1) And EsConsonant_PT(c2) Then
            If i < Len(texto) Then
                c3 = Mid$(texto, i + 1, 1)
                If EsVocal_PT(c3) Then
                    col.Add Array(ini, i - 1)
                    ini = i
                    GoTo siguiente
                End If
            End If
        End If

        ' ---------------------------------------------------------
        ' 3. CCV --> C | CV
        ' ---------------------------------------------------------
        If EsConsonant_PT(c1) And EsConsonant_PT(c2) Then
            If i < Len(texto) Then
                c3 = Mid$(texto, i + 1, 1)
                If EsVocal_PT(c3) Then
                    If Not EsGrupInseparable_PT(par) Then
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

    Set Silabear_PT_BR = col

End Function

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

    ' R inicial --> aspirado (lo mapeamos a H suave: 38)
    If g = "R" And ant = "" Then ReglasPortugues_PT_BR = 38: Exit Function

    ' SS --> /s/
    If g = "SS" Then ReglasPortugues_PT_BR = 34: Exit Function

    ' S entre vocales --> /z/
    If g = "S" And (ant Like "[AEIOUÃÕÁÉÍÓÚÂÊÔ]" And sig Like "[AEIOUÃÕÁÉÍÓÚÂÊÔ]") Then
        ReglasPortugues_PT_BR = 35: Exit Function
    End If

    ' S final --> /s/ (no /?/)
    If g = "S" And sig = "" Then ReglasPortugues_PT_BR = 34: Exit Function

    ' X --> /?/ estándar
    If g = "X" Then ReglasPortugues_PT_BR = 36: Exit Function

    ' J --> /?/
    If g = "J" Then ReglasPortugues_PT_BR = 37: Exit Function

    ' G + E/I --> /?/
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

Public Function MF_NormalizarVocales_PT_BR(ByVal texto As String) As String

    ' Nasales (idénticas a PT-EU)
    texto = Replace(texto, "Ã", "A~")
    texto = Replace(texto, "Õ", "O~")

    ' Cerradas --> se suavizan (PT-BR no las mantiene tan tensas)
    texto = Replace(texto, "Â", "A")
    texto = Replace(texto, "Ê", "E")
    texto = Replace(texto, "Ô", "O")

    ' Abiertas (agudas)
    texto = Replace(texto, "Á", "A´")
    texto = Replace(texto, "É", "E´")
    texto = Replace(texto, "Í", "I´")
    texto = Replace(texto, "Ó", "O´")
    texto = Replace(texto, "Ú", "U´")

    ' Graves (robustez)
    texto = Replace(texto, "À", "A")
    texto = Replace(texto, "È", "E")
    texto = Replace(texto, "Ì", "I")
    texto = Replace(texto, "Ò", "O")
    texto = Replace(texto, "Ù", "U")

    MF_NormalizarVocales_PT_BR = texto

End Function

