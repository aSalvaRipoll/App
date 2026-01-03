Attribute VB_Name = "modFonemasGA"

Option Compare Database
Option Explicit

' ============================================================================
' Módulo: modFonemasGalego (versión premium afinada)
' ============================================================================

Public Function ObtenerFonemasGalego(ByVal nombre As String, _
                                     Optional ByVal UsarHmuda As Boolean = True, _
                                     Optional ByVal UsarUmuda As Boolean = True) As String

    Dim txt As String
    Dim i As Long
    Dim f As String
    Dim Out As String
    
    txt = NormalizarTexto(nombre)
    i = 1
    Out = ""
    
    Do While i <= Len(txt)
        f = ExtraerFonema(txt, i, UsarHmuda, UsarUmuda)
        If f <> "" Then Out = Out & f
    Loop
    
    ObtenerFonemasGalego = Out
End Function

' ============================================================================
' NORMALIZACIÓN — PRESERVA LA Ü
' ============================================================================

Private Function NormalizarTexto(ByVal txt As String) As String
    Dim i As Long, c As String, sb As String
    
    txt = UCase$(txt)
    
    For i = 1 To Len(txt)
        c = Mid$(txt, i, 1)
        
        Select Case c
            Case "Á", "À": sb = sb & "A"
            Case "É", "È": sb = sb & "E"
            Case "Í", "Ï": sb = sb & "I"
            Case "Ó", "Ò": sb = sb & "O"
            Case "Ú": sb = sb & "U"
            Case "Ü": sb = sb & "Ü"   ' ¡se conserva!
            Case Else: sb = sb & c
        End Select
    Next i
    
    NormalizarTexto = sb
End Function

' ============================================================================
' TOKENIZADOR
' ============================================================================

Private Function ExtraerFonema(ByVal txt As String, _
                               ByRef i As Long, _
                               ByVal UsarHmuda As Boolean, _
                               ByVal UsarUmuda As Boolean) As String

    Dim c As String, sig As String, ant As String
    Dim c2 As String, c3 As String
    
    c = Mid$(txt, i, 1)
    
    If i < Len(txt) Then
        sig = Mid$(txt, i + 1, 1)
        c2 = c & sig
    Else
        sig = ""
        c2 = c
    End If
    
    If i < Len(txt) - 1 Then
        c3 = Mid$(txt, i, 3)
    Else
        c3 = c2
    End If
    
    If i > 1 Then ant = Mid$(txt, i - 1, 1) Else ant = ""

    ' ============================================================================
    ' 1) GÜE / GÜI ? GW + vocal
    ' ============================================================================
    If c = "G" And sig = "Ü" Then
        If i + 2 <= Len(txt) Then
            Dim sig2 As String
            sig2 = Mid$(txt, i + 2, 1)
            If sig2 = "E" Or sig2 = "I" Then
                ExtraerFonema = "GW"
                i = i + 2
                Exit Function
            End If
        End If
    End If

    ' ============================================================================
    ' 2) QÜE / QÜI ? KW + vocal
    ' ============================================================================
    If c = "Q" And sig = "Ü" Then
        If i + 2 <= Len(txt) Then
            Dim sig2 As String
            sig2 = Mid$(txt, i + 2, 1)
            If sig2 = "E" Or sig2 = "I" Then
                ExtraerFonema = "KW"
                i = i + 2
                Exit Function
            End If
        End If
    End If

    ' ============================================================================
    ' 3) Dígrafos CH, LL, RR
    ' ============================================================================
    Select Case c2
        Case "CH": ExtraerFonema = "CH": i = i + 2: Exit Function
        Case "LL": ExtraerFonema = "LL": i = i + 2: Exit Function
        Case "RR": ExtraerFonema = "RR": i = i + 2: Exit Function
    End Select

    ' ============================================================================
    ' 4) NH ? Ñ (coherencia con sistemas románicos)
    ' ============================================================================
    If c2 = "NH" Then
        ExtraerFonema = "Ñ"
        i = i + 2
        Exit Function
    End If

    ' ============================================================================
    ' 5) GH + E/I ? G (dialectal tradicional)
    ' ============================================================================
    If c2 = "GH" Then
        If sig = "H" And (Mid$(txt, i + 2, 1) = "E" Or Mid$(txt, i + 2, 1) = "I") Then
            ExtraerFonema = "G"
            i = i + 2
            Exit Function
        End If
    End If

    ' ============================================================================
    ' 6) H muda
    ' ============================================================================
    If c = "H" Then
        i = i + 1
        Exit Function
    End If

    ' ============================================================================
    ' 7) QU ? K
    ' ============================================================================
    If c = "Q" Then
        ExtraerFonema = "K"
        If sig = "U" Then i = i + 2 Else i = i + 1
        Exit Function
    End If

    ' ============================================================================
    ' 8) U muda en QU / GU
    ' ============================================================================
    If c = "U" Then
        If ant = "Q" Or (ant = "G" And (sig = "E" Or sig = "I")) Then
            i = i + 1
            Exit Function
        End If
    End If

    ' ============================================================================
    ' 9) C ? K / Z
    ' ============================================================================
    If c = "C" Then
        If sig = "E" Or sig = "I" Then ExtraerFonema = "Z" Else ExtraerFonema = "K"
        i = i + 1
        Exit Function
    End If

    ' ============================================================================
    ' 10) G ? G / J
    ' ============================================================================
    If c = "G" Then
        If sig = "E" Or sig = "I" Then ExtraerFonema = "J" Else ExtraerFonema = "G"
        i = i + 1
        Exit Function
    End If

    ' ============================================================================
    ' 11) J ? J
    ' ============================================================================
    If c = "J" Then
        ExtraerFonema = "J"
        i = i + 1
        Exit Function
    End If

    ' ============================================================================
    ' 12) X galega completa
    ' ============================================================================
    If c = "X" Then
        
        ' X inicial ? SH
        If i = 1 Then
            ExtraerFonema = "SH"
            i = i + 1
            Exit Function
        End If
        
        ' X + vocal ? GZ
        If sig Like "[AEIOU]" Then
            ExtraerFonema = "GZ"
            i = i + 1
            Exit Function
        End If
        
        ' X final ? KS
        If sig = "" Then
            ExtraerFonema = "KS"
            i = i + 1
            Exit Function
        End If
        
        ' Resto ? KS
        ExtraerFonema = "KS"
        i = i + 1
        Exit Function
    End If

    ' ============================================================================
    ' 13) V ? B
    ' ============================================================================
    If c = "V" Then
        ExtraerFonema = "B"
        i = i + 1
        Exit Function
    End If

    ' ============================================================================
    ' 14) Ñ, NH --> NY
    ' ============================================================================
    ' --- NH ? NY (galaico-portugués) ---
    If c2 = "NH" Then
        ExtraerFonema = "NY"
        i = i + 2
        Exit Function
    End If
    
    ' --- Ñ ? NY ---
    If c = "Ñ" Then
        ExtraerFonema = "NY"
        i = i + 1
        Exit Function
    End If

'    If c = "Ñ" Then
'        ExtraerFonema = "NY"
'        i = i + 1
'        Exit Function
'    End If

    ' ============================================================================
    ' 15) R inicial --> RR
    ' ============================================================================
    If c = "R" And i = 1 Then
        ExtraerFonema = "RR"
        i = i + 1
        Exit Function
    End If

    ' ============================================================================
    ' 16) R simple
    ' ============================================================================
    If c = "R" Then
        ExtraerFonema = "R"
        i = i + 1
        Exit Function
    End If

    ' ============================================================================
    ' 17) Y --> I / Y según contexto
    ' ============================================================================
    If c = "Y" Then
        ExtraerFonema = ProcesarY(ant, sig)
        i = i + 1
        Exit Function
    End If

    ' ============================================================================
    ' 18) W ? fonema según reglas
    ' ============================================================================
    If c = "W" Then
        ExtraerFonema = ProcesarW()
        i = i + 1
        Exit Function
    End If

    ' ============================================================================
    ' 19) Por defecto
    ' ============================================================================
    ExtraerFonema = c
    i = i + 1
End Function

'Option Compare Database
'Option Explicit
'
'' ============================================================================
'' Módulo: modFonemasGalego (versión corregida)
'' ============================================================================
'
'Public Function ObtenerFonemasGalego(ByVal nombre As String, _
'                                     Optional ByVal UsarHmuda As Boolean = True, _
'                                     Optional ByVal UsarUmuda As Boolean = True) As String
'
'    Dim txt As String
'    Dim i As Long
'    Dim f As String
'    Dim Out As String
'
'    txt = NormalizarTexto(nombre)
'    i = 1
'    Out = ""
'
'    Do While i <= Len(txt)
'        f = ExtraerFonema(txt, i)
'        If f <> "" Then Out = Out & f
'    Loop
'
'    ObtenerFonemasGalego = Out
'End Function
'
'' ============================================================================
'' NORMALIZACIÓN — PRESERVA LA Ü
'' ============================================================================
'
'Private Function NormalizarTexto(ByVal txt As String) As String
'    Dim i As Long, c As String, sb As String
'
'    txt = UCase$(txt)
'
'    For i = 1 To Len(txt)
'        c = Mid$(txt, i, 1)
'
'        Select Case c
'            Case "Á", "À": sb = sb & "A"
'            Case "É", "È": sb = sb & "E"
'            Case "Í", "Ï": sb = sb & "I"
'            Case "Ó", "Ò": sb = sb & "O"
'            Case "Ú": sb = sb & "U"
''            Case "Ü": sb = sb & "Ü"   ' ¡se conserva!
'            Case Else: sb = sb & c
'        End Select
'    Next i
'
'    NormalizarTexto = sb
'End Function
'
'' ============================================================================
'' TOKENIZADOR
'' ============================================================================
'
'Private Function ExtraerFonema(ByVal txt As String, ByRef i As Long) As String
'
'    Dim c As String, sig As String, ant As String
'    Dim c2 As String, c3 As String
'
'    c = Mid$(txt, i, 1)
'
'    If i < Len(txt) Then
'        sig = Mid$(txt, i + 1, 1)
'        c2 = c & sig
'    Else
'        sig = ""
'        c2 = c
'    End If
'
'    If i < Len(txt) - 1 Then
'        c3 = Mid$(txt, i, 3)
'    Else
'        c3 = c2
'    End If
'
'    If i > 1 Then ant = Mid$(txt, i - 1, 1) Else ant = ""
'
'    ' ============================================================================
'    ' 1) GÜE / GÜI ? GW + vocal
'    ' ============================================================================
'    If c = "G" And sig = "Ü" Then
'        If i + 2 <= Len(txt) Then
'            Dim sig2 As String
'            sig2 = Mid$(txt, i + 2, 1)
'
'            Select Case sig2
'                Case "E", "I"
'                    ExtraerFonema = "GW"
'                    i = i + 2
'                    Exit Function
'            End Select
'        End If
'    End If
'
'    ' ============================================================================
'    ' 2) QÜE / QÜI ? KW + vocal
'    ' ============================================================================
'    If c = "Q" And sig = "Ü" Then
'        If i + 2 <= Len(txt) Then
'            Dim sig2 As String
'            sig2 = Mid$(txt, i + 2, 1)
'
'            Select Case sig2
'                Case "E", "I"
'                    ExtraerFonema = "KW"
'                    i = i + 2
'                    Exit Function
'            End Select
'        End If
'    End If
'
'    ' ============================================================================
'    ' 3) Dígrafos
'    ' ============================================================================
'    Select Case c2
'        Case "CH": ExtraerFonema = "CH": i = i + 2: Exit Function
'        Case "LL": ExtraerFonema = "LL": i = i + 2: Exit Function
'        Case "RR": ExtraerFonema = "RR": i = i + 2: Exit Function
'    End Select
'
'    ' ============================================================================
'    ' 4) H muda (siempre muda en galego)
'    ' ============================================================================
'    If c = "H" Then
'        i = i + 1
'        Exit Function
'    End If
'
'    ' ============================================================================
'    ' 5) QU ? K
'    ' ============================================================================
'    If c = "Q" Then
'        ExtraerFonema = "K"
'        If sig = "U" Then i = i + 2 Else i = i + 1
'        Exit Function
'    End If
'
'    ' ============================================================================
'    ' 6) U muda en QU / GU
'    ' ============================================================================
'    If c = "U" Then
'        If ant = "Q" Or (ant = "G" And (sig = "E" Or sig = "I")) Then
'            i = i + 1
'            Exit Function
'        End If
'    End If
'
'    ' ============================================================================
'    ' 7) C ? K / Z
'    ' ============================================================================
'    If c = "C" Then
'        If sig = "E" Or sig = "I" Then ExtraerFonema = "Z" Else ExtraerFonema = "K"
'        i = i + 1
'        Exit Function
'    End If
'
'    ' ============================================================================
'    ' 8) G ? G / J
'    ' ============================================================================
'    If c = "G" Then
'        If sig = "E" Or sig = "I" Then ExtraerFonema = "J" Else ExtraerFonema = "G"
'        i = i + 1
'        Exit Function
'    End If
'
'    ' ============================================================================
'    ' 9) J ? J
'    ' ============================================================================
'    If c = "J" Then
'        ExtraerFonema = "J"
'        i = i + 1
'        Exit Function
'    End If
'
'    ' ============================================================================
'    ' 10) X ? SH (simplificación)
'    ' ============================================================================
'    If c = "X" Then
'        ExtraerFonema = "SH"
'        i = i + 1
'        Exit Function
'    End If
'
'    ' ============================================================================
'    ' 11) V ? B
'    ' ============================================================================
'    If c = "V" Then
'        ExtraerFonema = "B"
'        i = i + 1
'        Exit Function
'    End If
'
'    ' ============================================================================
'    ' 12) Ñ ? Ñ
'    ' ============================================================================
'    If c = "Ñ" Then
'        ExtraerFonema = "Ñ"
'        i = i + 1
'        Exit Function
'    End If
'
'    ' ============================================================================
'    ' 13) R simple
'    ' ============================================================================
'    If c = "R" Then
'        ExtraerFonema = "R"
'        i = i + 1
'        Exit Function
'    End If
'
'    ' ============================================================================
'    ' 14) Y ? I / Y según contexto
'    ' ============================================================================
'    If c = "Y" Then
'        ExtraerFonema = ProcesarY(ant, sig)
'        i = i + 1
'        Exit Function
'    End If
'
'    ' ============================================================================
'    ' 15) W ? fonema según reglas
'    ' ============================================================================
'    If c = "W" Then
'        ExtraerFonema = ProcesarW()
'        i = i + 1
'        Exit Function
'    End If
'
'    ' ============================================================================
'    ' 16) Por defecto
'    ' ============================================================================
'    ExtraerFonema = c
'    i = i + 1
'End Function


