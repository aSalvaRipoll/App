Attribute VB_Name = "modFonemasES"

Option Compare Database
Option Explicit

' ============================================================================
' Módulo: modFonemasCastellano (versión final corregida)
' ============================================================================

' ============================================================================
' FUNCIÓN PRINCIPAL
' ============================================================================

Public Function ObtenerFonemasCastellano(ByVal nombre As String, _
                                         Optional ByVal UsarHmuda As Boolean = True, _
                                         Optional ByVal UsarUmuda As Boolean = True) As String

    Dim texto As String
    Dim i As Integer
    Dim f As String
    Dim Out As String
    
    texto = NormalizarTexto(nombre)
    i = 1
    Out = ""
    
    Do While i <= Len(texto)
        f = ExtraerFonema(texto, i)
        If f <> "" Then Out = Out & f
    Loop
    
    ObtenerFonemasCastellano = Out
End Function

' ============================================================================
' NORMALIZACIÓN
'   - Mayúsculas
'   - Elimina tildes
'   - NO toca la Ü (fundamental para GÜE/GÜI)
' ============================================================================

Private Function NormalizarTexto(ByVal texto As String) As String
    Dim i As Long
    Dim c As String
    Dim sb As String

    texto = UCase$(texto)

    For i = 1 To Len(texto)
        c = Mid$(texto, i, 1)

        Select Case c
            Case "Á": sb = sb & "A"
            Case "É": sb = sb & "E"
            Case "Í": sb = sb & "I"
            Case "Ó": sb = sb & "O"
            Case "Ú": sb = sb & "U"
            Case Else
                sb = sb & c
        End Select
    Next i

    NormalizarTexto = sb
End Function

' ============================================================================
' TOKENIZADOR FONÉTICO
' ============================================================================

Private Function ExtraerFonema(ByVal texto As String, ByRef i As Integer) As String

    Dim c As String, sig As String, ant As String
    Dim c2 As String, c3 As String
    
    c = Mid$(texto, i, 1)
    
    If i < Len(texto) Then
        sig = Mid$(texto, i + 1, 1)
        c2 = c & sig
    Else
        sig = ""
        c2 = c
    End If
    
    If i < Len(texto) - 1 Then
        c3 = Mid$(texto, i, 3)
    Else
        c3 = c2
    End If
    
    If i > 1 Then ant = Mid$(texto, i - 1, 1) Else ant = ""

    ' ============================================================================
    ' 1) GÜE / GÜI ? GW + vocal
    ' ============================================================================
    If c = "G" And sig = "Ü" Then
        If i + 2 <= Len(texto) Then
            Dim sig2 As String
            sig2 = Mid$(texto, i + 2, 1)

            Select Case sig2
                Case "E", "I"
                    ExtraerFonema = "GW"
                    i = i + 2   ' saltamos G + Ü
                    Exit Function
            End Select
        End If
    End If

    ' ============================================================================
    ' 2) Dígrafos CH, LL, RR
    ' ============================================================================
    Select Case c2
        Case "CH": ExtraerFonema = "CH": i = i + 2: Exit Function
        Case "LL": ExtraerFonema = "LL": i = i + 2: Exit Function
        Case "RR": ExtraerFonema = "RR": i = i + 2: Exit Function
    End Select
    
    ' ============================================================================
    ' 3) H muda (siempre muda en castellano)
    ' ============================================================================
    If c = "H" Then
        i = i + 1
        Exit Function
    End If
    
    ' ============================================================================
    ' 4) Ü aislada ? U
    ' ============================================================================
    If c = "Ü" Then
        ExtraerFonema = "U"
        i = i + 1
        Exit Function
    End If
    
    ' ============================================================================
    ' 5) QU ? K
    ' ============================================================================
    If c = "Q" Then
        ExtraerFonema = "K"
        If sig = "U" Then i = i + 2 Else i = i + 1
        Exit Function
    End If
    
    ' ============================================================================
    ' 6) U muda en QU / GU
    ' ============================================================================
    If c = "U" Then
        If ant = "Q" Or (ant = "G" And (sig = "E" Or sig = "I")) Then
            i = i + 1
            Exit Function
        End If
    End If
    
    ' ============================================================================
    ' 7) C ? K / Z
    ' ============================================================================
    If c = "C" Then
        If sig = "E" Or sig = "I" Then
            ExtraerFonema = "Z"
        Else
            ExtraerFonema = "K"
        End If
        i = i + 1
        Exit Function
    End If
    
    ' ============================================================================
    ' 8) G ? G / J
    ' ============================================================================
    If c = "G" Then
        If sig = "E" Or sig = "I" Then
            ExtraerFonema = "J"
        Else
            ExtraerFonema = "G"
        End If
        i = i + 1
        Exit Function
    End If
    
    ' ============================================================================
    ' 9) X ? KS
    ' ============================================================================
    If c = "X" Then
        ExtraerFonema = "KS"
        i = i + 1
        Exit Function
    End If
    
    ' ============================================================================
    ' 10) V ? B
    ' ============================================================================
    If c = "V" Then
        ExtraerFonema = "B"
        i = i + 1
        Exit Function
    End If
    
    ' ============================================================================
    ' Ñ ? NY
    ' ============================================================================
    If c = "Ñ" Then
        ExtraerFonema = "NY"
        i = i + 1
        Exit Function
    End If
    
    ' ============================================================================
    ' 11) R simple
    ' ============================================================================
    If c = "R" Then
        ExtraerFonema = "R"
        i = i + 1
        Exit Function
    End If
    
    ' ============================================================================
    ' 12) Y ? I / Y según contexto
    ' ============================================================================
    If c = "Y" Then
        ExtraerFonema = ProcesarY(ant, sig)
        i = i + 1
        Exit Function
    End If
    
    ' ============================================================================
    ' 13) W ? fonema según reglas
    ' ============================================================================
    If c = "W" Then
        ExtraerFonema = ProcesarW()
        i = i + 1
        Exit Function
    End If
    
    ' ============================================================================
    ' 14) Por defecto, devolver la letra tal cual
    ' ============================================================================
    ExtraerFonema = c
    i = i + 1
End Function


