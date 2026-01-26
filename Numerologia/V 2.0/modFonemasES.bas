Attribute VB_Name = "modFonemasES"
' ------------------------------------------------------
' Nombre:    modFonemasES
' Tipo:      Módulo
' Propósito:
' Autor:     asalv
' Fecha:     15/01/2026
' ------------------------------------------------------

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
    Dim C As String
    Dim sb As String

    texto = UCase$(texto)

    For i = 1 To Len(texto)
        C = Mid$(texto, i, 1)

        Select Case C
            Case "Á": sb = sb & "A"
            Case "É": sb = sb & "E"
            Case "Í": sb = sb & "I"
            Case "Ó": sb = sb & "O"
            Case "Ú": sb = sb & "U"
            Case Else
                sb = sb & C
        End Select
    Next i

    NormalizarTexto = sb
End Function

' ============================================================================
' TOKENIZADOR FONÉTICO
' ============================================================================

Private Function ExtraerFonema(ByVal texto As String, ByRef i As Integer) As String

    Dim C As String, sig As String, ant As String
    Dim c2 As String, c3 As String
    
    C = Mid$(texto, i, 1)
    
    If i < Len(texto) Then
        sig = Mid$(texto, i + 1, 1)
        c2 = C & sig
    Else
        sig = ""
        c2 = C
    End If
    
    If i < Len(texto) - 1 Then
        c3 = Mid$(texto, i, 3)
    Else
        c3 = c2
    End If
    
    If i > 1 Then ant = Mid$(texto, i - 1, 1) Else ant = ""

    ' ============================================================================
    ' 1) GÜE / GÜI --> GW + vocal
    ' ============================================================================
    If C = "G" And sig = "Ü" Then
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
    If C = "H" Then
        i = i + 1
        Exit Function
    End If
    
    ' ============================================================================
    ' 4) Ü aislada --> U
    ' ============================================================================
    If C = "Ü" Then
        ExtraerFonema = "U"
        i = i + 1
        Exit Function
    End If
    
    ' ============================================================================
    ' 5) QU --> K
    ' ============================================================================
    If C = "Q" Then
        ExtraerFonema = "K"
        If sig = "U" Then i = i + 2 Else i = i + 1
        Exit Function
    End If
    
    ' ============================================================================
    ' 6) U muda en QU / GU
    ' ============================================================================
    If C = "U" Then
        If ant = "Q" Or (ant = "G" And (sig = "E" Or sig = "I")) Then
            i = i + 1
            Exit Function
        End If
    End If
    
    ' ============================================================================
    ' 7) C --> K / Z
    ' ============================================================================
    If C = "C" Then
        If sig = "E" Or sig = "I" Then
            ExtraerFonema = "Z"
        Else
            ExtraerFonema = "K"
        End If
        i = i + 1
        Exit Function
    End If
    
    ' ============================================================================
    ' 8) G --> G / J
    ' ============================================================================
    If C = "G" Then
        If sig = "E" Or sig = "I" Then
            ExtraerFonema = "J"
        Else
            ExtraerFonema = "G"
        End If
        i = i + 1
        Exit Function
    End If
    
    ' ============================================================================
    ' 9) X --> KS
    ' ============================================================================
    If C = "X" Then
        ExtraerFonema = "KS"
        i = i + 1
        Exit Function
    End If
    
    ' ============================================================================
    ' 10) V --> B
    ' ============================================================================
    If C = "V" Then
        ExtraerFonema = "B"
        i = i + 1
        Exit Function
    End If
    
    ' ============================================================================
    ' Ñ --> NY
    ' ============================================================================
    If C = "Ñ" Then
        ExtraerFonema = "NY"
        i = i + 1
        Exit Function
    End If
    
    ' ============================================================================
    ' 11) R simple
    ' ============================================================================
    If C = "R" Then
        ExtraerFonema = "R"
        i = i + 1
        Exit Function
    End If
    
    ' ============================================================================
    ' 12) Y --> I / Y según contexto
    ' ============================================================================
    If C = "Y" Then
        ExtraerFonema = ProcesarY(ant, sig)
        i = i + 1
        Exit Function
    End If
    
    ' ============================================================================
    ' 13) W --> fonema según reglas
    ' ============================================================================
    If C = "W" Then
        ExtraerFonema = ProcesarW()
        i = i + 1
        Exit Function
    End If
    
    ' ============================================================================
    ' 14) Por defecto, devolver la letra tal cual
    ' ============================================================================
    ExtraerFonema = C
    i = i + 1
End Function


