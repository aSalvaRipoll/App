Attribute VB_Name = "modFonemasEU"
' ------------------------------------------------------
' Nombre:    modFonemasEU
' Tipo:      Módulo
' Propósito:
' Autor:     asalv
' Fecha:     15/01/2026
' ------------------------------------------------------

Option Compare Database
Option Explicit

' ============================================================================
' Módulo: modFonemasEuskera (versión final afinada)
' ============================================================================

Public Function ObtenerFonemasEuskera(ByVal nombre As String) As String
    Dim txt As String
    Dim i As Long
    Dim f As String
    Dim Out As String
    
    txt = NormalizarTexto(nombre)
    i = 1
    Out = ""
    
    Do While i <= Len(txt)
        f = ExtraerFonema(txt, i)
        If f <> "" Then Out = Out & f
    Loop
    
    ObtenerFonemasEuskera = Out
End Function

' ============================================================================
' NORMALIZACIÓN
'   - Mayúsculas
'   - Elimina tildes
'   - Añade Ï para robustez
' ============================================================================

Private Function NormalizarTexto(ByVal txt As String) As String
    Dim i As Long, C As String, sb As String
    
    txt = UCase$(txt)
    
    For i = 1 To Len(txt)
        C = Mid$(txt, i, 1)
        
        Select Case C
            Case "Á", "À": sb = sb & "A"
            Case "É", "È": sb = sb & "E"
            Case "Í", "Ì", "Ï": sb = sb & "I"
            Case "Ó", "Ò": sb = sb & "O"
            Case "Ú", "Ù": sb = sb & "U"
            Case Else: sb = sb & C
        End Select
    Next i
    
    NormalizarTexto = sb
End Function

' ============================================================================
' TOKENIZADOR FONÉTICO
' ============================================================================

Private Function ExtraerFonema(ByVal txt As String, _
                               ByRef i As Long) As String

    Dim C As String, sig As String, ant As String
    Dim c2 As String
    
    C = Mid$(txt, i, 1)
    
    If i < Len(txt) Then
        sig = Mid$(txt, i + 1, 1)
        c2 = C & sig
    Else
        sig = ""
        c2 = C
    End If
    
    If i > 1 Then ant = Mid$(txt, i - 1, 1) Else ant = ""

    ' ============================================================================
    ' 1) Dígrafos propios del euskera
    ' ============================================================================
    Select Case c2
        Case "TX": ExtraerFonema = "TX": i = i + 2: Exit Function
        Case "TS": ExtraerFonema = "TS": i = i + 2: Exit Function
        Case "TZ": ExtraerFonema = "TZ": i = i + 2: Exit Function
        Case "RR": ExtraerFonema = "RR": i = i + 2: Exit Function
    End Select

    ' ============================================================================
    ' 2) NH --> NY (préstamos galaico-portugueses)
    ' ============================================================================
    If c2 = "NH" Then
        ExtraerFonema = "NY"
        i = i + 2
        Exit Function
    End If

    ' ============================================================================
    ' 3) Ñ --> NY
    ' ============================================================================
    If C = "Ñ" Then
        ExtraerFonema = "NY"
        i = i + 1
        Exit Function
    End If

    ' ============================================================================
    ' 4) X --> SH
    ' ============================================================================
    If C = "X" Then
        ExtraerFonema = "SH"
        i = i + 1
        Exit Function
    End If

    ' ============================================================================
    ' 5) H --> H (siempre suena)
    ' ============================================================================
    If C = "H" Then
        ExtraerFonema = "H"
        i = i + 1
        Exit Function
    End If

    ' ============================================================================
    ' 6) K --> K (siempre /k/)
    ' ============================================================================
    If C = "K" Then
        ExtraerFonema = "K"
        i = i + 1
        Exit Function
    End If

    ' ============================================================================
    ' 7) S y Z se mantienen tal cual
    ' ============================================================================
    If C = "S" Or C = "Z" Then
        ExtraerFonema = C
        i = i + 1
        Exit Function
    End If

    ' ============================================================================
    ' 8) R inicial ? RR
    ' ============================================================================
    If C = "R" And i = 1 Then
        ExtraerFonema = "RR"
        i = i + 1
        Exit Function
    End If

    ' ============================================================================
    ' 9) R simple
    ' ============================================================================
    If C = "R" Then
        ExtraerFonema = "R"
        i = i + 1
        Exit Function
    End If


' ============================================================================
' --- Q (préstamos) ? K ---
' ============================================================================
If C = "Q" Then
    ExtraerFonema = "K"
    ' Si viene QU, saltamos la U
    If sig = "U" Then
        i = i + 2
    Else
        i = i + 1
    End If
    Exit Function
End If

' ============================================================================
' --- U muda en GU + E/I (préstamos) ---
' ============================================================================
If C = "U" Then
    If ant = "G" And (sig = "E" Or sig = "I") Then
        ' U muda ? se omite
        i = i + 1
        Exit Function
    End If
End If

    ' ============================================================================
    ' --- C (préstamos) ? K / S ---
    '   CE / CI ? S
    '   CA / CO / CU ? K
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
    ' --- G (préstamos) ? G / J ---
    '   GE / GI ? J
    '   GA / GO / GU ? G
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
    ' 10) Y --> I / Y según contexto
    ' ============================================================================
    If C = "Y" Then
        ExtraerFonema = ProcesarY(ant, sig)
        i = i + 1
        Exit Function
    End If

    ' ============================================================================
    ' 11) W --> fonema según reglas
    ' ============================================================================
    If C = "W" Then
        ExtraerFonema = ProcesarW()
        i = i + 1
        Exit Function
    End If

    ' ============================================================================
    ' 12) Por defecto
    ' ============================================================================
    ExtraerFonema = C
    i = i + 1
End Function


