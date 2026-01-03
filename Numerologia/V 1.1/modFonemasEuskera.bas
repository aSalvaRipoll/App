Attribute VB_Name = "modFonemasEuskera"

Option Compare Database
Option Explicit

' ============================================================================
' Módulo: modFonemasEuskera
' Descripción: Tokenizador fonético para euskera (versión optimizada)
' ============================================================================

Public Function ObtenerFonemasEuskera(ByVal Nombre As String) As Collection
    Dim col As New Collection
    Dim txt As String
    Dim i As Long
    Dim f As String
    
    txt = NormalizarTexto(Nombre)
    i = 1
    
    Do While i <= Len(txt)
        f = ExtraerFonema(txt, i)
        If f <> "" Then col.Add f
    Loop
    
    Set ObtenerFonemasEuskera = col
End Function

' ============================================================================
' NORMALIZACIÓN
' ============================================================================

Private Function NormalizarTexto(ByVal txt As String) As String
    Dim i As Long, c As String, sb As String
    
    txt = UCase$(txt)
    
    For i = 1 To Len(txt)
        c = Mid$(txt, i, 1)
        
        Select Case c
            Case "Á", "À": sb = sb & "A"
            Case "É", "È": sb = sb & "E"
            Case "Í": sb = sb & "I"
            Case "Ó", "Ò": sb = sb & "O"
            Case "Ú": sb = sb & "U"
            Case Else: sb = sb & c
        End Select
    Next i
    
    NormalizarTexto = sb
End Function

' ============================================================================
' TOKENIZADOR
' ============================================================================

Private Function ExtraerFonema(ByVal txt As String, _
                                      ByRef i As Long) As String
    Dim c As String, sig As String, ant As String
    Dim c2 As String
    
    c = Mid$(txt, i, 1)
    
    If i < Len(txt) Then
        sig = Mid$(txt, i + 1, 1)
        c2 = c & sig
    Else
        sig = ""
        c2 = c
    End If
    
    If i > 1 Then ant = Mid$(txt, i - 1, 1) Else ant = ""
    
'    ' 1) Espacios
'    If c = " " Then i = i + 1: Exit Function
    
    ' 2) Dígrafos propios del euskera
    Select Case c2
        Case "TX": ExtraerFonema = "TX": i = i + 2: Exit Function
        Case "TS": ExtraerFonema = "TS": i = i + 2: Exit Function
        Case "TZ": ExtraerFonema = "TZ": i = i + 2: Exit Function
        Case "RR": ExtraerFonema = "RR": i = i + 2: Exit Function
    End Select
    
    ' 3) X ? SH
    If c = "X" Then
        ExtraerFonema = "SH"
        i = i + 1
        Exit Function
    End If
    
    ' 4) H ? H (siempre suena)
    If c = "H" Then
        ExtraerFonema = "H"
        i = i + 1
        Exit Function
    End If
    
    ' 5) K ? K (siempre /k/)
    If c = "K" Then
        ExtraerFonema = "K"
        i = i + 1
        Exit Function
    End If
    
    ' 6) S y Z se mantienen tal cual
    If c = "S" Or c = "Z" Then
        ExtraerFonema = c
        i = i + 1
        Exit Function
    End If
    
    ' 7) R simple
    If c = "R" Then
        ExtraerFonema = "R"
        i = i + 1
        Exit Function
    End If
    
    ' Tratamiento de la Y
    If c = "Y" Then
        ExtraerFonema = ProcesarY(ant, c, sig)
        i = i + 1
        Exit Function
    End If

    ' Tratamiento de la W
    If c = "W" Then
        ExtraerFonema = ProcesarW(c)
        i = i + 1
        Exit Function
    End If

    ' 8) Por defecto, devolver la letra tal cual
    ExtraerFonema = c
    i = i + 1
End Function


