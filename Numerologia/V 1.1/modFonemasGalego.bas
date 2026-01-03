Attribute VB_Name = "modFonemasGalego"

Option Compare Database
Option Explicit

' ============================================================================
' Módulo: modFonemasGalego
' Descripción: Tokenizador fonético para galego (versión optimizada)
' ============================================================================

Public Function ObtenerFonemasGalego(ByVal Nombre As String, _
                                     Optional ByVal UsarHmuda As Boolean = True, _
                                     Optional ByVal UsarUmuda As Boolean = True) As Collection
    Dim col As New Collection
    Dim txt As String
    Dim i As Long
    Dim f As String
    
    txt = NormalizarTexto(Nombre)
    i = 1
    
    Do While i <= Len(txt)
        f = ExtraerFonema(txt, i, UsarHmuda, UsarUmuda)
        If f <> "" Then col.Add f
    Loop
    
    Set ObtenerFonemasGalego = col
End Function

' ============================================================================
' NORMALIZACIÓN
'   - Mayúsculas
'   - Elimina tildes
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
            Case "Ú", "Ü": sb = sb & "U"
            Case Else
                sb = sb & c
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
    
'    ' 1) Espacios
'    If c = " " Then
'        i = i + 1
'        Exit Function
'    End If
    
    ' 2) Dígrafos
    Select Case c2
        Case "CH": ExtraerFonema = "CH": i = i + 2: Exit Function
        Case "LL": ExtraerFonema = "LL": i = i + 2: Exit Function
        Case "RR": ExtraerFonema = "RR": i = i + 2: Exit Function
    End Select
    
    ' 3) H muda
    If c = "H" And UsarHmuda Then
        i = i + 1
        Exit Function
    End If
    
    ' 4) QU ? K
    If c = "Q" Then
        ExtraerFonema = "K"
        If sig = "U" Then
            i = i + 2
        Else
            i = i + 1
        End If
        Exit Function
    End If
    
    ' 5) U muda en QU / GU
    If c = "U" And UsarUmuda Then
        If ant = "Q" Or (ant = "G" And (sig = "E" Or sig = "I")) Then
            i = i + 1
            Exit Function
        End If
    End If
    
    ' 6) C ? K o Z
    If c = "C" Then
        If sig = "E" Or sig = "I" Then
            ExtraerFonema = "Z"
        Else
            ExtraerFonema = "K"
        End If
        i = i + 1
        Exit Function
    End If
    
    ' 7) G ? J o G
    If c = "G" Then
        If sig = "E" Or sig = "I" Then
            ExtraerFonema = "J"   ' /x/
        Else
            ExtraerFonema = "G"
        End If
        i = i + 1
        Exit Function
    End If
    
    ' 8) J ? J (/x/)
    If c = "J" Then
        ExtraerFonema = "J"
        i = i + 1
        Exit Function
    End If
    
    ' 9) X ? SH (simplificación estándar)
    If c = "X" Then
        ExtraerFonema = "SH"
        i = i + 1
        Exit Function
    End If
    
    ' 10) V ? B
    If c = "V" Then
        ExtraerFonema = "B"
        i = i + 1
        Exit Function
    End If
    
    ' 11) Ñ ? Ñ (o NY si quisieras unificar)
    If c = "Ñ" Then
        ExtraerFonema = "Ñ"
        i = i + 1
        Exit Function
    End If
    
    ' 12) R simple
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

    ' 13) Por defecto, letra tal cual
    ExtraerFonema = c
    i = i + 1
End Function


