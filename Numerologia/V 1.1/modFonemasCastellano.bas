Attribute VB_Name = "modFonemasCastellano"

Option Compare Database
Option Explicit

' ============================================================================
' Módulo: modFonemasCastellano (versión optimizada)
' ============================================================================

' ============================================================================
' FUNCIÓN PRINCIPAL
' ============================================================================

Public Function ObtenerFonemasCastellano(ByVal Nombre As String, _
                                         Optional ByVal UsarHmuda As Boolean = True, _
                                         Optional ByVal UsarUmuda As Boolean = True) As Collection
    Dim col As New Collection
    Dim texto As String
    Dim i As Long
    Dim fonema As String
    
    texto = NormalizarTexto(Nombre)
    i = 1
    
    Do While i <= Len(texto)
        fonema = ExtraerFonema(texto, i, UsarHmuda, UsarUmuda)
        If fonema <> "" Then col.Add fonema
    Loop
    
    Set ObtenerFonemasCastellano = col
End Function

' ============================================================================
' NORMALIZACIÓN OPTIMIZADA
'   - Mayúsculas
'   - Elimina tildes
'   - NO toca la Ü
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
' TOKENIZADOR OPTIMIZADO
' ============================================================================

Private Function ExtraerFonema(ByVal texto As String, _
                                         ByRef i As Long, _
                                         ByVal UsarHmuda As Boolean, _
                                         ByVal UsarUmuda As Boolean) As String
    Dim c As String, sig As String, ant As String
    Dim c2 As String, c3 As String
    
    c = Mid$(texto, i, 1)
    
    ' Evitar llamadas repetidas
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
    
'    ' 1) Espacios
'    If c = " " Then
'        i = i + 1
'        Exit Function
'    End If
    
'    ' 1) Espacios ? se añaden como fonema literal
'    If c = " " Then
'        ExtraerFonema = " "
'        i = i + 1
'        Exit Function
'    End If
    
    ' 2) Dígrafos frecuentes
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
    
    ' 4) Ü ? U sonora
    If c = "Ü" Then
        ExtraerFonema = "U"
        i = i + 1
        Exit Function
    End If
    
    ' 5) GÜE / GÜI ? G
    If c3 = "GÜE" Or c3 = "GÜI" Then
        ExtraerFonema = "G"
        i = i + 1
        Exit Function
    End If
    
    ' 6) QU ? K
    If c = "Q" Then
        ExtraerFonema = "K"
        If sig = "U" Then i = i + 2 Else i = i + 1
        Exit Function
    End If
    
    ' 7) U muda en QU / GU
    If c = "U" And UsarUmuda Then
        If ant = "Q" Or (ant = "G" And (sig = "E" Or sig = "I")) Then
            i = i + 1
            Exit Function
        End If
    End If
    
    ' 8) C ? K o Z
    If c = "C" Then
        If sig = "E" Or sig = "I" Then
            ExtraerFonema = "Z"
        Else
            ExtraerFonema = "K"
        End If
        i = i + 1
        Exit Function
    End If
    
    ' 9) G ? G o J
    If c = "G" Then
        If sig = "E" Or sig = "I" Then
            ExtraerFonema = "J"
        Else
            ExtraerFonema = "G"
        End If
        i = i + 1
        Exit Function
    End If
    
    ' 10) X ? KS
    If c = "X" Then
        ExtraerFonema = "KS"
        i = i + 1
        Exit Function
    End If
    
    ' 11) V ? B
    If c = "V" Then
        ExtraerFonema = "B"
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
    
    ' 13) Por defecto, devolver la letra tal cual
    ExtraerFonema = c
    i = i + 1
End Function


