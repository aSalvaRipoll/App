Attribute VB_Name = "modFonemasCatalan"

Option Compare Database
Option Explicit

' ============================================================================
' Módulo: modFonemasCatalan
' Descripción: Tokenizador fonético para catalán (versión optimizada)
' ============================================================================

Public Function ObtenerFonemasCatalan(ByVal Nombre As String, _
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
    
    Set ObtenerFonemasCatalan = col
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
            Case "Í", "Ï": sb = sb & "I"
            Case "Ó", "Ò": sb = sb & "O"
            Case "Ú", "Ü": sb = sb & "U"
            Case "·": sb = sb & "·"   ' Ela geminada
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
    
'    ' 1) Espacios
'    If c = " " Then i = i + 1: Exit Function
   
'   ' 1) Espacios ? se añaden como fonema literal
'    If c = " " Then
'        ExtraerFonemaCastellano = " "
'        i = i + 1
'        Exit Function
'    End If

    ' 2) NY
    If c2 = "NY" Then ExtraerFonema = "NY": i = i + 2: Exit Function
    
    ' 3) L·L ? LL
    If c3 = "L·L" Then ExtraerFonema = "LL": i = i + 3: Exit Function
    
    ' 4) TX ? TX
    If c2 = "TX" Then ExtraerFonema = "TX": i = i + 2: Exit Function
    
    ' 5) IG final ? TX
    If c2 = "IG" And (i + 1 = Len(txt)) Then
        ExtraerFonema = "TX"
        i = i + 2
        Exit Function
    End If
    
    ' 6) TG / DJ ? DJ
    If c2 = "TG" Or c2 = "DJ" Then
        ExtraerFonema = "DJ"
        i = i + 2
        Exit Function
    End If
    
    ' 7) X según contexto
    If c = "X" Then
        ' Inicial ? SH
'        If i = 1 Then
            ExtraerFonema = "SH"
'        ' EX + vocal ? GZ
'        ElseIf ant = "E" And (sig Like "[AEIOU]") Then
'            ExtraerFonema = "GZ"
'        ' Intervocálica ? GZ
'        ElseIf (ant Like "[AEIOU]") And (sig Like "[AEIOU]") Then
'            ExtraerFonema = "GZ"
'        ' Final ? KS
'        ElseIf sig = "" Then
'            ExtraerFonema = "KS"
'        Else
'            ExtraerFonema = "KS"
'        End If
        
        i = i + 1
        Exit Function
    End If
    
    ' 8) G + E/I ? J
    If c = "G" And (sig = "E" Or sig = "I") Then
        ExtraerFonema = "J"
        i = i + 1
        Exit Function
    End If
    
    ' 9) J ? J
    If c = "J" Then
        ExtraerFonema = "J"
        i = i + 1
        Exit Function
    End If
    
    ' 10) H muda
    If c = "H" And UsarHmuda Then
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

    ' 11) Por defecto, letra tal cual
    ExtraerFonema = c
    i = i + 1
End Function




