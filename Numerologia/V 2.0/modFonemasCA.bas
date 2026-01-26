Attribute VB_Name = "modFonemasCA"
' ------------------------------------------------------
' Nombre:    modFonemasCA
' Tipo:      Módulo
' Propósito:
' Autor:     asalv
' Fecha:     15/01/2026
' ------------------------------------------------------

Option Compare Database
Option Explicit

' ============================================================================
' Módulo: modFonemasCatalan (versión final corregida)
' ============================================================================

Public Function ObtenerFonemasCatalan(ByVal nombre As String, _
                                      Optional ByVal UsarHmuda As Boolean = True) As String

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
        
    ObtenerFonemasCatalan = Out
End Function

' ============================================================================
' NORMALIZACIÓN
'   - Mayúsculas
'   - Elimina tildes
'   - PRESERVA LA Ü (fundamental para GÜE/GÜI y QÜE/QÜI)
' ============================================================================

Private Function NormalizarTexto(ByVal txt As String) As String
    Dim i As Long, C As String, sb As String
    
    txt = UCase$(txt)
    
    For i = 1 To Len(txt)
        C = Mid$(txt, i, 1)
        
        Select Case C
            Case "Á", "À": sb = sb & "A"
            Case "É", "È": sb = sb & "E"
            Case "Í", "Ï": sb = sb & "I"
            Case "Ó", "Ò": sb = sb & "O"
            Case "Ú", "Ù": sb = sb & "U"
            Case "Ü": sb = sb & "Ü"   ' ¡se conserva!
            Case "·": sb = sb & "·"   ' Ela geminada
            Case Else: sb = sb & C
        End Select
    Next i
    
    NormalizarTexto = sb
End Function

' ============================================================================
' TOKENIZADOR FONÉTICO
' ============================================================================

Private Function ExtraerFonema(ByVal txt As String, ByRef i As Long) As String

    Dim C As String, sig As String, ant As String
    Dim c2 As String, c3 As String
    
    C = Mid$(txt, i, 1)
    
    If i < Len(txt) Then
        sig = Mid$(txt, i + 1, 1)
        c2 = C & sig
    Else
        sig = ""
        c2 = C
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
    If C = "G" And sig = "Ü" Then
        If i + 2 <= Len(txt) Then
            Dim sig2 As String
            sig2 = Mid$(txt, i + 2, 1)

            Select Case sig2
                Case "E", "I"
                    ExtraerFonema = "GW"
                    i = i + 2
                    Exit Function
            End Select
        End If
    End If

    ' ============================================================================
    ' 2) QÜE / QÜI ? KW + vocal
    ' ============================================================================
    If C = "Q" And sig = "Ü" Then
        If i + 2 <= Len(txt) Then
            'Dim sig2 As String
            sig2 = Mid$(txt, i + 2, 1)

            Select Case sig2
                Case "E", "I"
                    ExtraerFonema = "KW"
                    i = i + 2
                    Exit Function
            End Select
        End If
    End If

    ' ============================================================================
    ' 3) NY
    ' ============================================================================
    If c2 = "NY" Then
        ExtraerFonema = "NY"
        i = i + 2
        Exit Function
    End If
    
    ' ============================================================================
    ' 4) L·L ? LL
    ' ============================================================================
    If c3 = "L·L" Then
        ExtraerFonema = "LL"
        i = i + 3
        Exit Function
    End If
    
    ' ============================================================================
    ' 5) TX
    ' ============================================================================
    If c2 = "TX" Then
        ExtraerFonema = "TX"
        i = i + 2
        Exit Function
    End If
    
    ' ============================================================================
    ' 6) IG final ? TX
    ' ============================================================================
    If c2 = "IG" And (i + 1 = Len(txt)) Then
        ExtraerFonema = "TX"
        i = i + 2
        Exit Function
    End If
    
    ' ============================================================================
    ' 7) TG / DJ ? DJ
    ' ============================================================================
    If c2 = "TG" Or c2 = "DJ" Then
        ExtraerFonema = "DJ"
        i = i + 2
        Exit Function
    End If
    
    ' ============================================================================
    ' 8) X ? SH (versión simplificada)
    ' ============================================================================
    If C = "X" Then
        ExtraerFonema = "SH"
        i = i + 1
        Exit Function
    End If
    
    ' 9) G + E/I ? Y (fonema catalán /?/)
    If C = "G" And (sig = "E" Or sig = "I") Then
        ExtraerFonema = "Y"
        i = i + 1
        Exit Function
    End If
    
    ' 10) J ? Y (misma pronunciación /?/)
    If C = "J" Then
        ExtraerFonema = "Y"
        i = i + 1
        Exit Function
    End If
    
'    ' ============================================================================
'    ' 9) G + E/I --> J
'    ' ============================================================================
'    If c = "G" And (sig = "E" Or sig = "I") Then
'        ExtraerFonema = "J"
'        i = i + 1
'        Exit Function
'    End If
'
'    ' ============================================================================
'    ' 10) J --> J
'    ' ============================================================================
'    If c = "J" Then
'        ExtraerFonema = "J"
'        i = i + 1
'        Exit Function
'    End If
    
    
    
    ' ============================================================================
    ' 11) H muda (siempre muda en catalán)
    ' ============================================================================
    If C = "H" Then
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
    ' Ñ --> NY
    ' ============================================================================
    If C = "Ñ" Then
        ExtraerFonema = "NY"
        i = i + 1
        Exit Function
    End If

    ' ============================================================================
    ' 14) Por defecto, letra tal cual
    ' ============================================================================
    ExtraerFonema = C
    i = i + 1
End Function


