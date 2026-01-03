Attribute VB_Name = "modFonemasCastellano"

Option Compare Database
Option Explicit

' ============================================================================
' Módulo: modFonemasCastellano
' Descripción: Tokenizador fonético para castellano (versión corregida)
' ============================================================================

' ============================================================================
' FUNCIÓN PRINCIPAL
' ============================================================================

Public Function ObtenerFonemasCastellano(ByVal nombre As String, _
                                         Optional ByVal UsarHmuda As Boolean = True, _
                                         Optional ByVal UsarUmuda As Boolean = True) As Collection
    Dim col As New Collection
    Dim texto As String
    Dim i As Long
    Dim fonema As String
    
    texto = NormalizarTextoCastellano(nombre)
    i = 1
    
    Do While i <= Len(texto)
        fonema = ExtraerFonemaCastellano(texto, i, UsarHmuda, UsarUmuda)
        If fonema <> "" Then col.Add fonema
    Loop
    
    Set ObtenerFonemasCastellano = col
End Function

' ============================================================================
' NORMALIZACIÓN ESPECÍFICA DEL CASTELLANO
'   - Mayúsculas
'   - Elimina tildes
'   - NO toca la Ü (para no romper reglas fonéticas)
' ============================================================================

Private Function NormalizarTextoCastellano(ByVal texto As String) As String
    texto = UCase$(texto)
    
    ' Vocales acentuadas ? vocal simple
    texto = Replace(texto, "Á", "A")
    texto = Replace(texto, "É", "E")
    texto = Replace(texto, "Í", "I")
    texto = Replace(texto, "Ó", "O")
    texto = Replace(texto, "Ú", "U")
    
    ' OJO: NO convertir Ü aquí. La tratamos en el tokenizador.
    ' texto = Replace(texto, "Ü", "U")
    
    NormalizarTextoCastellano = texto
End Function

' ============================================================================
' TOKENIZADOR DE FONEMAS
' ============================================================================

Private Function ExtraerFonemaCastellano(ByVal texto As String, _
                                         ByRef i As Long, _
                                         ByVal UsarHmuda As Boolean, _
                                         ByVal UsarUmuda As Boolean) As String
    Dim c As String, c2 As String, c3 As String
    Dim sig As String, ant As String
    
    c = Mid$(texto, i, 1)
    c2 = Mid$(texto, i, 2)
    c3 = Mid$(texto, i, 3)
    sig = Mid$(texto, i + 1, 1)
    ant = IIf(i > 1, Mid$(texto, i - 1, 1), "")
    
    ' 1) Espacios
    If c = " " Then
        i = i + 1
        Exit Function
    End If
    
    ' 2) Dígrafos reales
    If c2 = "CH" Then
        ExtraerFonemaCastellano = "CH"
        i = i + 2
        Exit Function
    End If
    
    If c2 = "LL" Then
        ExtraerFonemaCastellano = "LL"
        i = i + 2
        Exit Function
    End If
    
    If c2 = "RR" Then
        ExtraerFonemaCastellano = "RR"
        i = i + 2
        Exit Function
    End If
    
    ' 3) H muda
    If c = "H" And UsarHmuda Then
        i = i + 1
        Exit Function
    End If
    
    ' 4) Ü ? U sonora (siempre fonema independiente)
    If c = "Ü" Then
        ExtraerFonemaCastellano = "U"
        i = i + 1
        Exit Function
    End If
    
    ' 5) GÜE / GÜI: aquí la G es oclusiva /G/, la Ü ya se tratará en el siguiente paso
    If c3 = "GÜE" Or c3 = "GÜI" Then
        ExtraerFonemaCastellano = "G"
        i = i + 1
        Exit Function
    End If
    
    ' 6) QU ? K (U no suena nunca)
    If c = "Q" Then
        ExtraerFonemaCastellano = "K"
        If sig = "U" Then
            i = i + 2
        Else
            i = i + 1
        End If
        Exit Function
    End If
    
    ' 7) U muda en QU / GU (solo para U, no para Ü)
    If c = "U" And UsarUmuda Then
        If ant = "Q" Or (ant = "G" And (sig = "E" Or sig = "I")) Then
            i = i + 1
            Exit Function
        End If
    End If
    
    ' 8) C ? K o Z
    If c = "C" Then
        If sig = "E" Or sig = "I" Then
            ExtraerFonemaCastellano = "Z"
        Else
            ExtraerFonemaCastellano = "K"
        End If
        i = i + 1
        Exit Function
    End If
    
    ' 9) G ? G o J
    If c = "G" Then
        If sig = "E" Or sig = "I" Then
            ExtraerFonemaCastellano = "J"
        Else
            ExtraerFonemaCastellano = "G"
        End If
        i = i + 1
        Exit Function
    End If
    
    ' 10) X ? KS
    If c = "X" Then
        ExtraerFonemaCastellano = "KS"
        i = i + 1
        Exit Function
    End If
    
    ' 11) V ? B
    If c = "V" Then
        ExtraerFonemaCastellano = "B"
        i = i + 1
        Exit Function
    End If
    
    ' 12) R simple
    If c = "R" Then
        ExtraerFonemaCastellano = "R"
        i = i + 1
        Exit Function
    End If
    
    ' 13) Por defecto, devolver la letra tal cual
    ExtraerFonemaCastellano = c
    i = i + 1
End Function


