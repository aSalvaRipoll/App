Attribute VB_Name = "modAuxCalculos"

Option Compare Database
Option Explicit

' ============================================================================
'  FONEMAS COMPUESTOS (14 sistemas multilingües)
' ============================================================================

Private Function EsFonemaCompuesto(f As String) As Boolean
    Select Case f
        Case "CH", "LL", "RR", "NY", _
             "TX", "TS", "TZ", "DJ", _
             "SH", "KS", "GZ", _
             "GW", "KW"
             
            EsFonemaCompuesto = True
        Case Else
            EsFonemaCompuesto = False
    End Select
End Function



' ============================================================================
'  TABLA PITAGÓRICA TRADICIONAL (Sistema = 1)
' ============================================================================
' ASR: Agregada distinción de sistema anglosajón antiguo (1) y moderno (3)
Public Function ConvertirLetraANumero(letra As String, modoFon As ModoFonetico, modoCalc As ModoCalculo) As Byte
    
    ' En fonético no se procesan letras
    If modoFon = mfFonetico Then
        ConvertirLetraANumero = 0
        Exit Function
    End If

    letra = UCase(letra)

    ' ============================================================================
    ' EXCEPCIONES SEGÚN SISTEMA (Q, K, y futuras)
    ' ============================================================================
    Select Case letra

        Case "Q"
            If modoCalc = mcClasico Then
                ConvertirLetraANumero = 8
            Else
                ConvertirLetraANumero = 3
            End If
            Exit Function

        Case "K"
            If modoCalc = mcClasico Then
                ConvertirLetraANumero = 2
            Else
                ConvertirLetraANumero = 3
            End If
            Exit Function

        'Case "Ñ"
        '    If modoCalc = Clasico Then
        '        ConvertirLetraANumero = 5
        '    Else
        '        ConvertirLetraANumero = 6
        '    End If
        '    Exit Function

        'Case "Ç"
        '    ConvertirLetraANumero = 3
        '    Exit Function

        Case Else ' TABLA GENERAL (clásica por ahora)
            ConvertirLetraANumero = TablaClasica(letra)
            
    End Select

End Function

Private Function TablaClasica(letra As String) As Byte

    Select Case letra
        Case "A", "J", "S": TablaClasica = 1
        Case "B", "T": TablaClasica = 2
        Case "C", "L", "U": TablaClasica = 3
        Case "D", "M", "V": TablaClasica = 4
        Case "E", "N", "W": TablaClasica = 5
        Case "F", "O", "X": TablaClasica = 6
        Case "G", "P", "Y": TablaClasica = 7
        Case "H", "Z": TablaClasica = 8
        Case "I", "R": TablaClasica = 9
        Case Else: TablaClasica = 0
    End Select

End Function

Private Function TablaModerna(letra As String) As Byte

    ' Por ahora es igual, pero ya está preparada para ampliaciones
    Select Case letra
        Case "A", "J", "S": TablaModerna = 1
        Case "B", "T": TablaModerna = 2
        Case "C", "L", "U": TablaModerna = 3
        Case "D", "M", "V": TablaModerna = 4
        Case "E", "N", "W": TablaModerna = 5
        Case "F", "O", "X": TablaModerna = 6
        Case "G", "P", "Y": TablaModerna = 7
        Case "H", "Z": TablaModerna = 8
        Case "I", "R": TablaModerna = 9
        Case Else: TablaModerna = 0
    End Select

End Function

'Public Function ConvertirLetraANumero(letra As String, modoFon As ModoFonetico, modoCalc As ModoCalculo) As Byte
'
'    ' Sistema 1 = Fonético --> esta función no aplica
'    If modoFon = Fonetico Then
'        ConvertirLetraANumero = 0
'        Exit Function
'    End If
'
'    letra = UCase(letra)
'
'    ' Normalización universal para sistemas 1 y 3
'    ' ASR Ya no es necesario normalizar aquí, se hace en la llamada
''    letra = NormalizarLetraTradicional(letra)
'
'    ' ============================================================================
'    '  EXCEPCIONES SEGÚN SISTEMA (Q y K)
'    ' ============================================================================
'
'    If letra = "Q" Then
'        If modoCalc = Clasico Then        ' Tradicional anglosajón antiguo
'            ConvertirLetraANumero = 8
'        ElseIf modoCalc = Moderno Then    ' Tradicional anglosajón moderno
'            ConvertirLetraANumero = 3
'        End If
'        Exit Function
'    End If
'
'    If letra = "K" Then
'        If modoCalc = Clasico Then        ' Tradicional anglosajón antiguo
'            ConvertirLetraANumero = 2
'        ElseIf modoCalc = Moderno Then    ' Tradicional anglosajón moderno
'            ConvertirLetraANumero = 3
'        End If
'        Exit Function
'    End If
'
'    ' ============================================================================
'    '  TABLA GENERAL (válida para sistemas 1)
'    ' ============================================================================
'
'    Select Case letra
'        Case "A", "J", "S": ConvertirLetraANumero = 1
'        Case "B", "T": ConvertirLetraANumero = 2
'        Case "C", "L", "U": ConvertirLetraANumero = 3
'        'Case "C", "Ç", "L", "U": ConvertirLetraANumero = 3
'        Case "D", "M", "V": ConvertirLetraANumero = 4
'        Case "E", "N", "W": ConvertirLetraANumero = 5
'        'Case "E", "N", "Ñ", "W": ConvertirLetraANumero = 5
'        Case "F", "O", "X": ConvertirLetraANumero = 6
'        Case "G", "P", "Y": ConvertirLetraANumero = 7
'        Case "H", "Z": ConvertirLetraANumero = 8
'        Case "I", "R": ConvertirLetraANumero = 9
'        Case Else
'            ConvertirLetraANumero = 0
'    End Select
'
'End Function

' ============================================================================
'  NORMALIZADOR PARA USAR SISTEMA TRADICIONAL (Sistema = 1)
' ============================================================================

Public Function NormalizarLetraTradicional(ByVal letra As String) As String
    
    ' Convertir a mayúscula para unificar tratamiento
    letra = UCase$(letra)

    ' Normalización de vocales acentuadas y variantes
    Select Case letra
        Case "Á", "À", "Ä"
            letra = "A"

        Case "É", "È", "Ë"
            letra = "E"

        Case "Í", "Ì", "Ï"
            letra = "I"

        Case "Ó", "Ò", "Ö"
            letra = "O"

        Case "Ú", "Ù", "Ü"
            letra = "U"

        ' Consonantes especiales
        Case "Ç"
            letra = "C"

        Case "Ñ"
            letra = "N"

        Case Else
            ' No se modifica
    End Select

    NormalizarLetraTradicional = letra
End Function

'Public Function NormalizarLetraTradicional(letra As String) As String
'    letra = UCase$(letra)
'
'    Select Case letra
'        Case "Á", "À", "Ä": NormalizarLetraTradicional = "A"
'        Case "É", "È", "Ë": NormalizarLetraTradicional = "E"
'        Case "Í", "Ì", "Ï": NormalizarLetraTradicional = "I"
'        Case "Ó", "Ò", "Ö": NormalizarLetraTradicional = "O"
'        Case "Ú", "Ù", "Ü": NormalizarLetraTradicional = "U"
'        Case "Ç": NormalizarLetraTradicional = "C"
'        Case "Ñ": NormalizarLetraTradicional = "N"
'        Case Else
'            NormalizarLetraTradicional = letra
'    End Select
'End Function


' ============================================================================
'  TABLA FONÉTICA UNIVERSAL (Sistema = 2)
' ============================================================================

Public Function ConvertirFonemaANumero(f As String) As Integer

    ' Asegurar mayúsculas
    f = UCase$(f)

    Select Case f
        ' ============================
        ' FONEMAS COMPUESTOS
        ' ============================
        Case "NY": ConvertirFonemaANumero = 7
        Case "CH": ConvertirFonemaANumero = 6
        Case "LL": ConvertirFonemaANumero = 3
        Case "RR": ConvertirFonemaANumero = 9
        Case "TX": ConvertirFonemaANumero = 6
        Case "TS": ConvertirFonemaANumero = 2
        Case "TZ": ConvertirFonemaANumero = 8
        Case "DJ": ConvertirFonemaANumero = 1
        Case "SH": ConvertirFonemaANumero = 1
        Case "KS": ConvertirFonemaANumero = 6
        Case "GZ": ConvertirFonemaANumero = 8
        Case "GW": ConvertirFonemaANumero = 7
        Case "KW": ConvertirFonemaANumero = 3
        
        ' ============================
        ' VOCALES
        ' ============================
        Case "A": ConvertirFonemaANumero = 1
        Case "E": ConvertirFonemaANumero = 5
        Case "I": ConvertirFonemaANumero = 9
        Case "O": ConvertirFonemaANumero = 6
        Case "U": ConvertirFonemaANumero = 3

        ' ============================
        ' CONSONANTES SIMPLES
        ' ============================
        Case "B": ConvertirFonemaANumero = 2
        Case "C", "K", "Q": ConvertirFonemaANumero = 3
        Case "D": ConvertirFonemaANumero = 4
        Case "F": ConvertirFonemaANumero = 6
        Case "G": ConvertirFonemaANumero = 7
        Case "H": ConvertirFonemaANumero = 8
        Case "J": ConvertirFonemaANumero = 1
        Case "L": ConvertirFonemaANumero = 3
        Case "M": ConvertirFonemaANumero = 4
        Case "N": ConvertirFonemaANumero = 5
        Case "P": ConvertirFonemaANumero = 7
        Case "R": ConvertirFonemaANumero = 9
        Case "S": ConvertirFonemaANumero = 1
        Case "T": ConvertirFonemaANumero = 2
        Case "W": ConvertirFonemaANumero = 5
        Case "X": ConvertirFonemaANumero = 6
        Case "Y": ConvertirFonemaANumero = 7
        Case "Z": ConvertirFonemaANumero = 8

        Case Else
            ConvertirFonemaANumero = 0

    End Select

End Function


' ============================================================================
'  EXTRAER FONEMAS (desde texto ya normalizado)
' ============================================================================

' Segmentación fonética greedy con patrón maximal munch:
' intenta primero fonemas de 3 letras, luego de 2, y finalmente 1.
' Garantiza que CH, LL, RR, NY, etc. no se dividan incorrectamente.

Public Function ExtraerFonemasFinales(ByVal texto As String) As Collection
    Dim col As New Collection
    Dim i As Long
    Dim f3 As String, f2 As String, f1 As String

    ' Normalizar a mayúsculas para unificar
    texto = UCase$(texto)

    i = 1
    Do While i <= Len(texto)

        ' --- Intentar fonema triple ---
        If i <= Len(texto) - 2 Then
            f3 = Mid$(texto, i, 3)
            If EsFonemaCompuesto(f3) Then
                col.Add f3
                i = i + 3
                GoTo Siguiente
            End If
        End If

        ' --- Intentar fonema doble ---
        If i <= Len(texto) - 1 Then
            f2 = Mid$(texto, i, 2)
            If EsFonemaCompuesto(f2) Then
                col.Add f2
                i = i + 2
                GoTo Siguiente
            End If
        End If

        ' --- Fonema simple ---
        f1 = Mid$(texto, i, 1)
        col.Add f1
        i = i + 1

Siguiente:
    Loop

    Set ExtraerFonemasFinales = col
End Function

' ============================================================================
'  SUMA DE DÍGITOS (corregida)
' ============================================================================

Public Function SumarDigitos(num As Integer) As Integer
    Dim s As String, i As Integer
    Dim total As Integer

    s = CStr(num)
    For i = 1 To Len(s)
        total = total + CLng(Mid$(s, i, 1))
    Next i

    SumarDigitos = total
End Function

' ============================================================================
'  REDUCCIÓN SIMBÓLICA
' ============================================================================

Public Function ReducirSimbolico(num As Integer) As String
    Dim original As Integer
    Dim intermedio As Integer
    Dim reducido As Integer

    original = num
    intermedio = SumarDigitos(original)
    reducido = SumarDigitos(intermedio)
    
    If num < 10 Then
        ReducirSimbolico = original
        Exit Function
    End If

    If esMaestro(CStr(intermedio)) Or esKarmico(CStr(intermedio)) Then
        ReducirSimbolico = original & "/" & intermedio & "/" & reducido
        Exit Function
    End If

    If intermedio >= 10 Then
        ReducirSimbolico = original & "/" & intermedio & "/" & reducido
        Exit Function
    End If

    If intermedio = reducido Then
        ReducirSimbolico = original & "/" & reducido
        Exit Function
    End If

    ReducirSimbolico = original & "/" & intermedio & "/" & reducido
End Function

' ============================================================================
'  MAESTROS Y KÁRMICOS
' ============================================================================

Public Function esMaestro(Valor As String) As Boolean
    Select Case Valor
        Case "11", "22", "33", "44"
            esMaestro = True
    End Select
End Function

Public Function esKarmico(Valor As String) As Boolean
    Select Case Valor
        Case "13", "14", "16", "19"
            esKarmico = True
    End Select
End Function

' ============================================================================
'  FECHAS
' ============================================================================

Public Function FechaValida(D As Integer, M As Integer, a As Integer) As Boolean
    On Error GoTo ErrHandler
    
    ' Intentar construir la fecha
    Dim f As Date
    f = DateSerial(a, M, D)
    
    ' Comprobar que el día no ha sido corregido por DateSerial
    If Day(f) = D And Month(f) = M And Year(f) = a Then
        FechaValida = True
    Else
        FechaValida = False
    End If
    
    Exit Function

ErrHandler:
    FechaValida = False
End Function


