Attribute VB_Name = "modAuxCalculos"

Option Compare Database
Option Explicit


Public Function ConvertirLetraANumero(letra As String, sistema As String) As Byte
    letra = UCase(letra)

    Select Case letra
        Case "A": ConvertirLetraANumero = 1
        Case "B": ConvertirLetraANumero = 2
        Case "K": ConvertirLetraANumero = 3
        Case "D": ConvertirLetraANumero = 4
        Case "E": ConvertirLetraANumero = 5
        Case "F": ConvertirLetraANumero = 6
        Case "G": ConvertirLetraANumero = 7
        Case "H": ConvertirLetraANumero = 8
        Case "I": ConvertirLetraANumero = 9
        Case "J": ConvertirLetraANumero = 1
        Case "L": ConvertirLetraANumero = 3
        Case "M": ConvertirLetraANumero = 4
        Case "N": ConvertirLetraANumero = 5
        Case "O": ConvertirLetraANumero = 6
        Case "P": ConvertirLetraANumero = 7
        Case "R": ConvertirLetraANumero = 9
        Case "S": ConvertirLetraANumero = 1
        Case "T": ConvertirLetraANumero = 2
        Case "U": ConvertirLetraANumero = 3
        Case "V": ConvertirLetraANumero = 4
        Case "W": ConvertirLetraANumero = 5
        Case "X": ConvertirLetraANumero = 6
        Case "Y": ConvertirLetraANumero = 7
        Case "Z": ConvertirLetraANumero = 8
        Case Else
            ConvertirLetraANumero = 0
    End Select
End Function


Public Function EsMaestro(valor As String) As Boolean
    Select Case valor
        Case "11", "22", "33", "44"
            EsMaestro = True
    End Select
End Function

Public Function EsKarmico(valor As String) As Boolean
    Select Case valor
        Case "13", "14", "16", "19"
            EsKarmico = True
    End Select
End Function

Public Function SumarDigitos(num As Integer) As Integer
    Dim s As String, i As Integer
    s = CStr(num)
    For i = 1 To Len(s)
        SumarDigitos = SumarDigitos + CLng(Mid(s, i, 1))
    Next i
End Function

Public Function ReducirSimbolico(num As Integer) As String
    Dim original As Long
    Dim intermedio As Long
    Dim reducido As Long

    original = num
    intermedio = SumarDigitos(original)
    reducido = SumarDigitos(intermedio)

    ' 1. Si intermedio es Maestro o Kármico ? siempre 3 grupos
    If EsMaestro(CStr(intermedio)) Or EsKarmico(CStr(intermedio)) Then
        ReducirSimbolico = original & "/" & intermedio & "/" & reducido
        Exit Function
    End If

    ' 2. Si intermedio tiene 2 cifras ? siempre 3 grupos
    If intermedio >= 10 Then
        ReducirSimbolico = original & "/" & intermedio & "/" & reducido
        Exit Function
    End If

    ' 3. Si intermedio = reducido ? solo 2 grupos
    If intermedio = reducido Then
        ReducirSimbolico = original & "/" & reducido
        Exit Function
    End If

    ' 4. Caso general ? 3 grupos
    ReducirSimbolico = original & "/" & intermedio & "/" & reducido
End Function

Public Function ExtraerFonemasFinales(ByVal texto As String) As Collection
    Dim col As New Collection
    Dim i As Long
    Dim f3 As String, f2 As String, f1 As String

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


Private Function EsFonemaCompuesto(f As String) As Boolean
    Select Case f
        ' --- Fonemas triples ---
'        Case "GWE", "GWI"

        ' --- Fonemas dobles ---
        Case "CH", "LL", "RR", "NY", "SH", "TS", "TX", "KS", "TZ", "DJ", "GW"

            EsFonemaCompuesto = True
        Case Else
            EsFonemaCompuesto = False
    End Select
End Function


Public Function ConvertirFonemaANumero(f As String) As Integer
    Select Case f

        ' --- Vocales ---
        Case "A": ConvertirFonemaANumero = 1
        Case "E": ConvertirFonemaANumero = 5
        Case "I": ConvertirFonemaANumero = 9
        Case "O": ConvertirFonemaANumero = 6
        Case "U": ConvertirFonemaANumero = 3

        ' --- Consonantes simples ---
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
        Case "X": ConvertirFonemaANumero = 6
        Case "Y": ConvertirFonemaANumero = 7
        Case "Z": ConvertirFonemaANumero = 8

        ' --- Fonemas dobles ---
        Case "NY": ConvertirFonemaANumero = 7
        Case "CH": ConvertirFonemaANumero = 6
        Case "SH": ConvertirFonemaANumero = 1
        Case "TS": ConvertirFonemaANumero = 2
        Case "TX": ConvertirFonemaANumero = 6
        Case "KS": ConvertirFonemaANumero = 6
        Case "TZ": ConvertirFonemaANumero = 8
        Case "DJ": ConvertirFonemaANumero = 1
        Case "GW": ConvertirFonemaANumero = 7

    End Select
End Function

