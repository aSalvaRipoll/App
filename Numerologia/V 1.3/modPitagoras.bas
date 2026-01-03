Attribute VB_Name = "modPitagoras"
Option Compare Database

' ================================
'  modPitagoras.bas
'  Tabla pitagórica fonética
' ================================

Option Explicit
Private mPitagoras As Scripting.Dictionary

Public Sub InicializarPitagoras()
    If Not mPitagoras Is Nothing Then Exit Sub
    
    Set mPitagoras = New Scripting.Dictionary
    mPitagoras.CompareMode = TextCompare

    ' Vocales
    mPitagoras.Add "A", 1
    mPitagoras.Add "E", 5
    mPitagoras.Add "I", 9
    mPitagoras.Add "O", 7
    mPitagoras.Add "U", 7

    ' Consonantes simples
    mPitagoras.Add "B", 2
    mPitagoras.Add "C", 3
    mPitagoras.Add "K", 3
    mPitagoras.Add "Q", 3
    mPitagoras.Add "D", 4
    mPitagoras.Add "F", 6
    mPitagoras.Add "G", 7
    mPitagoras.Add "H", 8
    mPitagoras.Add "L", 2
    mPitagoras.Add "M", 4
    mPitagoras.Add "N", 5
    mPitagoras.Add "P", 8
    mPitagoras.Add "R", 9
    mPitagoras.Add "S", 2
    mPitagoras.Add "T", 3
    mPitagoras.Add "V", 8
    mPitagoras.Add "X", 9
    mPitagoras.Add "Z", 2

    ' Fonemas complejos
    mPitagoras.Add "LL", 3
    mPitagoras.Add "RR", 1
    mPitagoras.Add "CH", 4
    mPitagoras.Add "NY", 6
    mPitagoras.Add "TX", 4
    mPitagoras.Add "TZ", 5
    mPitagoras.Add "TS", 6
    mPitagoras.Add "SH", 1
    mPitagoras.Add "J", 1
End Sub

Public Function ValorPitagorico(fonema As String) As Integer
    Call InicializarPitagoras
    If mPitagoras.Exists(fonema) Then
        ValorPitagorico = mPitagoras(fonema)
    Else
        ValorPitagorico = 0 ' o error controlado
    End If
End Function



'Public Function AnalizarNumero(ByVal valor As Integer) As String
'    Dim entrada As Integer
'    Dim reduccion As Byte
'
'    entrada = valor
'
'    ' 1) Maestro
'    If valor <= 255 Then
'        If EsMaestro(CByte(valor)) Then
'            AnalizarNumero = CStr(entrada) & "/Maestro"
'            Exit Function
'        End If
'    End If
'
'    ' 2) Karmico
'    If valor <= 255 Then
'        If EsKarmico(CByte(valor)) Then
'            reduccion = ReducirKarmico(CByte(valor))
'            AnalizarNumero = CStr(entrada) & "/Karmico/" & CStr(reduccion)
'            Exit Function
'        End If
'    End If
'
'    ' 3) Normal
'    reduccion = ReducirNormal(valor)
'    AnalizarNumero = CStr(entrada) & "/" & CStr(reduccion)
'End Function

'Public Function AnalizarNumero(ByVal valor As Integer) As tResultado
'    Dim res As tResultado
'    Dim cadena As String
'    Dim actual As Integer
'    Dim siguiente As Integer
'
'    ' Validación
'    If valor < 0 Then valor = Abs(valor)
'
'    actual = valor
'    cadena = CStr(actual)
'
'    ' Reducir hasta un dígito, guardando la cadena
'    Do While actual > 9
'        actual = SumaDigitos(actual)
'        cadena = cadena & "/" & CStr(actual)
'    Loop
'
'    ' Guardar cadena completa
'    res.cadena = cadena
'
'    ' Inicial y final
'    Dim partes() As String
'    partes = Split(cadena, "/")
'
'    res.Inicial = CByte(partes(0))
'    res.Final = CByte(partes(UBound(partes)))
'
'    ' Detectar Maestro y Karmico en cualquier etapa
'    Dim i As Long, v As Byte
'    For i = 0 To UBound(partes)
'        v = CByte(partes(i))
'
'        If EsMaestro(v) Then res.Maestro = v
'        If EsKarmico(v) Then res.Karmico = v
'    Next i
'
'    AnalizarNumero = res
'End Function
'
'Public Function EsMaestro(ByVal valor As Byte) As Boolean
'    Select Case valor
'        Case 11, 22, 33, 44
'            EsMaestro = True
'    End Select
'End Function
'
'Public Function EsKarmico(ByVal valor As Byte) As Boolean
'    Select Case valor
'        Case 13, 14, 16, 19
'            EsKarmico = True
'    End Select
'End Function
'
'Private Function SumaDigitos(ByVal n As Integer) As Integer
'    Dim s As Integer
'    Do While n > 0
'        s = s + (n Mod 10)
'        n = n \ 10
'    Loop
'    SumaDigitos = s
'End Function
'
