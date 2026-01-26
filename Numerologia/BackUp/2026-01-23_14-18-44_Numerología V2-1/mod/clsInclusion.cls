VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInclusion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ====================================================================
' Clase generada automáticamente desde la tabla: tbuinclusion
' ====================================================================

' clsInclusion:
' DTO de Inclusión: gestiona frecuencias del 1 al 9
' No calcula valores numerológicos: solo interpreta frecuencias ya calculadas.
' Gestiona las frecuencias del 1 al 9 y expone:
'   - Dominantes
'   - Ausentes
'   - Valores individuales (N1…N9)

Option Compare Database
Option Explicit

'Public Class clsInclusion

    Public IDFonetica As Long ' Campo: IDFonetica
    Public IDResultado As Long ' Campo: IDResultado
    Public IDPersona As Long ' Campo: IDPersona
    Public IDInclusion As Long
    
    Public N1 As Byte ' Campo: N1
    Public N2 As Byte ' Campo: N2
    Public N3 As Byte ' Campo: N3
    Public N4 As Byte ' Campo: N4
    Public N5 As Byte ' Campo: N5
    Public N6 As Byte ' Campo: N6
    Public N7 As Byte ' Campo: N7
    Public N8 As Byte ' Campo: N8
    Public N9 As Byte ' Campo: N9

'End Class
' ====================================================================

Public Sub AddNumero(ByVal n As Byte)
    Select Case n
        Case 1: N1 = N1 + 1
        Case 2: N2 = N2 + 1
        Case 3: N3 = N3 + 1
        Case 4: N4 = N4 + 1
        Case 5: N5 = N5 + 1
        Case 6: N6 = N6 + 1
        Case 7: N7 = N7 + 1
        Case 8: N8 = N8 + 1
        Case 9: N9 = N9 + 1
    End Select
End Sub

Public Function GetMaximos() As Collection
    Dim col As New Collection
    Dim maxVal As Byte
    Dim i As Integer
    Dim Valores(1 To 9) As Byte
    
    ' Cargar valores
    Call CargarValores(Valores)
    
    ' Encontrar el máximo
    For i = 1 To 9
        If Valores(i) > maxVal Then maxVal = Valores(i)
    Next
    
    ' Si el máximo es 0, no hay máximos (nombre vacío)
    If maxVal = 0 Then
        Set GetMaximos = col
        Exit Function
    End If
    
    ' Añadir todos los números que tengan ese máximo
    For i = 1 To 9
        If Valores(i) = maxVal Then col.Add i
    Next
    
    Set GetMaximos = col
End Function

Public Function GetAusentes() As Collection
    Dim col As New Collection
    Dim i As Integer
    Dim Valores(1 To 9) As Byte
    
    Call CargarValores(Valores)
    
    For i = 1 To 9
        If Valores(i) = 0 Then col.Add i
    Next
    
    Set GetAusentes = col
End Function

Public Function GetDominantes() As Collection
    Dim col As New Collection
    Dim Valores(1 To 9) As Byte
    Dim i As Integer
    Dim maxVal As Byte

    Call CargarValores(Valores)
    
    ' máximo solo entre presentes (>0)
    For i = 1 To 9
        If Valores(i) > 0 And Valores(i) > maxVal Then
            maxVal = Valores(i)
        End If
    Next

    If maxVal = 0 Then
        Set GetDominantes = col
        Exit Function
    End If

    ' dominantes = = (maxVal - 1) y >0
    For i = 1 To 9
        If Valores(i) > 0 And Valores(i) >= maxVal - 1 Then
            col.Add i
        End If
    Next

    Set GetDominantes = col

End Function

Public Property Get Valores(i As Integer) As Byte
    Valores = Choose(i, N1, N2, N3, N4, N5, N6, N7, N8, N9)
End Property


'Public Function GetDominantes() As Collection
'    Dim col As New Collection
'    Dim valores(1 To 9) As Byte
'    Dim i As Integer
'    Dim suma As Long
'    Dim media As Double
'    Dim umbral As Double
'
'    Call CargarValores(valores)
'
''    valores(1) = N1
''    valores(2) = N2
''    valores(3) = N3
''    valores(4) = N4
''    valores(5) = N5
''    valores(6) = N6
''    valores(7) = N7
''    valores(8) = N8
''    valores(9) = N9
'
'    For i = 1 To 9
'        suma = suma + valores(i)
'    Next
'
'    media = suma / 9
'    umbral = media * 0.5
'
'    For i = 1 To 9
'        If valores(i) > umbral Then col.Add i
'    Next
'
'    Set GetDominantes = col
'End Function

Public Function GetCuenta() As Byte
    
    Dim Valores(1 To 9) As Byte
    Dim i As Integer
    Dim suma As Byte
    
    suma = 0
    Call CargarValores(Valores)
    
    For i = 1 To 9
        If Valores(i) > 0 Then
            suma = suma + 1
        End If
    Next
 
    GetCuenta = suma
    
End Function

Private Sub CargarValores(ByRef v() As Byte)
    
    v(1) = N1
    v(2) = N2
    v(3) = N3
    v(4) = N4
    v(5) = N5
    v(6) = N6
    v(7) = N7
    v(8) = N8
    v(9) = N9
End Sub

Private Sub Class_Initialize()
    N1 = 0: N2 = 0: N3 = 0: N4 = 0: N5 = 0
    N6 = 0: N7 = 0: N8 = 0: N9 = 0
End Sub

