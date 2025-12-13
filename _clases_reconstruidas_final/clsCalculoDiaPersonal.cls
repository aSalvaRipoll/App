VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCalculoDiaPersonal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


' ============================================================================
' Clase: clsCalculoDiaPersonal
' Descripción: Calcula el Día Personal
' Autor: Sistema de Numerología
' Fecha: 2024
' ============================================================================

Private objBase As clsCalculoBase
Private pFechaCalculo As Date
Private pNumeroMesPersonal As Integer

' ============================================================================
' PROPIEDADES PÚBLICAS
' ============================================================================

Public Property Get FechaNacimiento() As Date
    FechaNacimiento = objBase.FechaNacimiento
End Property

Public Property Let FechaNacimiento(ByVal valor As Date)
    objBase.FechaNacimiento = valor
End Property

Public Property Get FechaCalculo() As Date
    FechaCalculo = pFechaCalculo
End Property

Public Property Let FechaCalculo(ByVal valor As Date)
    If Not ValidarFecha(valor) Then
        err.Raise vbObjectError + 2012, "clsCalculoDiaPersonal", ERR_FECHA_INVALIDA
    End If
    pFechaCalculo = valor
End Property

Public Property Get Resultado() As Integer
    Resultado = objBase.Resultado
End Property

Public Property Get NumeroMesPersonal() As Integer
    NumeroMesPersonal = pNumeroMesPersonal
End Property

' ============================================================================
' MÉTODOS PÚBLICOS
' ============================================================================

Public Function Calcular() As Integer
    Dim objMesPersonal As clsCalculoMesPersonal
    Dim diaCalculo As Integer
    Dim suma As Long
    
    On Error GoTo ErrorHandler
    
    ' Validar datos
    If objBase.FechaNacimiento = 0 Then
        err.Raise vbObjectError + 2013, "clsCalculoDiaPersonal", ERR_FECHA_INVALIDA
    End If
    
    If pFechaCalculo = 0 Then
        pFechaCalculo = Date  ' Usar fecha actual si no se especificó
    End If
    
    ' Calcular Mes Personal
    Set objMesPersonal = New clsCalculoMesPersonal
    objMesPersonal.FechaNacimiento = objBase.FechaNacimiento
    objMesPersonal.MesCalculo = Month(pFechaCalculo)
    objMesPersonal.AnoCalculo = Year(pFechaCalculo)
    pNumeroMesPersonal = objMesPersonal.Calcular()
    
    ' Obtener día de la fecha de cálculo
    diaCalculo = Day(pFechaCalculo)
    
    ' Sumar día + mes personal
    suma = CLng(diaCalculo) + CLng(pNumeroMesPersonal)
    
    ' Reducir a dígito
    objBase.Resultado = ReducirADigito(suma, False)
    
    Calcular = objBase.Resultado
    
    ' Limpiar objetos
    Set objMesPersonal = Nothing
    
    Exit Function
    
ErrorHandler:
    Set objMesPersonal = Nothing
    err.Raise err.Number, "clsCalculoDiaPersonal.Calcular", err.Description
End Function

' ============================================================================
' EVENTOS DE CLASE
' ============================================================================

Private Sub Class_Initialize()
    Set objBase = New clsCalculoBase
    pFechaCalculo = 0
    pNumeroMesPersonal = 0
End Sub

Private Sub Class_Terminate()
    Set objBase = Nothing
End Sub

