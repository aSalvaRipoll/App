VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCalculoMesPersonal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ============================================================================
' Clase: clsCalculoMesPersonal
' Descripción: Calcula el Mes Personal
' Autor: Sistema de Numerología
' Fecha: 2024
' ============================================================================

Private objBase As clsCalculoBase
Private pMesCalculo As Integer
Private pAnoCalculo As Integer
Private pNumeroAnoPersonal As Integer

' ============================================================================
' PROPIEDADES PÚBLICAS
' ============================================================================

Public Property Get FechaNacimiento() As Date
    FechaNacimiento = objBase.FechaNacimiento
End Property

Public Property Let FechaNacimiento(ByVal valor As Date)
    objBase.FechaNacimiento = valor
End Property

Public Property Get MesCalculo() As Integer
    MesCalculo = pMesCalculo
End Property

Public Property Let MesCalculo(ByVal valor As Integer)
    If valor < 1 Or valor > 12 Then
        Err.Raise vbObjectError + 2009, "clsCalculoMesPersonal", "Mes fuera de rango válido (1-12)"
    End If
    pMesCalculo = valor
End Property

Public Property Get AnoCalculo() As Integer
    AnoCalculo = pAnoCalculo
End Property

Public Property Let AnoCalculo(ByVal valor As Integer)
    If valor < 1900 Or valor > 2100 Then
        Err.Raise vbObjectError + 2010, "clsCalculoMesPersonal", "Año fuera de rango válido"
    End If
    pAnoCalculo = valor
End Property

Public Property Get Resultado() As Integer
    Resultado = objBase.Resultado
End Property

Public Property Get NumeroAnoPersonal() As Integer
    NumeroAnoPersonal = pNumeroAnoPersonal
End Property

' ============================================================================
' MÉTODOS PÚBLICOS
' ============================================================================

Public Function Calcular() As Integer
    Dim objAnoPersonal As clsCalculoAnoPersonal
    Dim suma As Long
    
    On Error GoTo ErrorHandler
    
    ' Validar datos
    If objBase.FechaNacimiento = 0 Then
        Err.Raise vbObjectError + 2011, "clsCalculoMesPersonal", ERR_FECHA_INVALIDA
    End If
    
    If pMesCalculo = 0 Then
        pMesCalculo = Month(Date)  ' Usar mes actual si no se especificó
    End If
    
    If pAnoCalculo = 0 Then
        pAnoCalculo = Year(Date)  ' Usar año actual si no se especificó
    End If
    
    ' Calcular Año Personal
    Set objAnoPersonal = New clsCalculoAnoPersonal
    objAnoPersonal.FechaNacimiento = objBase.FechaNacimiento
    objAnoPersonal.AnoCalculo = pAnoCalculo
    pNumeroAnoPersonal = objAnoPersonal.Calcular()
    
    ' Sumar mes + año personal
    suma = CLng(pMesCalculo) + CLng(pNumeroAnoPersonal)
    
    ' Reducir a dígito
    objBase.Resultado = ReducirADigito(suma, False)
    
    Calcular = objBase.Resultado
    
    ' Limpiar objetos
    Set objAnoPersonal = Nothing
    
    Exit Function
    
ErrorHandler:
    Set objAnoPersonal = Nothing
    Err.Raise Err.Number, "clsCalculoMesPersonal.Calcular", Err.Description
End Function

' ============================================================================
' EVENTOS DE CLASE
' ============================================================================

Private Sub Class_Initialize()
    Set objBase = New clsCalculoBase
    pMesCalculo = 0
    pAnoCalculo = 0
    pNumeroAnoPersonal = 0
End Sub

Private Sub Class_Terminate()
    Set objBase = Nothing
End Sub