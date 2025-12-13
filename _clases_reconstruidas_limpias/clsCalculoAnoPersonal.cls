VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCalculoAnoPersonal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ============================================================================
' Proyecto:     Sistema de Numerología Tradicional y Fonético
' Clase: clsCalculoAnoPersonal
' Descripción: Calcula el Año Personal
' Autor: Sistema de Numerología
' Fecha: 2024
' ============================================================================

Private objBase As clsCalculoBase
Private pAnoCalculo As Integer

' ============================================================================
' PROPIEDADES PÚBLICAS
' ============================================================================

Public Property Get FechaNacimiento() As Date
    FechaNacimiento = objBase.FechaNacimiento
End Property

Public Property Let FechaNacimiento(ByVal valor As Date)
    objBase.FechaNacimiento = valor
End Property

Public Property Get AnoCalculo() As Integer
    AnoCalculo = pAnoCalculo
End Property

Public Property Let AnoCalculo(ByVal valor As Integer)
    If valor < 1900 Or valor > 2100 Then
        err.Raise vbObjectError + 2007, "clsCalculoAnoPersonal", "Año fuera de rango válido"
    End If
    pAnoCalculo = valor
End Property

Public Property Get Resultado() As Integer
    Resultado = objBase.Resultado
End Property

Public Property Get Interpretacion() As String
    Interpretacion = objBase.Interpretacion
End Property

' ============================================================================
' MÉTODOS PÚBLICOS
' ============================================================================

Public Function Calcular() As Integer
    Dim dia As Integer
    Dim mes As Integer
    Dim sumaFecha As Long
    
    On Error GoTo ErrorHandler
    
    ' Validar datos
    If objBase.FechaNacimiento = 0 Then
        err.Raise vbObjectError + 2008, "clsCalculoAnoPersonal", ERR_FECHA_INVALIDA
    End If
    
    If pAnoCalculo = 0 Then
        pAnoCalculo = Year(Date)  ' Usar año actual si no se especificó
    End If
    
    ' Obtener día y mes de nacimiento
    dia = Day(objBase.FechaNacimiento)
    mes = Month(objBase.FechaNacimiento)
    
    ' Sumar día + mes + año de cálculo
    sumaFecha = CLng(dia) + CLng(mes) + CLng(pAnoCalculo)
    
    ' Reducir a dígito (Año Personal no tiene maestros)
    objBase.Resultado = ReducirADigito(sumaFecha, False)
    
    Calcular = objBase.Resultado
    Exit Function
    
ErrorHandler:
    err.Raise err.Number, "clsCalculoAnoPersonal.Calcular", err.Description
End Function

Public Function ObtenerInterpretacion() As String
    Dim nombreArchivo As String
    
    On Error GoTo ErrorHandler
    
    ' Verificar que se haya calculado
    If objBase.Resultado = 0 Then
        Calcular
    End If
    
    ' Construir nombre de archivo
    nombreArchivo = objBase.ObtenerNombreArchivoInterpretacion("AnoPersonal", objBase.Resultado)
    
    ' Cargar interpretación
    objBase.Interpretacion = objBase.CargarInterpretacionDesdeArchivo(nombreArchivo)
    
    ObtenerInterpretacion = objBase.Interpretacion
    Exit Function
    
ErrorHandler:
    err.Raise err.Number, "clsCalculoAnoPersonal.ObtenerInterpretacion", err.Description
End Function

' ============================================================================
' EVENTOS DE CLASE
' ============================================================================

Private Sub Class_Initialize()
    Set objBase = New clsCalculoBase
    pAnoCalculo = 0
End Sub

Private Sub Class_Terminate()
    Set objBase = Nothing
End Sub
