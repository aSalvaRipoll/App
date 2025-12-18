VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCalculoDesafio3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


' ============================================================================
' Clase: clsCalculoDesafio3
' Descripción: Calcula el Tercer Desafío (Desafío Mayor)
' Autor: Sistema de Numerología
' Fecha: 2024
' ============================================================================

Private objBase As clsCalculoBase
Private pEdadInicio As Integer
Private pEdadFin As Integer

' ============================================================================
' PROPIEDADES PÚBLICAS
' ============================================================================

Public Property Get FechaNacimiento() As Date
    FechaNacimiento = objBase.FechaNacimiento
End Property

Public Property Let FechaNacimiento(ByVal valor As Date)
    objBase.FechaNacimiento = valor
End Property

Public Property Get Resultado() As Integer
    Resultado = objBase.Resultado
End Property

Public Property Get Interpretacion() As String
    Interpretacion = objBase.Interpretacion
End Property

Public Property Get EdadInicio() As Integer
    EdadInicio = pEdadInicio
End Property

Public Property Get EdadFin() As Integer
    EdadFin = pEdadFin
End Property

' ============================================================================
' MÉTODOS PÚBLICOS
' ============================================================================

Public Function Calcular() As Integer
    Dim objDesafio1 As clsCalculoDesafio1
    Dim objDesafio2 As clsCalculoDesafio2
    Dim desafio1Valor As Integer
    Dim desafio2Valor As Integer
    
    On Error GoTo ErrorHandler
    
    ' Validar fecha
    If objBase.FechaNacimiento = 0 Then
        err.Raise vbObjectError + 2024, "clsCalculoDesafio3", ERR_FECHA_INVALIDA
    End If
    
    ' El Desafío 3 = |Desafío 1 - Desafío 2|
    Set objDesafio1 = New clsCalculoDesafio1
    objDesafio1.FechaNacimiento = objBase.FechaNacimiento
    desafio1Valor = objDesafio1.Calcular()
    
    Set objDesafio2 = New clsCalculoDesafio2
    objDesafio2.FechaNacimiento = objBase.FechaNacimiento
    desafio2Valor = objDesafio2.Calcular()
    
    ' Diferencia absoluta
    objBase.Resultado = Abs(desafio1Valor - desafio2Valor)
    
    ' Calcular edades del desafío
    pEdadInicio = objDesafio2.EdadFin + 1
    pEdadFin = pEdadInicio + 8
    
    Calcular = objBase.Resultado
    
    ' Limpiar objetos
    Set objDesafio1 = Nothing
    Set objDesafio2 = Nothing
    
    Exit Function
    
ErrorHandler:
    Set objDesafio1 = Nothing
    Set objDesafio2 = Nothing
    err.Raise err.Number, "clsCalculoDesafio3.Calcular", err.Description
End Function

Public Function ObtenerInterpretacion() As String
    Dim nombreArchivo As String
    
    On Error GoTo ErrorHandler
    
    ' Verificar que se haya calculado
    If objBase.Resultado < 0 Then
        Calcular
    End If
    
    ' Construir nombre de archivo
    nombreArchivo = objBase.ObtenerNombreArchivoInterpretacion("Desafio3", objBase.Resultado)
    If objBase.Resultado = 0 Then
        nombreArchivo = "00_Desafio3"
    End If
    
    ' Cargar interpretación
    objBase.Interpretacion = objBase.CargarInterpretacionDesdeArchivo(nombreArchivo)
    
    ObtenerInterpretacion = objBase.Interpretacion
    Exit Function
    
ErrorHandler:
    err.Raise err.Number, "clsCalculoDesafio3.ObtenerInterpretacion", err.Description
End Function

Public Function ObtenerRangoEdades() As String
    If pEdadInicio = 0 And pEdadFin = 0 Then
        Calcular
    End If
    
    ObtenerRangoEdades = pEdadInicio & " - " & pEdadFin & " años"
End Function

' ============================================================================
' EVENTOS DE CLASE
' ============================================================================

Private Sub Class_Initialize()
    Set objBase = New clsCalculoBase
    pEdadInicio = 0
    pEdadFin = 0
End Sub

Private Sub Class_Terminate()
    Set objBase = Nothing
End Sub

