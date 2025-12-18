VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCalculoDestino"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ============================================================================
' Clase: clsCalculoDestino
' Descripción: Calcula el Número de Destino (Expresión)
' Autor: Sistema de Numerología
' Fecha: 2024
' ============================================================================

Private objBase As clsCalculoBase

' ============================================================================
' PROPIEDADES PÚBLICAS
' ============================================================================

Public Property Get NombreCompleto() As String
    NombreCompleto = objBase.NombreCompleto
End Property

Public Property Let NombreCompleto(ByVal valor As String)
    objBase.NombreCompleto = valor
End Property

Public Property Get Resultado() As Integer
    Resultado = objBase.Resultado
End Property

Public Property Get Interpretacion() As String
    Interpretacion = objBase.Interpretacion
End Property

Public Property Get EsKarmico() As Boolean
    EsKarmico = objBase.EsKarmico
End Property

Public Property Get EsMaestro() As Boolean
    EsMaestro = objBase.EsMaestro
End Property

' ============================================================================
' MÉTODOS PÚBLICOS
' ============================================================================

Public Function Calcular() As Integer
    On Error GoTo ErrorHandler
    
    ' Validar nombre
    If Len(objBase.NombreCompleto) = 0 Then
        err.Raise vbObjectError + 2002, "clsCalculoDestino", ERR_NOMBRE_VACIO
    End If
    
    ' Sumar todos los valores del nombre completo
    objBase.Resultado = objBase.SumarValoresNombre(objBase.NombreCompleto, False, False)
    
    Calcular = objBase.Resultado
    Exit Function
    
ErrorHandler:
    err.Raise err.Number, "clsCalculoDestino.Calcular", err.Description
End Function

Public Function ObtenerInterpretacion() As String
    Dim nombreArchivo As String
    
    On Error GoTo ErrorHandler
    
    ' Verificar que se haya calculado
    If objBase.Resultado = 0 Then
        Calcular
    End If
    
    ' Construir nombre de archivo
    nombreArchivo = objBase.ObtenerNombreArchivoInterpretacion("Destino", objBase.Resultado)
    
    ' Cargar interpretación
    objBase.Interpretacion = objBase.CargarInterpretacionDesdeArchivo(nombreArchivo)
    
    ObtenerInterpretacion = objBase.Interpretacion
    Exit Function
    
ErrorHandler:
    err.Raise err.Number, "clsCalculoDestino.ObtenerInterpretacion", err.Description
End Function

' ============================================================================
' EVENTOS DE CLASE
' ============================================================================

Private Sub Class_Initialize()
    Set objBase = New clsCalculoBase
End Sub

Private Sub Class_Terminate()
    Set objBase = Nothing
End Sub
