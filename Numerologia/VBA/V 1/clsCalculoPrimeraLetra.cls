VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCalculoPrimeraLetra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ============================================================================
' Clase: clsCalculoPrimeraLetra
' Descripción: Calcula el Número de la Primera Letra (Piedra Angular)
' Autor: Sistema de Numerología
' Fecha: 2024
' ============================================================================

Private objBase As clsCalculoBase
Private pPrimeraLetra As String

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

Public Property Get PrimeraLetra() As String
    PrimeraLetra = pPrimeraLetra
End Property

' ============================================================================
' MÉTODOS PÚBLICOS
' ============================================================================

Public Function Calcular() As Integer
    Dim nombreNormalizado As String
    
    On Error GoTo ErrorHandler
    
    ' Validar nombre
    If Len(objBase.NombreCompleto) = 0 Then
        Err.Raise vbObjectError + 2028, "clsCalculoPrimeraLetra", ERR_NOMBRE_VACIO
    End If
    
    nombreNormalizado = objBase.NombreCompleto
    
    ' Obtener primera letra (ignorando espacios)
    pPrimeraLetra = Left(Trim(nombreNormalizado), 1)
    
    ' Obtener valor numérico
    objBase.Resultado = ObtenerValorLetra(pPrimeraLetra)
    
    Calcular = objBase.Resultado
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, "clsCalculoPrimeraLetra.Calcular", Err.Description
End Function

Public Function ObtenerInterpretacion() As String
    Dim nombreArchivo As String
    
    On Error GoTo ErrorHandler
    
    ' Verificar que se haya calculado
    If objBase.Resultado = 0 Then
        Calcular
    End If
    
    ' Construir nombre de archivo usando el número
    nombreArchivo = objBase.ObtenerNombreArchivoInterpretacion("PrimeraLetra", objBase.Resultado)
    
    ' Cargar interpretación
    objBase.Interpretacion = objBase.CargarInterpretacionDesdeArchivo(nombreArchivo)
    
    ObtenerInterpretacion = objBase.Interpretacion
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, "clsCalculoPrimeraLetra.ObtenerInterpretacion", Err.Description
End Function

' ============================================================================
' EVENTOS DE CLASE
' ============================================================================

Private Sub Class_Initialize()
    Set objBase = New clsCalculoBase
    pPrimeraLetra = ""
End Sub

Private Sub Class_Terminate()
    Set objBase = Nothing
End Sub