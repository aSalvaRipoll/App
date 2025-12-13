VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCalculoNumeroPoder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ============================================================================
' Clase: clsCalculoNumeroPoder
' Descripción: Calcula el Número de Poder (suma de Primera Letra + Primera Vocal)
' Autor: Sistema de Numerología
' Fecha: 2024
' ============================================================================

Private objBase As clsCalculoBase
Private pNumeroPrimeraLetra As Integer
Private pNumeroPrimeraVocal As Integer
Private pPrimeraLetra As String
Private pPrimeraVocal As String

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

Public Property Get NumeroPrimeraLetra() As Integer
    NumeroPrimeraLetra = pNumeroPrimeraLetra
End Property

Public Property Get NumeroPrimeraVocal() As Integer
    NumeroPrimeraVocal = pNumeroPrimeraVocal
End Property

Public Property Get PrimeraLetra() As String
    PrimeraLetra = pPrimeraLetra
End Property

Public Property Get PrimeraVocal() As String
    PrimeraVocal = pPrimeraVocal
End Property

' ============================================================================
' MÉTODOS PÚBLICOS
' ============================================================================

Public Function Calcular() As Integer
    Dim objPrimeraLetra As clsCalculoPrimeraLetra
    Dim objPrimeraVocal As clsCalculoPrimeraVocal
    Dim suma As Long
    
    On Error GoTo ErrorHandler
    
    ' Validar nombre
    If Len(objBase.NombreCompleto) = 0 Then
        Err.Raise vbObjectError + 2036, "clsCalculoNumeroPoder", ERR_NOMBRE_VACIO
    End If
    
    ' Calcular Primera Letra
    Set objPrimeraLetra = New clsCalculoPrimeraLetra
    objPrimeraLetra.NombreCompleto = objBase.NombreCompleto
    pNumeroPrimeraLetra = objPrimeraLetra.Calcular()
    pPrimeraLetra = objPrimeraLetra.PrimeraLetra
    
    ' Calcular Primera Vocal
    Set objPrimeraVocal = New clsCalculoPrimeraVocal
    objPrimeraVocal.NombreCompleto = objBase.NombreCompleto
    pNumeroPrimeraVocal = objPrimeraVocal.Calcular()
    pPrimeraVocal = objPrimeraVocal.PrimeraVocal
    
    ' Sumar y reducir
    suma = CLng(pNumeroPrimeraLetra) + CLng(pNumeroPrimeraVocal)
    objBase.Resultado = ReducirADigito(suma, True)
    
    Calcular = objBase.Resultado
    
    ' Limpiar objetos
    Set objPrimeraLetra = Nothing
    Set objPrimeraVocal = Nothing
    
    Exit Function
    
ErrorHandler:
    Set objPrimeraLetra = Nothing
    Set objPrimeraVocal = Nothing
    Err.Raise Err.Number, "clsCalculoNumeroPoder.Calcular", Err.Description
End Function

Public Function ObtenerInterpretacion() As String
    Dim nombreArchivo As String
    
    On Error GoTo ErrorHandler
    
    ' Verificar que se haya calculado
    If objBase.Resultado = 0 Then
        Calcular
    End If
    
    ' Construir nombre de archivo
    nombreArchivo = objBase.ObtenerNombreArchivoInterpretacion("Poder", objBase.Resultado)
    
    ' Cargar interpretación
    objBase.Interpretacion = objBase.CargarInterpretacionDesdeArchivo(nombreArchivo)
    
    ObtenerInterpretacion = objBase.Interpretacion
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, "clsCalculoNumeroPoder.ObtenerInterpretacion", Err.Description
End Function

Public Function ObtenerDetalles() As String
    Dim detalles As String
    
    If objBase.Resultado = 0 Then
        Calcular
    End If
    
    detalles = "Primera Letra: " & pPrimeraLetra & " = " & pNumeroPrimeraLetra & vbCrLf
    detalles = detalles & "Primera Vocal: " & pPrimeraVocal & " = " & pNumeroPrimeraVocal & vbCrLf
    detalles = detalles & "Número de Poder: " & objBase.Resultado
    
    ObtenerDetalles = detalles
End Function

' ============================================================================
' EVENTOS DE CLASE
' ============================================================================

Private Sub Class_Initialize()
    Set objBase = New clsCalculoBase
    pNumeroPrimeraLetra = 0
    pNumeroPrimeraVocal = 0
    pPrimeraLetra = ""
    pPrimeraVocal = ""
End Sub

Private Sub Class_Terminate()
    Set objBase = Nothing
End Sub