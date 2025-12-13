VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCalculoMadurez"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ============================================================================
' Clase: clsCalculoMadurez
' Descripción: Calcula el Número de Madurez (Realización)
' Autor: Sistema de Numerología
' Fecha: 2024
' ============================================================================

Private objBase As clsCalculoBase
Private pNumeroCaminoVida As Integer
Private pNumeroDestino As Integer

' ============================================================================
' PROPIEDADES PÚBLICAS
' ============================================================================

Public Property Get FechaNacimiento() As Date
    FechaNacimiento = objBase.FechaNacimiento
End Property

Public Property Let FechaNacimiento(ByVal valor As Date)
    objBase.FechaNacimiento = valor
End Property

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

Public Property Get NumeroCaminoVida() As Integer
    NumeroCaminoVida = pNumeroCaminoVida
End Property

Public Property Get NumeroDestino() As Integer
    NumeroDestino = pNumeroDestino
End Property

' ============================================================================
' MÉTODOS PÚBLICOS
' ============================================================================

Public Function Calcular() As Integer
    Dim objCaminoVida As clsCalculoCaminoVida
    Dim objDestino As clsCalculoDestino
    Dim suma As Long
    
    On Error GoTo ErrorHandler
    
    ' Validar datos
    If objBase.FechaNacimiento = 0 Then
        Err.Raise vbObjectError + 2005, "clsCalculoMadurez", ERR_FECHA_INVALIDA
    End If
    
    If Len(objBase.NombreCompleto) = 0 Then
        Err.Raise vbObjectError + 2006, "clsCalculoMadurez", ERR_NOMBRE_VACIO
    End If
    
    ' Calcular Camino de Vida
    Set objCaminoVida = New clsCalculoCaminoVida
    objCaminoVida.FechaNacimiento = objBase.FechaNacimiento
    pNumeroCaminoVida = objCaminoVida.Calcular()
    
    ' Calcular Destino
    Set objDestino = New clsCalculoDestino
    objDestino.NombreCompleto = objBase.NombreCompleto
    pNumeroDestino = objDestino.Calcular()
    
    ' Sumar y reducir
    suma = CLng(pNumeroCaminoVida) + CLng(pNumeroDestino)
    objBase.Resultado = ReducirADigito(suma, True)
    
    Calcular = objBase.Resultado
    
    ' Limpiar objetos
    Set objCaminoVida = Nothing
    Set objDestino = Nothing
    
    Exit Function
    
ErrorHandler:
    Set objCaminoVida = Nothing
    Set objDestino = Nothing
    Err.Raise Err.Number, "clsCalculoMadurez.Calcular", Err.Description
End Function

Public Function ObtenerInterpretacion() As String
    Dim nombreArchivo As String
    
    On Error GoTo ErrorHandler
    
    ' Verificar que se haya calculado
    If objBase.Resultado = 0 Then
        Calcular
    End If
    
    ' Construir nombre de archivo
    nombreArchivo = objBase.ObtenerNombreArchivoInterpretacion("Madurez", objBase.Resultado)
    
    ' Cargar interpretación
    objBase.Interpretacion = objBase.CargarInterpretacionDesdeArchivo(nombreArchivo)
    
    ObtenerInterpretacion = objBase.Interpretacion
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, "clsCalculoMadurez.ObtenerInterpretacion", Err.Description
End Function

' ============================================================================
' EVENTOS DE CLASE
' ============================================================================

Private Sub Class_Initialize()
    Set objBase = New clsCalculoBase
    pNumeroCaminoVida = 0
    pNumeroDestino = 0
End Sub

Private Sub Class_Terminate()
    Set objBase = Nothing
End Sub