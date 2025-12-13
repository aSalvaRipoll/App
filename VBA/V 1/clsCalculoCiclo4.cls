VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCalculoCiclo4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ============================================================================
' Clase: clsCalculoCiclo4
' Descripción: Calcula el Cuarto Ciclo de Vida (Ciclo Espiritual - Opcional)
' Nota: Algunos sistemas numerológicos solo usan 3 ciclos
' Autor: Sistema de Numerología
' Fecha: 2024
' ============================================================================

Private objBase As clsCalculoBase
Private pEdadInicio As Integer
Private pEdadFin As Integer
Private pEsAplicable As Boolean

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

Public Property Get EsKarmico() As Boolean
    EsKarmico = objBase.EsKarmico
End Property

Public Property Get EsMaestro() As Boolean
    EsMaestro = objBase.EsMaestro
End Property

Public Property Get EdadInicio() As Integer
    EdadInicio = pEdadInicio
End Property

Public Property Get EdadFin() As Integer
    EdadFin = pEdadFin
End Property

Public Property Get EsAplicable() As Boolean
    EsAplicable = pEsAplicable
End Property

' ============================================================================
' MÉTODOS PÚBLICOS
' ============================================================================

Public Function Calcular() As Integer
    Dim objCaminoVida As clsCalculoCaminoVida
    Dim numeroCaminoVida As Integer
    
    On Error GoTo ErrorHandler
    
    ' Validar fecha
    If objBase.FechaNacimiento = 0 Then
        Err.Raise vbObjectError + 2017, "clsCalculoCiclo4", ERR_FECHA_INVALIDA
    End If
    
    ' Calcular Camino de Vida
    Set objCaminoVida = New clsCalculoCaminoVida
    objCaminoVida.FechaNacimiento = objBase.FechaNacimiento
    numeroCaminoVida = objCaminoVida.Calcular()
    
    ' El Ciclo 4 es igual al Camino de Vida
    objBase.Resultado = numeroCaminoVida
    
    ' Este ciclo solo aplica en algunas interpretaciones
    ' Se considera activo después de los 60-70 años
    pEdadInicio = 63
    pEdadFin = 999
    pEsAplicable = True  ' Puede configurarse según el sistema usado
    
    Calcular = objBase.Resultado
    
    ' Limpiar objetos
    Set objCaminoVida = Nothing
    
    Exit Function
    
ErrorHandler:
    Set objCaminoVida = Nothing
    Err.Raise Err.Number, "clsCalculoCiclo4.Calcular", Err.Description
End Function

Public Function ObtenerInterpretacion() As String
    Dim nombreArchivo As String
    
    On Error GoTo ErrorHandler
    
    ' Verificar que se haya calculado
    If objBase.Resultado = 0 Then
        Calcular
    End If
    
    If Not pEsAplicable Then
        ObtenerInterpretacion = "El Cuarto Ciclo es opcional en este sistema numerológico."
        Exit Function
    End If
    
    ' Construir nombre de archivo
    nombreArchivo = objBase.ObtenerNombreArchivoInterpretacion("Ciclo4", objBase.Resultado)
    
    ' Cargar interpretación
    objBase.Interpretacion = objBase.CargarInterpretacionDesdeArchivo(nombreArchivo)
    
    ObtenerInterpretacion = objBase.Interpretacion
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, "clsCalculoCiclo4.ObtenerInterpretacion", Err.Description
End Function

Public Function ObtenerRangoEdades() As String
    If pEdadInicio = 0 Then
        Calcular
    End If
    
    If Not pEsAplicable Then
        ObtenerRangoEdades = "No aplicable"
    Else
        ObtenerRangoEdades = pEdadInicio & " años en adelante"
    End If
End Function

' ============================================================================
' EVENTOS DE CLASE
' ============================================================================

Private Sub Class_Initialize()
    Set objBase = New clsCalculoBase
    pEdadInicio = 63
    pEdadFin = 999
    pEsAplicable = True
End Sub

Private Sub Class_Terminate()
    Set objBase = Nothing
End Sub