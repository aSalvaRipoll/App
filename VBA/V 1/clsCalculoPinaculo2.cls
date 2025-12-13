VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCalculoPinaculo2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ============================================================================
' Clase: clsCalculoPinaculo2
' Descripción: Calcula el Segundo Pináculo
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

' ============================================================================
' MÉTODOS PÚBLICOS
' ============================================================================

Public Function Calcular() As Integer
    Dim dia As Integer
    Dim ano As Integer
    Dim suma As Long
    Dim objCaminoVida As clsCalculoCaminoVida
    Dim objPinaculo1 As clsCalculoPinaculo1
    Dim numeroCaminoVida As Integer
    
    On Error GoTo ErrorHandler
    
    ' Validar fecha
    If objBase.FechaNacimiento = 0 Then
        Err.Raise vbObjectError + 2019, "clsCalculoPinaculo2", ERR_FECHA_INVALIDA
    End If
    
    ' El Pináculo 2 = Día + Año reducidos
    dia = Day(objBase.FechaNacimiento)
    ano = Year(objBase.FechaNacimiento)
    
    suma = CLng(ReducirADigito(dia, True)) + CLng(ReducirADigito(ano, True))
    objBase.Resultado = ReducirADigito(suma, True)
    
    ' Calcular edades del pináculo
    Set objCaminoVida = New clsCalculoCaminoVida
    objCaminoVida.FechaNacimiento = objBase.FechaNacimiento
    numeroCaminoVida = objCaminoVida.Calcular()
    
    Set objPinaculo1 = New clsCalculoPinaculo1
    objPinaculo1.FechaNacimiento = objBase.FechaNacimiento
    objPinaculo1.Calcular
    
    pEdadInicio = objPinaculo1.EdadFin + 1
    pEdadFin = pEdadInicio + 8
    
    Calcular = objBase.Resultado
    
    ' Limpiar objetos
    Set objCaminoVida = Nothing
    Set objPinaculo1 = Nothing
    
    Exit Function
    
ErrorHandler:
    Set objCaminoVida = Nothing
    Set objPinaculo1 = Nothing
    Err.Raise Err.Number, "clsCalculoPinaculo2.Calcular", Err.Description
End Function

Public Function ObtenerInterpretacion() As String
    Dim nombreArchivo As String
    
    On Error GoTo ErrorHandler
    
    ' Verificar que se haya calculado
    If objBase.Resultado = 0 Then
        Calcular
    End If
    
    ' Construir nombre de archivo
    nombreArchivo = objBase.ObtenerNombreArchivoInterpretacion("Pinaculo2", objBase.Resultado)
    
    ' Cargar interpretación
    objBase.Interpretacion = objBase.CargarInterpretacionDesdeArchivo(nombreArchivo)
    
    ObtenerInterpretacion = objBase.Interpretacion
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, "clsCalculoPinaculo2.ObtenerInterpretacion", Err.Description
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