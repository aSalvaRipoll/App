VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCalculoPinaculo4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ============================================================================
' Clase: clsCalculoPinaculo4
' Descripción: Calcula el Cuarto Pináculo
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
    Dim mes As Integer
    Dim ano As Integer
    Dim suma As Long
    Dim objPinaculo3 As clsCalculoPinaculo3
    
    On Error GoTo ErrorHandler
    
    ' Validar fecha
    If objBase.FechaNacimiento = 0 Then
        Err.Raise vbObjectError + 2021, "clsCalculoPinaculo4", ERR_FECHA_INVALIDA
    End If
    
    ' El Pináculo 4 = Mes + Año reducidos
    mes = Month(objBase.FechaNacimiento)
    ano = Year(objBase.FechaNacimiento)
    
    suma = CLng(ReducirADigito(mes, True)) + CLng(ReducirADigito(ano, True))
    objBase.Resultado = ReducirADigito(suma, True)
    
    ' Calcular edades del pináculo
    Set objPinaculo3 = New clsCalculoPinaculo3
    objPinaculo3.FechaNacimiento = objBase.FechaNacimiento
    objPinaculo3.Calcular
    
    pEdadInicio = objPinaculo3.EdadFin + 1
    pEdadFin = 999  ' Hasta el final de la vida
    
    Calcular = objBase.Resultado
    
    ' Limpiar objetos
    Set objPinaculo3 = Nothing
    
    Exit Function
    
ErrorHandler:
    Set objPinaculo3 = Nothing
    Err.Raise Err.Number, "clsCalculoPinaculo4.Calcular", Err.Description
End Function

Public Function ObtenerInterpretacion() As String
    Dim nombreArchivo As String
    
    On Error GoTo ErrorHandler
    
    ' Verificar que se haya calculado
    If objBase.Resultado = 0 Then
        Calcular
    End If
    
    ' Construir nombre de archivo
    nombreArchivo = objBase.ObtenerNombreArchivoInterpretacion("Pinaculo4", objBase.Resultado)
    
    ' Cargar interpretación
    objBase.Interpretacion = objBase.CargarInterpretacionDesdeArchivo(nombreArchivo)
    
    ObtenerInterpretacion = objBase.Interpretacion
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, "clsCalculoPinaculo4.ObtenerInterpretacion", Err.Description
End Function

Public Function ObtenerRangoEdades() As String
    If pEdadInicio = 0 Then
        Calcular
    End If
    
    ObtenerRangoEdades = pEdadInicio & " años en adelante"
End Function

' ============================================================================
' EVENTOS DE CLASE
' ============================================================================

Private Sub Class_Initialize()
    Set objBase = New clsCalculoBase
    pEdadInicio = 0
    pEdadFin = 999
End Sub

Private Sub Class_Terminate()
    Set objBase = Nothing
End Sub