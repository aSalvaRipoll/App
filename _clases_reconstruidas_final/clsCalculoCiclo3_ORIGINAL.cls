VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCalculoCiclo3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ============================================================================
' Clase: clsCalculoCiclo3
' Descripción: Calcula el Tercer Ciclo de Vida (Ciclo de Cosecha)
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
    Dim ano As Integer
    Dim objCiclo2 As clsCalculoCiclo2
    
    On Error GoTo ErrorHandler
    
    ' Validar fecha
    If objBase.FechaNacimiento = 0 Then
        err.Raise vbObjectError + 2016, "clsCalculoCiclo3", ERR_FECHA_INVALIDA
    End If
    
    ' El Ciclo 3 se basa en el año de nacimiento
    ano = Year(objBase.FechaNacimiento)
    objBase.Resultado = ReducirADigito(ano, True)
    
    ' Calcular edades del ciclo
    Set objCiclo2 = New clsCalculoCiclo2
    objCiclo2.FechaNacimiento = objBase.FechaNacimiento
    objCiclo2.Calcular
    
    pEdadInicio = objCiclo2.EdadFin + 1
    pEdadFin = 999  ' Hasta el final de la vida
    
    Calcular = objBase.Resultado
    
    ' Limpiar objetos
    Set objCiclo2 = Nothing
    
    Exit Function
    
ErrorHandler:
    Set objCiclo2 = Nothing
    err.Raise err.Number, "clsCalculoCiclo3.Calcular", err.Description
End Function

Public Function ObtenerInterpretacion() As String
    Dim nombreArchivo As String
    
    On Error GoTo ErrorHandler
    
    ' Verificar que se haya calculado
    If objBase.Resultado = 0 Then
        Calcular
    End If
    
    ' Construir nombre de archivo
    nombreArchivo = objBase.ObtenerNombreArchivoInterpretacion("Ciclo3", objBase.Resultado)
    
    ' Cargar interpretación
    objBase.Interpretacion = objBase.CargarInterpretacionDesdeArchivo(nombreArchivo)
    
    ObtenerInterpretacion = objBase.Interpretacion
    Exit Function
    
ErrorHandler:
    err.Raise err.Number, "clsCalculoCiclo3.ObtenerInterpretacion", err.Description
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
