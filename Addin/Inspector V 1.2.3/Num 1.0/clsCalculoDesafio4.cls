VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCalculoDesafio4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


' ============================================================================
' Clase: clsCalculoDesafio4
' Descripción: Calcula el Cuarto Desafío
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
    Dim mes As Integer
    Dim ano As Integer
    Dim mesReducido As Integer
    Dim anoReducido As Integer
    Dim objDesafio3 As clsCalculoDesafio3
    
    On Error GoTo ErrorHandler
    
    ' Validar fecha
    If objBase.FechaNacimiento = 0 Then
        err.Raise vbObjectError + 2025, "clsCalculoDesafio4", ERR_FECHA_INVALIDA
    End If
    
    ' El Desafío 4 = |Mes - Año| (valor absoluto de la diferencia)
    mes = Month(objBase.FechaNacimiento)
    ano = Year(objBase.FechaNacimiento)
    
    mesReducido = ReducirADigito(mes, False)
    anoReducido = ReducirADigito(ano, False)
    
    ' Diferencia absoluta
    objBase.Resultado = Abs(mesReducido - anoReducido)
    
    ' Calcular edades del desafío
    Set objDesafio3 = New clsCalculoDesafio3
    objDesafio3.FechaNacimiento = objBase.FechaNacimiento
    objDesafio3.Calcular
    
    pEdadInicio = objDesafio3.EdadFin + 1
    pEdadFin = 999  ' Hasta el final de la vida
    
    Calcular = objBase.Resultado
    
    ' Limpiar objetos
    Set objDesafio3 = Nothing
    
    Exit Function
    
ErrorHandler:
    Set objDesafio3 = Nothing
    err.Raise err.Number, "clsCalculoDesafio4.Calcular", err.Description
End Function

Public Function ObtenerInterpretacion() As String
    Dim nombreArchivo As String
    
    On Error GoTo ErrorHandler
    
    ' Verificar que se haya calculado
    If objBase.Resultado < 0 Then
        Calcular
    End If
    
    ' Construir nombre de archivo
    nombreArchivo = objBase.ObtenerNombreArchivoInterpretacion("Desafio4", objBase.Resultado)
    If objBase.Resultado = 0 Then
        nombreArchivo = "00_Desafio4"
    End If
    
    ' Cargar interpretación
    objBase.Interpretacion = objBase.CargarInterpretacionDesdeArchivo(nombreArchivo)
    
    ObtenerInterpretacion = objBase.Interpretacion
    Exit Function
    
ErrorHandler:
    err.Raise err.Number, "clsCalculoDesafio4.ObtenerInterpretacion", err.Description
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

