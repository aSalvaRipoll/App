VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCalculoDesafio1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


' ============================================================================
' Clase: clsCalculoDesafio1
' Descripción: Calcula el Primer Desafío
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
    Dim dia As Integer
    Dim mes As Integer
    Dim diaReducido As Integer
    Dim mesReducido As Integer
    Dim objCaminoVida As clsCalculoCaminoVida
    Dim numeroCaminoVida As Integer
    
    On Error GoTo ErrorHandler
    
    ' Validar fecha
    If objBase.FechaNacimiento = 0 Then
        err.Raise vbObjectError + 2022, "clsCalculoDesafio1", ERR_FECHA_INVALIDA
    End If
    
    ' El Desafío 1 = |Mes - Día| (valor absoluto de la diferencia)
    dia = Day(objBase.FechaNacimiento)
    mes = Month(objBase.FechaNacimiento)
    
    diaReducido = ReducirADigito(dia, False)
    mesReducido = ReducirADigito(mes, False)
    
    ' Diferencia absoluta (los desafíos son siempre 0-8, sin maestros)
    objBase.Resultado = Abs(mesReducido - diaReducido)
    
    ' Calcular edades del desafío
    Set objCaminoVida = New clsCalculoCaminoVida
    objCaminoVida.FechaNacimiento = objBase.FechaNacimiento
    numeroCaminoVida = objCaminoVida.Calcular()
    
    pEdadInicio = 0
    pEdadFin = 36 - numeroCaminoVida
    
    Calcular = objBase.Resultado
    
    ' Limpiar objetos
    Set objCaminoVida = Nothing
    
    Exit Function
    
ErrorHandler:
    Set objCaminoVida = Nothing
    err.Raise err.Number, "clsCalculoDesafio1.Calcular", err.Description
End Function

Public Function ObtenerInterpretacion() As String
    Dim nombreArchivo As String
    
    On Error GoTo ErrorHandler
    
    ' Verificar que se haya calculado
    If objBase.Resultado < 0 Then
        Calcular
    End If
    
    ' Construir nombre de archivo (formato: 00_Desafio1.md para desafío 0)
    nombreArchivo = objBase.ObtenerNombreArchivoInterpretacion("Desafio1", objBase.Resultado)
    If objBase.Resultado = 0 Then
        nombreArchivo = "00_Desafio1"
    End If
    
    ' Cargar interpretación
    objBase.Interpretacion = objBase.CargarInterpretacionDesdeArchivo(nombreArchivo)
    
    ObtenerInterpretacion = objBase.Interpretacion
    Exit Function
    
ErrorHandler:
    err.Raise err.Number, "clsCalculoDesafio1.ObtenerInterpretacion", err.Description
End Function

Public Function ObtenerRangoEdades() As String
    If pEdadInicio = 0 And pEdadFin = 0 Then
        Calcular
    End If
    
    ObtenerRangoEdades = "Nacimiento - " & pEdadFin & " años"
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

