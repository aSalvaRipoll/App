VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCalculoDesafio2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ============================================================================
' Clase: clsCalculoDesafio2
' Descripción: Calcula el Segundo Desafío
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
    Dim ano As Integer
    Dim diaReducido As Integer
    Dim anoReducido As Integer
    Dim objCaminoVida As clsCalculoCaminoVida
    Dim objDesafio1 As clsCalculoDesafio1
    Dim numeroCaminoVida As Integer
    
    On Error GoTo ErrorHandler
    
    ' Validar fecha
    If objBase.FechaNacimiento = 0 Then
        Err.Raise vbObjectError + 2023, "clsCalculoDesafio2", ERR_FECHA_INVALIDA
    End If
    
    ' El Desafío 2 = |Día - Año| (valor absoluto de la diferencia)
    dia = Day(objBase.FechaNacimiento)
    ano = Year(objBase.FechaNacimiento)
    
    diaReducido = ReducirADigito(dia, False)
    anoReducido = ReducirADigito(ano, False)
    
    ' Diferencia absoluta
    objBase.Resultado = Abs(diaReducido - anoReducido)
    
    ' Calcular edades del desafío
    Set objCaminoVida = New clsCalculoCaminoVida
    objCaminoVida.FechaNacimiento = objBase.FechaNacimiento
    numeroCaminoVida = objCaminoVida.Calcular()
    
    Set objDesafio1 = New clsCalculoDesafio1
    objDesafio1.FechaNacimiento = objBase.FechaNacimiento
    objDesafio1.Calcular
    
    pEdadInicio = objDesafio1.EdadFin + 1
    pEdadFin = pEdadInicio + 8
    
    Calcular = objBase.Resultado
    
    ' Limpiar objetos
    Set objCaminoVida = Nothing
    Set objDesafio1 = Nothing
    
    Exit Function
    
ErrorHandler:
    Set objCaminoVida = Nothing
    Set objDesafio1 = Nothing
    Err.Raise Err.Number, "clsCalculoDesafio2.Calcular", Err.Description
End Function

Public Function ObtenerInterpretacion() As String
    Dim nombreArchivo As String
    
    On Error GoTo ErrorHandler
    
    ' Verificar que se haya calculado
    If objBase.Resultado < 0 Then
        Calcular
    End If
    
    ' Construir nombre de archivo
    nombreArchivo = objBase.ObtenerNombreArchivoInterpretacion("Desafio2", objBase.Resultado)
    If objBase.Resultado = 0 Then
        nombreArchivo = "00_Desafio2"
    End If
    
    ' Cargar interpretación
    objBase.Interpretacion = objBase.CargarInterpretacionDesdeArchivo(nombreArchivo)
    
    ObtenerInterpretacion = objBase.Interpretacion
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, "clsCalculoDesafio2.ObtenerInterpretacion", Err.Description
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