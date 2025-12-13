VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCalculoEdadPersonal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ============================================================================
' Clase: clsCalculoEdadPersonal
' Descripción: Calcula el Número de Edad Personal
' Autor: Sistema de Numerología
' Fecha: 2024
' ============================================================================

Private objBase As clsCalculoBase
Private pEdadActual As Integer

' ============================================================================
' PROPIEDADES PÚBLICAS
' ============================================================================

Public Property Get FechaNacimiento() As Date
    FechaNacimiento = objBase.FechaNacimiento
End Property

Public Property Let FechaNacimiento(ByVal valor As Date)
    objBase.FechaNacimiento = valor
End Property

Public Property Get EdadActual() As Integer
    EdadActual = pEdadActual
End Property

Public Property Get Resultado() As Integer
    Resultado = objBase.Resultado
End Property

Public Property Get Interpretacion() As String
    Interpretacion = objBase.Interpretacion
End Property

' ============================================================================
' MÉTODOS PÚBLICOS
' ============================================================================

Public Function Calcular() As Integer
    Dim fechaHoy As Date
    Dim anos As Integer
    
    On Error GoTo ErrorHandler
    
    ' Validar fecha
    If objBase.FechaNacimiento = 0 Then
        Err.Raise vbObjectError + 2035, "clsCalculoEdadPersonal", ERR_FECHA_INVALIDA
    End If
    
    fechaHoy = Date
    
    ' Calcular edad actual
    anos = Year(fechaHoy) - Year(objBase.FechaNacimiento)
    
    ' Ajustar si aún no ha cumplido años este año
    If Month(fechaHoy) < Month(objBase.FechaNacimiento) Or _
       (Month(fechaHoy) = Month(objBase.FechaNacimiento) And Day(fechaHoy) < Day(objBase.FechaNacimiento)) Then
        anos = anos - 1
    End If
    
    pEdadActual = anos
    
    ' Reducir edad a número (permitiendo maestros)
    objBase.Resultado = ReducirADigito(anos, True)
    
    Calcular = objBase.Resultado
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, "clsCalculoEdadPersonal.Calcular", Err.Description
End Function

Public Function ObtenerInterpretacion() As String
    Dim nombreArchivo As String
    
    On Error GoTo ErrorHandler
    
    ' Verificar que se haya calculado
    If objBase.Resultado = 0 Then
        Calcular
    End If
    
    ' Construir nombre de archivo
    nombreArchivo = objBase.ObtenerNombreArchivoInterpretacion("EdadPersonal", objBase.Resultado)
    
    ' Cargar interpretación
    objBase.Interpretacion = objBase.CargarInterpretacionDesdeArchivo(nombreArchivo)
    
    ObtenerInterpretacion = objBase.Interpretacion
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, "clsCalculoEdadPersonal.ObtenerInterpretacion", Err.Description
End Function

Public Function ObtenerInformacionCompleta() As String
    Dim info As String
    
    If pEdadActual = 0 Then
        Calcular
    End If
    
    info = "Edad actual: " & pEdadActual & " años" & vbCrLf
    info = info & "Número de Edad Personal: " & FormatearNumeroResultado(objBase.Resultado, True)
    
    ObtenerInformacionCompleta = info
End Function

' ============================================================================
' EVENTOS DE CLASE
' ============================================================================

Private Sub Class_Initialize()
    Set objBase = New clsCalculoBase
    pEdadActual = 0
End Sub

Private Sub Class_Terminate()
    Set objBase = Nothing
End Sub