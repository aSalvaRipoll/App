VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCalculoCaminoVida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ============================================================================
' Proyecto:     Sistema de Numerología Tradicional y Fonético
' Clase: clsCalculoCaminoVida
' Descripción: Calcula el Número de Camino de Vida
' Autor: Sistema de Numerología
' Fecha: 2024
' ============================================================================

' Hereda de clsCalculoBase (conceptualmente)
Implements clsCalculoBase

Private objBase As clsCalculoBase

' ============================================================================
' IMPLEMENTACIÓN DE PROPIEDADES BASE
' ============================================================================

Private Property Get clsCalculoBase_FechaNacimiento() As Date
    clsCalculoBase_FechaNacimiento = objBase.FechaNacimiento
End Property

Private Property Let clsCalculoBase_FechaNacimiento(ByVal valor As Date)
    objBase.FechaNacimiento = valor
End Property

Private Property Get clsCalculoBase_NombreCompleto() As String
    clsCalculoBase_NombreCompleto = objBase.NombreCompleto
End Property

Private Property Let clsCalculoBase_NombreCompleto(ByVal valor As String)
    objBase.NombreCompleto = valor
End Property

Private Property Get clsCalculoBase_Resultado() As Integer
    clsCalculoBase_Resultado = objBase.Resultado
End Property

Private Property Get clsCalculoBase_Interpretacion() As String
    clsCalculoBase_Interpretacion = objBase.Interpretacion
End Property

Private Property Get clsCalculoBase_EsKarmico() As Boolean
    clsCalculoBase_EsKarmico = objBase.EsKarmico
End Property

Private Property Get clsCalculoBase_EsMaestro() As Boolean
    clsCalculoBase_EsMaestro = objBase.EsMaestro
End Property

' ============================================================================
' PROPIEDADES PÚBLICAS ESPECÍFICAS
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

' ============================================================================
' MÉTODOS PÚBLICOS
' ============================================================================

Public Function Calcular() As Integer
    Dim dia As Integer
    Dim mes As Integer
    Dim ano As Integer
    Dim sumaDia As Integer
    Dim sumaMes As Integer
    Dim sumaAno As Integer
    Dim sumaTotal As Long
    
    On Error GoTo ErrorHandler
    
    ' Validar fecha
    If objBase.FechaNacimiento = 0 Then
        err.Raise vbObjectError + 2001, "clsCalculoCaminoVida", ERR_FECHA_INVALIDA
    End If
    
    ' Obtener componentes de la fecha
    dia = Day(objBase.FechaNacimiento)
    mes = Month(objBase.FechaNacimiento)
    ano = Year(objBase.FechaNacimiento)
    
    ' Reducir cada componente preservando maestros
    sumaDia = ReducirADigito(dia, True)
    sumaMes = ReducirADigito(mes, True)
    sumaAno = ReducirADigito(ano, True)
    
    ' Sumar componentes reducidos
    sumaTotal = CLng(sumaDia) + CLng(sumaMes) + CLng(sumaAno)
    
    ' Reducir suma total
    objBase.Resultado = ReducirADigito(sumaTotal, True)
    
    Calcular = objBase.Resultado
    Exit Function
    
ErrorHandler:
    err.Raise err.Number, "clsCalculoCaminoVida.Calcular", err.Description
End Function

Public Function ObtenerInterpretacion() As String
    Dim nombreArchivo As String
    
    On Error GoTo ErrorHandler
    
    ' Verificar que se haya calculado
    If objBase.Resultado = 0 Then
        Calcular
    End If
    
    ' Construir nombre de archivo
    nombreArchivo = objBase.ObtenerNombreArchivoInterpretacion("CaminoVida", objBase.Resultado)
    
    ' Cargar interpretación
    objBase.Interpretacion = objBase.CargarInterpretacionDesdeArchivo(nombreArchivo)
    
    ObtenerInterpretacion = objBase.Interpretacion
    Exit Function
    
ErrorHandler:
    err.Raise err.Number, "clsCalculoCaminoVida.ObtenerInterpretacion", err.Description
End Function

' ============================================================================
' IMPLEMENTACIÓN DE MÉTODOS BASE
' ============================================================================

Private Function clsCalculoBase_Calcular() As Integer
    clsCalculoBase_Calcular = Calcular()
End Function

Private Function clsCalculoBase_ObtenerInterpretacion() As String
    clsCalculoBase_ObtenerInterpretacion = ObtenerInterpretacion()
End Function

Private Function clsCalculoBase_ObtenerInformacionCompleta() As String
    clsCalculoBase_ObtenerInformacionCompleta = objBase.ObtenerInformacionCompleta()
End Function

' ============================================================================
' EVENTOS DE CLASE
' ============================================================================

Private Sub Class_Initialize()
    Set objBase = New clsCalculoBase
End Sub

Private Sub Class_Terminate()
    Set objBase = Nothing
End Sub
