VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCalculoPrimeraConsonante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ============================================================================
' Clase: clsCalculoPrimeraConsonante
' Descripción: Calcula el Número de la Primera Consonante
' Autor: Sistema de Numerología
' Fecha: 2024
' ============================================================================

Private objBase As clsCalculoBase
Private pPrimeraConsonante As String

' ============================================================================
' PROPIEDADES PÚBLICAS
' ============================================================================

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

Public Property Get PrimeraConsonante() As String
    PrimeraConsonante = pPrimeraConsonante
End Property

' ============================================================================
' MÉTODOS PÚBLICOS
' ============================================================================

Public Function Calcular() As Integer
    Dim i As Integer
    Dim letra As String
    Dim nombreNormalizado As String
    
    On Error GoTo ErrorHandler
    
    ' Validar nombre
    If Len(objBase.NombreCompleto) = 0 Then
        Err.Raise vbObjectError + 2029, "clsCalculoPrimeraConsonante", ERR_NOMBRE_VACIO
    End If
    
    nombreNormalizado = objBase.NombreCompleto
    pPrimeraConsonante = ""
    
    ' Buscar la primera consonante del nombre
    For i = 1 To Len(nombreNormalizado)
        letra = Mid(nombreNormalizado, i, 1)
        
        ' Saltar espacios
        If letra = " " Then GoTo SiguienteLetra
        
        ' Saltar letras mudas
        If EsLetraMuda(letra, nombreNormalizado, i) Then GoTo SiguienteLetra
        
        ' Verificar si es consonante
        If Not EsVocal(letra, nombreNormalizado) Then
            pPrimeraConsonante = letra
            objBase.Resultado = ObtenerValorLetra(letra)
            Exit For
        End If
        
SiguienteLetra:
    Next i
    
    ' Verificar que se encontró una consonante
    If Len(pPrimeraConsonante) = 0 Then
        Err.Raise vbObjectError + 2030, "clsCalculoPrimeraConsonante", "No se encontró ninguna consonante en el nombre"
    End If
    
    Calcular = objBase.Resultado
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, "clsCalculoPrimeraConsonante.Calcular", Err.Description
End Function

Public Function ObtenerInterpretacion() As String
    Dim nombreArchivo As String
    
    On Error GoTo ErrorHandler
    
    ' Verificar que se haya calculado
    If objBase.Resultado = 0 Then
        Calcular
    End If
    
    ' Construir nombre de archivo usando la consonante
    nombreArchivo = pPrimeraConsonante & "_PrimeraConsonante"
    
    ' Cargar interpretación
    objBase.Interpretacion = objBase.CargarInterpretacionDesdeArchivo(nombreArchivo)
    
    ObtenerInterpretacion = objBase.Interpretacion
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, "clsCalculoPrimeraConsonante.ObtenerInterpretacion", Err.Description
End Function

' ============================================================================
' EVENTOS DE CLASE
' ============================================================================

Private Sub Class_Initialize()
    Set objBase = New clsCalculoBase
    pPrimeraConsonante = ""
End Sub

Private Sub Class_Terminate()
    Set objBase = Nothing
End Sub