VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCalculoNumeroDominante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ============================================================================
' Clase: clsCalculoNumeroDominante
' Descripción: Calcula el Número Dominante (más repetido en el nombre)
' Autor: Sistema de Numerología
' Fecha: 2024
' ============================================================================

Private objBase As clsCalculoBase
Private pFrecuencias() As Integer  ' Array del 1 al 9
Private pMaximaFrecuencia As Integer

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

Public Property Get MaximaFrecuencia() As Integer
    MaximaFrecuencia = pMaximaFrecuencia
End Property

' ============================================================================
' MÉTODOS PÚBLICOS
' ============================================================================

Public Function Calcular() As Integer
    Dim i As Integer
    Dim letra As String
    Dim valor As Integer
    Dim nombreNormalizado As String
    
    On Error GoTo ErrorHandler
    
    ' Validar nombre
    If Len(objBase.NombreCompleto) = 0 Then
        Err.Raise vbObjectError + 2032, "clsCalculoNumeroDominante", ERR_NOMBRE_VACIO
    End If
    
    ' Inicializar array de frecuencias
    ReDim pFrecuencias(1 To 9)
    For i = 1 To 9
        pFrecuencias(i) = 0
    Next i
    
    nombreNormalizado = objBase.NombreCompleto
    
    ' Contar frecuencia de cada número
    For i = 1 To Len(nombreNormalizado)
        letra = Mid(nombreNormalizado, i, 1)
        
        ' Saltar espacios
        If letra = " " Then GoTo SiguienteLetra
        
        ' Saltar letras mudas
        If EsLetraMuda(letra, nombreNormalizado, i) Then GoTo SiguienteLetra
        
        ' Obtener valor
        valor = ObtenerValorLetra(letra)
        
        ' Incrementar contador
        If valor >= 1 And valor <= 9 Then
            pFrecuencias(valor) = pFrecuencias(valor) + 1
        End If
        
SiguienteLetra:
    Next i
    
    ' Encontrar el número más frecuente
    pMaximaFrecuencia = 0
    objBase.Resultado = 0
    
    For i = 1 To 9
        If pFrecuencias(i) > pMaximaFrecuencia Then
            pMaximaFrecuencia = pFrecuencias(i)
            objBase.Resultado = i
        End If
    Next i
    
    Calcular = objBase.Resultado
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, "clsCalculoNumeroDominante.Calcular", Err.Description
End Function

Public Function ObtenerInterpretacion() As String
    Dim nombreArchivo As String
    
    On Error GoTo ErrorHandler
    
    ' Verificar que se haya calculado
    If objBase.Resultado = 0 Then
        Calcular
    End If
    
    ' Construir nombre de archivo
    nombreArchivo = objBase.ObtenerNombreArchivoInterpretacion("Dominante", objBase.Resultado)
    
    ' Cargar interpretación
    objBase.Interpretacion = objBase.CargarInterpretacionDesdeArchivo(nombreArchivo)
    
    ObtenerInterpretacion = objBase.Interpretacion
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, "clsCalculoNumeroDominante.ObtenerInterpretacion", Err.Description
End Function

Public Function ObtenerFrecuencia(ByVal numero As Integer) As Integer
    If numero < 1 Or numero > 9 Then
        ObtenerFrecuencia = 0
    Else
        ObtenerFrecuencia = pFrecuencias(numero)
    End If
End Function

' ============================================================================
' EVENTOS DE CLASE
' ============================================================================

Private Sub Class_Initialize()
    Set objBase = New clsCalculoBase
    ReDim pFrecuencias(1 To 9)
    pMaximaFrecuencia = 0
End Sub

Private Sub Class_Terminate()
    Set objBase = Nothing
    Erase pFrecuencias
End Sub