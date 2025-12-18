VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCalculoPrimeraVocal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


' ============================================================================
' Clase: clsCalculoPrimeraVocal
' Descripción: Calcula el Número de la Primera Vocal
' Autor: Sistema de Numerología
' Fecha: 2024
' ============================================================================

Private objBase As clsCalculoBase
Private pPrimeraVocal As String

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

Public Property Get PrimeraVocal() As String
    PrimeraVocal = pPrimeraVocal
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
        err.Raise vbObjectError + 2026, "clsCalculoPrimeraVocal", ERR_NOMBRE_VACIO
    End If
    
    nombreNormalizado = objBase.NombreCompleto
    pPrimeraVocal = ""
    
    ' Buscar la primera vocal del nombre
    For i = 1 To Len(nombreNormalizado)
        letra = Mid(nombreNormalizado, i, 1)
        
        ' Saltar espacios
        If letra = " " Then GoTo SiguienteLetra
        
        ' Verificar si es vocal
        If EsVocal(letra, nombreNormalizado) Then
            pPrimeraVocal = letra
            objBase.Resultado = ObtenerValorLetra(letra)
            Exit For
        End If
        
SiguienteLetra:
    Next i
    
    ' Verificar que se encontró una vocal
    If Len(pPrimeraVocal) = 0 Then
        err.Raise vbObjectError + 2027, "clsCalculoPrimeraVocal", "No se encontró ninguna vocal en el nombre"
    End If
    
    Calcular = objBase.Resultado
    Exit Function
    
ErrorHandler:
    err.Raise err.Number, "clsCalculoPrimeraVocal.Calcular", err.Description
End Function

Public Function ObtenerInterpretacion() As String
    Dim nombreArchivo As String
    
    On Error GoTo ErrorHandler
    
    ' Verificar que se haya calculado
    If objBase.Resultado = 0 Then
        Calcular
    End If
    
    ' Construir nombre de archivo usando la vocal
    nombreArchivo = pPrimeraVocal & "_PrimeraVocal"
    
    ' Cargar interpretación
    objBase.Interpretacion = objBase.CargarInterpretacionDesdeArchivo(nombreArchivo)
    
    ObtenerInterpretacion = objBase.Interpretacion
    Exit Function
    
ErrorHandler:
    err.Raise err.Number, "clsCalculoPrimeraVocal.ObtenerInterpretacion", err.Description
End Function

' ============================================================================
' EVENTOS DE CLASE
' ============================================================================

Private Sub Class_Initialize()
    Set objBase = New clsCalculoBase
    pPrimeraVocal = ""
End Sub

Private Sub Class_Terminate()
    Set objBase = Nothing
End Sub

