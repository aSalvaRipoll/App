VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCalculoRespuestaSubconsciente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


' ============================================================================
' Clase: clsCalculoRespuestaSubconsciente
' Descripción: Calcula la Respuesta Subconsciente (cantidad de números presentes)
' Autor: Sistema de Numerología
' Fecha: 2024
' ============================================================================

Private objBase As clsCalculoBase
Private pNumerosPresentes() As Boolean  ' Array del 1 al 9

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

' ============================================================================
' MÉTODOS PÚBLICOS
' ============================================================================

Public Function Calcular() As Integer
    Dim i As Integer
    Dim letra As String
    Dim valor As Integer
    Dim nombreNormalizado As String
    Dim contador As Integer
    
    On Error GoTo ErrorHandler
    
    ' Validar nombre
    If Len(objBase.NombreCompleto) = 0 Then
        err.Raise vbObjectError + 2031, "clsCalculoRespuestaSubconsciente", ERR_NOMBRE_VACIO
    End If
    
    ' Inicializar array
    ReDim pNumerosPresentes(1 To 9)
    For i = 1 To 9
        pNumerosPresentes(i) = False
    Next i
    
    nombreNormalizado = objBase.NombreCompleto
    
    ' Analizar cada letra del nombre
    For i = 1 To Len(nombreNormalizado)
        letra = Mid(nombreNormalizado, i, 1)
        
        ' Saltar espacios
        If letra = " " Then GoTo SiguienteLetra
        
        ' Saltar letras mudas
        If EsLetraMuda(letra, nombreNormalizado, i) Then GoTo SiguienteLetra
        
        ' Obtener valor
        valor = ObtenerValorLetra(letra)
        
        ' Marcar número como presente
        If valor >= 1 And valor <= 9 Then
            pNumerosPresentes(valor) = True
        End If
        
SiguienteLetra:
    Next i
    
    ' Contar cuántos números diferentes están presentes
    contador = 0
    For i = 1 To 9
        If pNumerosPresentes(i) Then
            contador = contador + 1
        End If
    Next i
    
    objBase.Resultado = contador
    Calcular = objBase.Resultado
    Exit Function
    
ErrorHandler:
    err.Raise err.Number, "clsCalculoRespuestaSubconsciente.Calcular", err.Description
End Function

Public Function ObtenerInterpretacion() As String
    Dim nombreArchivo As String
    
    On Error GoTo ErrorHandler
    
    ' Verificar que se haya calculado
    If objBase.Resultado = 0 Then
        Calcular
    End If
    
    ' Construir nombre de archivo
    nombreArchivo = objBase.ObtenerNombreArchivoInterpretacion("RespuestaSubconsciente", objBase.Resultado)
    
    ' Cargar interpretación
    objBase.Interpretacion = objBase.CargarInterpretacionDesdeArchivo(nombreArchivo)
    
    ObtenerInterpretacion = objBase.Interpretacion
    Exit Function
    
ErrorHandler:
    err.Raise err.Number, "clsCalculoRespuestaSubconsciente.ObtenerInterpretacion", err.Description
End Function

Public Function NumeroEstaPresente(ByVal numero As Integer) As Boolean
    If numero < 1 Or numero > 9 Then
        NumeroEstaPresente = False
    Else
        NumeroEstaPresente = pNumerosPresentes(numero)
    End If
End Function

' ============================================================================
' EVENTOS DE CLASE
' ============================================================================

Private Sub Class_Initialize()
    Set objBase = New clsCalculoBase
    ReDim pNumerosPresentes(1 To 9)
End Sub

Private Sub Class_Terminate()
    Set objBase = Nothing
    Erase pNumerosPresentes
End Sub

