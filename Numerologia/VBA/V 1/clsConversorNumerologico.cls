VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsConversorNumerologico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ============================================================================
' Clase: clsConversorNumerologico
' Descripción: Clase para conversión entre sistema tradicional y fonético español
' Autor: Sistema de Numerología
' Fecha: 2024
' ============================================================================

Private pNombreTradicional As String
Private pNombreFonetico As String

' ============================================================================
' PROPIEDADES PÚBLICAS
' ============================================================================

Public Property Get NombreTradicional() As String
    NombreTradicional = pNombreTradicional
End Property

Public Property Let NombreTradicional(ByVal valor As String)
    pNombreTradicional = NormalizarTexto(valor)
End Property

Public Property Get NombreFonetico() As String
    NombreFonetico = pNombreFonetico
End Property

Public Property Let NombreFonetico(ByVal valor As String)
    pNombreFonetico = NormalizarTexto(valor)
End Property

' ============================================================================
' MÉTODOS PÚBLICOS
' ============================================================================

Public Function ConvertirTradicionalAFonetico(ByVal nombreTradicional As String) As String
    Dim resultado As String
    Dim i As Integer
    Dim letra As String
    Dim letraAnterior As String
    Dim letraSiguiente As String
    
    On Error GoTo ErrorHandler
    
    pNombreTradicional = NormalizarTexto(nombreTradicional)
    resultado = ""
    
    For i = 1 To Len(pNombreTradicional)
        letra = Mid(pNombreTradicional, i, 1)
        
        ' Obtener letras adyacentes
        If i > 1 Then letraAnterior = Mid(pNombreTradicional, i - 1, 1)
        If i < Len(pNombreTradicional) Then letraSiguiente = Mid(pNombreTradicional, i + 1, 1)
        
        ' Reglas de conversión fonética
        Select Case letra
            Case "H"
                ' H muda: no se incluye
                
            Case "U"
                ' U después de Q o G (antes de E/I) es muda
                If (letraAnterior = "Q") Or _
                   (letraAnterior = "G" And (letraSiguiente = "E" Or letraSiguiente = "I")) Then
                    ' No incluir U muda
                Else
                    resultado = resultado & letra
                End If
                
            Case "C"
                ' C antes de E o I suena como S
                If letraSiguiente = "E" Or letraSiguiente = "I" Then
                    resultado = resultado & "S"
                ' CH suena como una sola letra
                ElseIf letraSiguiente = "H" Then
                    resultado = resultado & "CH"
                Else
                    resultado = resultado & letra
                End If
                
            Case "G"
                ' G antes de E o I suena como J
                If letraSiguiente = "E" Or letraSiguiente = "I" Then
                    resultado = resultado & "J"
                Else
                    resultado = resultado & letra
                End If
                
            Case "Q"
                ' Q siempre va con U, suena como K
                resultado = resultado & "K"
                
            Case "Z"
                ' Z suena como S en español
                resultado = resultado & "S"
                
            Case "V"
                ' V suena como B en español
                resultado = resultado & "B"
                
            Case "LL"
                ' LL es una sola letra en español
                resultado = resultado & "Y"
                
            Case "X"
                ' X puede sonar como KS o S según contexto
                resultado = resultado & "KS"
                
            Case Else
                ' Mantener la letra
                resultado = resultado & letra
        End Select
        
        letraAnterior = letra
    Next i
    
    pNombreFonetico = resultado
    ConvertirTradicionalAFonetico = resultado
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, "clsConversorNumerologico.ConvertirTradicionalAFonetico", Err.Description
End Function

Public Function CompararCalculos(ByVal nombreTradicional As String) As String
    Dim objDestinoTradicional As clsCalculoDestino
    Dim objDestinoFonetico As clsCalculoDestino
    Dim nombreFonetico As String
    Dim resultado As String
    
    On Error GoTo ErrorHandler
    
    ' Convertir a fonético
    nombreFonetico = ConvertirTradicionalAFonetico(nombreTradicional)
    
    ' Calcular Destino tradicional
    Set objDestinoTradicional = New clsCalculoDestino
    objDestinoTradicional.NombreCompleto = nombreTradicional
    
    ' Calcular Destino fonético
    Set objDestinoFonetico = New clsCalculoDestino
    objDestinoFonetico.NombreCompleto = nombreFonetico
    
    ' Construir comparación
    resultado = "=== COMPARACIÓN DE SISTEMAS ===" & vbCrLf & vbCrLf
    resultado = resultado & "Nombre Tradicional: " & nombreTradicional & vbCrLf
    resultado = resultado & "Número de Destino: " & objDestinoTradicional.Calcular() & vbCrLf & vbCrLf
    resultado = resultado & "Nombre Fonético: " & nombreFonetico & vbCrLf
    resultado = resultado & "Número de Destino: " & objDestinoFonetico.Calcular() & vbCrLf & vbCrLf
    
    If objDestinoTradicional.Resultado = objDestinoFonetico.Resultado Then
        resultado = resultado & "Los dos sistemas producen el MISMO resultado."
    Else
        resultado = resultado & "Los dos sistemas producen DIFERENTES resultados."
    End If
    
    CompararCalculos = resultado
    
    ' Limpiar objetos
    Set objDestinoTradicional = Nothing
    Set objDestinoFonetico = Nothing
    
    Exit Function
    
ErrorHandler:
    Set objDestinoTradicional = Nothing
    Set objDestinoFonetico = Nothing
    Err.Raise Err.Number, "clsConversorNumerologico.CompararCalculos", Err.Description
End Function

' ============================================================================
' EVENTOS DE CLASE
' ============================================================================

Private Sub Class_Initialize()
    pNombreTradicional = ""
    pNombreFonetico = ""
End Sub

Private Sub Class_Terminate()
    ' Limpiar
End Sub