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

Public Property Get nombreTradicional() As String
    nombreTradicional = pNombreTradicional
End Property

Public Property Let nombreTradicional(ByVal valor As String)
    pNombreTradicional = NormalizarTexto(valor)
End Property

Public Property Get nombreFonetico() As String
    nombreFonetico = pNombreFonetico
End Property

Public Property Let nombreFonetico(ByVal valor As String)
    pNombreFonetico = NormalizarTexto(valor)
End Property

' ============================================================================
' MÉTODOS PÚBLICOS
' ============================================================================

Public Function ConvertirTradicionalAFonetico(ByVal nombreTradicional As String) As String
    Dim Resultado As String
    Dim i As Integer
    Dim letra As String
    Dim letraAnterior As String
    Dim letraSiguiente As String
    
    On Error GoTo ErrorHandler
    
    pNombreTradicional = NormalizarTexto(nombreTradicional)
    Resultado = ""
    
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
                    Resultado = Resultado & letra
                End If
                
            Case "C"
                ' C antes de E o I suena como S
                If letraSiguiente = "E" Or letraSiguiente = "I" Then
                    Resultado = Resultado & "S"
                ' CH suena como una sola letra
                ElseIf letraSiguiente = "H" Then
                    Resultado = Resultado & "CH"
                Else
                    Resultado = Resultado & letra
                End If
                
            Case "G"
                ' G antes de E o I suena como J
                If letraSiguiente = "E" Or letraSiguiente = "I" Then
                    Resultado = Resultado & "J"
                Else
                    Resultado = Resultado & letra
                End If
                
            Case "Q"
                ' Q siempre va con U, suena como K
                Resultado = Resultado & "K"
                
            Case "Z"
                ' Z suena como S en español
                Resultado = Resultado & "S"
                
            Case "V"
                ' V suena como B en español
                Resultado = Resultado & "B"
                
            Case "LL"
                ' LL es una sola letra en español
                Resultado = Resultado & "Y"
                
            Case "X"
                ' X puede sonar como KS o S según contexto
                Resultado = Resultado & "KS"
                
            Case Else
                ' Mantener la letra
                Resultado = Resultado & letra
        End Select
        
        letraAnterior = letra
    Next i
    
    pNombreFonetico = Resultado
    ConvertirTradicionalAFonetico = Resultado
    Exit Function
    
ErrorHandler:
    err.Raise err.Number, "clsConversorNumerologico.ConvertirTradicionalAFonetico", err.Description
End Function

Public Function CompararCalculos(ByVal nombreTradicional As String) As String
    Dim objDestinoTradicional As clsCalculoDestino
    Dim objDestinoFonetico As clsCalculoDestino
    Dim nombreFonetico As String
    Dim Resultado As String
    
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
    Resultado = "=== COMPARACIÓN DE SISTEMAS ===" & vbCrLf & vbCrLf
    Resultado = Resultado & "Nombre Tradicional: " & nombreTradicional & vbCrLf
    Resultado = Resultado & "Número de Destino: " & objDestinoTradicional.Calcular() & vbCrLf & vbCrLf
    Resultado = Resultado & "Nombre Fonético: " & nombreFonetico & vbCrLf
    Resultado = Resultado & "Número de Destino: " & objDestinoFonetico.Calcular() & vbCrLf & vbCrLf
    
    If objDestinoTradicional.Resultado = objDestinoFonetico.Resultado Then
        Resultado = Resultado & "Los dos sistemas producen el MISMO resultado."
    Else
        Resultado = Resultado & "Los dos sistemas producen DIFERENTES resultados."
    End If
    
    CompararCalculos = Resultado
    
    ' Limpiar objetos
    Set objDestinoTradicional = Nothing
    Set objDestinoFonetico = Nothing
    
    Exit Function
    
ErrorHandler:
    Set objDestinoTradicional = Nothing
    Set objDestinoFonetico = Nothing
    err.Raise err.Number, "clsConversorNumerologico.CompararCalculos", err.Description
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

