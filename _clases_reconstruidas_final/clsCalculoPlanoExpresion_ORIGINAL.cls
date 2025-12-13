VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCalculoPlanoExpresion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ============================================================================
' Clase: clsCalculoPlanoExpresion
' Descripción: Calcula los Planos de Expresión (Mental, Físico, Emocional, Intuitivo)
' Autor: Sistema de Numerología
' Fecha: 2024
' ============================================================================

Private objBase As clsCalculoBase
Private pPlanoMental As Integer
Private pPlanoFisico As Integer
Private pPlanoEmocional As Integer
Private pPlanoIntuitivo As Integer

' Definición de letras por plano
Private Const LETRAS_MENTAL As String = "ADGJMP"
Private Const LETRAS_FISICO As String = "EHKNQX"
Private Const LETRAS_EMOCIONAL As String = "BILORSTÑ"
Private Const LETRAS_INTUITIVO As String = "CFUWYZÇ"

' ============================================================================
' PROPIEDADES PÚBLICAS
' ============================================================================

Public Property Get NombreCompleto() As String
    NombreCompleto = objBase.NombreCompleto
End Property

Public Property Let NombreCompleto(ByVal valor As String)
    objBase.NombreCompleto = valor
End Property

Public Property Get PlanoMental() As Integer
    PlanoMental = pPlanoMental
End Property

Public Property Get PlanoFisico() As Integer
    PlanoFisico = pPlanoFisico
End Property

Public Property Get PlanoEmocional() As Integer
    PlanoEmocional = pPlanoEmocional
End Property

Public Property Get PlanoIntuitivo() As Integer
    PlanoIntuitivo = pPlanoIntuitivo
End Property

Public Property Get PlanoDominante() As String
    Dim maxValor As Integer
    Dim planoDom As String
    
    maxValor = pPlanoMental
    planoDom = "Mental"
    
    If pPlanoFisico > maxValor Then
        maxValor = pPlanoFisico
        planoDom = "Físico"
    End If
    
    If pPlanoEmocional > maxValor Then
        maxValor = pPlanoEmocional
        planoDom = "Emocional"
    End If
    
    If pPlanoIntuitivo > maxValor Then
        maxValor = pPlanoIntuitivo
        planoDom = "Intuitivo"
    End If
    
    PlanoDominante = planoDom
End Property

' ============================================================================
' MÉTODOS PÚBLICOS
' ============================================================================

Public Function Calcular() As Boolean
    Dim i As Integer
    Dim letra As String
    Dim nombreNormalizado As String
    
    On Error GoTo ErrorHandler
    
    ' Validar nombre
    If Len(objBase.NombreCompleto) = 0 Then
        err.Raise vbObjectError + 2034, "clsCalculoPlanoExpresion", ERR_NOMBRE_VACIO
    End If
    
    ' Inicializar contadores
    pPlanoMental = 0
    pPlanoFisico = 0
    pPlanoEmocional = 0
    pPlanoIntuitivo = 0
    
    nombreNormalizado = objBase.NombreCompleto
    
    ' Analizar cada letra
    For i = 1 To Len(nombreNormalizado)
        letra = Mid(nombreNormalizado, i, 1)
        
        ' Saltar espacios
        If letra = " " Then GoTo SiguienteLetra
        
        ' Saltar letras mudas
        If EsLetraMuda(letra, nombreNormalizado, i) Then GoTo SiguienteLetra
        
        ' Clasificar en el plano correspondiente
        If InStr(1, LETRAS_MENTAL, letra) > 0 Then
            pPlanoMental = pPlanoMental + 1
        ElseIf InStr(1, LETRAS_FISICO, letra) > 0 Then
            pPlanoFisico = pPlanoFisico + 1
        ElseIf InStr(1, LETRAS_EMOCIONAL, letra) > 0 Then
            pPlanoEmocional = pPlanoEmocional + 1
        ElseIf InStr(1, LETRAS_INTUITIVO, letra) > 0 Then
            pPlanoIntuitivo = pPlanoIntuitivo + 1
        End If
        
SiguienteLetra:
    Next i
    
    Calcular = True
    Exit Function
    
ErrorHandler:
    err.Raise err.Number, "clsCalculoPlanoExpresion.Calcular", err.Description
End Function

Public Function ObtenerInterpretacion() As String
    Dim Interpretacion As String
    
    On Error GoTo ErrorHandler
    
    ' Verificar que se haya calculado
    If pPlanoMental = 0 And pPlanoFisico = 0 And pPlanoEmocional = 0 And pPlanoIntuitivo = 0 Then
        Calcular
    End If
    
    ' Construir interpretación completa
    Interpretacion = "=== PLANOS DE EXPRESIÓN ===" & vbCrLf & vbCrLf
    
    Interpretacion = Interpretacion & "Plano Mental: " & pPlanoMental & " letras" & vbCrLf
    Interpretacion = Interpretacion & objBase.CargarInterpretacionDesdeArchivo("PlanoExpresion_MENTAL") & vbCrLf & vbCrLf
    
    Interpretacion = Interpretacion & "Plano Físico: " & pPlanoFisico & " letras" & vbCrLf
    Interpretacion = Interpretacion & objBase.CargarInterpretacionDesdeArchivo("PlanoExpresion_FISICO") & vbCrLf & vbCrLf
    
    Interpretacion = Interpretacion & "Plano Emocional: " & pPlanoEmocional & " letras" & vbCrLf
    Interpretacion = Interpretacion & objBase.CargarInterpretacionDesdeArchivo("PlanoExpresion_EMOCIONAL") & vbCrLf & vbCrLf
    
    Interpretacion = Interpretacion & "Plano Intuitivo: " & pPlanoIntuitivo & " letras" & vbCrLf
    Interpretacion = Interpretacion & objBase.CargarInterpretacionDesdeArchivo("PlanoExpresion_INTUITIVO") & vbCrLf & vbCrLf
    
    Interpretacion = Interpretacion & "PLANO DOMINANTE: " & PlanoDominante
    
    ObtenerInterpretacion = Interpretacion
    Exit Function
    
ErrorHandler:
    err.Raise err.Number, "clsCalculoPlanoExpresion.ObtenerInterpretacion", err.Description
End Function

Public Function ObtenerResumen() As String
    Dim resumen As String
    
    If pPlanoMental = 0 And pPlanoFisico = 0 And pPlanoEmocional = 0 And pPlanoIntuitivo = 0 Then
        Calcular
    End If
    
    resumen = "Mental: " & pPlanoMental & " | "
    resumen = resumen & "Físico: " & pPlanoFisico & " | "
    resumen = resumen & "Emocional: " & pPlanoEmocional & " | "
    resumen = resumen & "Intuitivo: " & pPlanoIntuitivo
    resumen = resumen & vbCrLf & "Dominante: " & PlanoDominante
    
    ObtenerResumen = resumen
End Function

' ============================================================================
' EVENTOS DE CLASE
' ============================================================================

Private Sub Class_Initialize()
    Set objBase = New clsCalculoBase
    pPlanoMental = 0
    pPlanoFisico = 0
    pPlanoEmocional = 0
    pPlanoIntuitivo = 0
End Sub

Private Sub Class_Terminate()
    Set objBase = Nothing
End Sub
