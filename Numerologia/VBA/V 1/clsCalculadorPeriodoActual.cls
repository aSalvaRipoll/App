VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCalculadorPeriodoActual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ============================================================================
' Clase: clsCalculadorPeriodoActual
' Descripción: Determina el Ciclo, Pináculo y Desafío actual según la edad
' Autor: Sistema de Numerología
' Fecha: 2024
' ============================================================================

Private pFechaNacimiento As Date
Private pEdadActual As Integer

' ============================================================================
' PROPIEDADES PÚBLICAS
' ============================================================================

Public Property Get FechaNacimiento() As Date
    FechaNacimiento = pFechaNacimiento
End Property

Public Property Let FechaNacimiento(ByVal valor As Date)
    If Not ValidarFecha(valor) Then
        Err.Raise vbObjectError + 3001, "clsCalculadorPeriodoActual", ERR_FECHA_INVALIDA
    End If
    pFechaNacimiento = valor
    pEdadActual = CalcularEdadActual()
End Property

Public Property Get EdadActual() As Integer
    EdadActual = pEdadActual
End Property

' ============================================================================
' MÉTODOS PÚBLICOS - CICLOS
' ============================================================================

Public Function ObtenerCicloActual() As Integer
    Dim objCiclo1 As clsCalculoCiclo1
    Dim objCiclo2 As clsCalculoCiclo2
    Dim objCiclo3 As clsCalculoCiclo3
    
    On Error GoTo ErrorHandler
    
    If pFechaNacimiento = 0 Then
        Err.Raise vbObjectError + 3002, "clsCalculadorPeriodoActual", ERR_FECHA_INVALIDA
    End If
    
    ' Verificar Ciclo 1
    Set objCiclo1 = New clsCalculoCiclo1
    objCiclo1.FechaNacimiento = pFechaNacimiento
    objCiclo1.Calcular
    
    If pEdadActual <= objCiclo1.EdadFin Then
        ObtenerCicloActual = 1
        Set objCiclo1 = Nothing
        Exit Function
    End If
    
    ' Verificar Ciclo 2
    Set objCiclo2 = New clsCalculoCiclo2
    objCiclo2.FechaNacimiento = pFechaNacimiento
    objCiclo2.Calcular
    
    If pEdadActual >= objCiclo2.EdadInicio And pEdadActual <= objCiclo2.EdadFin Then
        ObtenerCicloActual = 2
        Set objCiclo1 = Nothing
        Set objCiclo2 = Nothing
        Exit Function
    End If
    
    ' Si no está en Ciclo 1 ni 2, está en Ciclo 3
    ObtenerCicloActual = 3
    
    Set objCiclo1 = Nothing
    Set objCiclo2 = Nothing
    Set objCiclo3 = Nothing
    
    Exit Function
    
ErrorHandler:
    Set objCiclo1 = Nothing
    Set objCiclo2 = Nothing
    Set objCiclo3 = Nothing
    Err.Raise Err.Number, "clsCalculadorPeriodoActual.ObtenerCicloActual", Err.Description
End Function

Public Function ObtenerInformacionCicloActual() As String
    Dim numeroCiclo As Integer
    Dim objCiclo1 As clsCalculoCiclo1
    Dim objCiclo2 As clsCalculoCiclo2
    Dim objCiclo3 As clsCalculoCiclo3
    Dim valorCiclo As Integer
    Dim rangoEdades As String
    Dim info As String
    
    On Error GoTo ErrorHandler
    
    numeroCiclo = ObtenerCicloActual()
    
    Select Case numeroCiclo
        Case 1
            Set objCiclo1 = New clsCalculoCiclo1
            objCiclo1.FechaNacimiento = pFechaNacimiento
            valorCiclo = objCiclo1.Calcular()
            rangoEdades = objCiclo1.ObtenerRangoEdades()
            
        Case 2
            Set objCiclo2 = New clsCalculoCiclo2
            objCiclo2.FechaNacimiento = pFechaNacimiento
            valorCiclo = objCiclo2.Calcular()
            rangoEdades = objCiclo2.ObtenerRangoEdades()
            
        Case 3
            Set objCiclo3 = New clsCalculoCiclo3
            objCiclo3.FechaNacimiento = pFechaNacimiento
            valorCiclo = objCiclo3.Calcular()
            rangoEdades = objCiclo3.ObtenerRangoEdades()
    End Select
    
    info = "=== CICLO ACTUAL ===" & vbCrLf
    info = info & "Edad actual: " & pEdadActual & " años" & vbCrLf
    info = info & "Ciclo activo: Ciclo " & numeroCiclo & vbCrLf
    info = info & "Valor del ciclo: " & FormatearNumeroResultado(valorCiclo, True) & vbCrLf
    info = info & "Rango de edades: " & rangoEdades
    
    ObtenerInformacionCicloActual = info
    
    Set objCiclo1 = Nothing
    Set objCiclo2 = Nothing
    Set objCiclo3 = Nothing
    
    Exit Function
    
ErrorHandler:
    Set objCiclo1 = Nothing
    Set objCiclo2 = Nothing
    Set objCiclo3 = Nothing
    Err.Raise Err.Number, "clsCalculadorPeriodoActual.ObtenerInformacionCicloActual", Err.Description
End Function

' ============================================================================
' MÉTODOS PÚBLICOS - PINÁCULOS
' ============================================================================

Public Function ObtenerPinaculoActual() As Integer
    Dim objPinaculo1 As clsCalculoPinaculo1
    Dim objPinaculo2 As clsCalculoPinaculo2
    Dim objPinaculo3 As clsCalculoPinaculo3
    
    On Error GoTo ErrorHandler
    
    If pFechaNacimiento = 0 Then
        Err.Raise vbObjectError + 3003, "clsCalculadorPeriodoActual", ERR_FECHA_INVALIDA
    End If
    
    ' Verificar Pináculo 1
    Set objPinaculo1 = New clsCalculoPinaculo1
    objPinaculo1.FechaNacimiento = pFechaNacimiento
    objPinaculo1.Calcular
    
    If pEdadActual <= objPinaculo1.EdadFin Then
        ObtenerPinaculoActual = 1
        Set objPinaculo1 = Nothing
        Exit Function
    End If
    
    ' Verificar Pináculo 2
    Set objPinaculo2 = New clsCalculoPinaculo2
    objPinaculo2.FechaNacimiento = pFechaNacimiento
    objPinaculo2.Calcular
    
    If pEdadActual >= objPinaculo2.EdadInicio And pEdadActual <= objPinaculo2.EdadFin Then
        ObtenerPinaculoActual = 2
        Set objPinaculo1 = Nothing
        Set objPinaculo2 = Nothing
        Exit Function
    End If
    
    ' Verificar Pináculo 3
    Set objPinaculo3 = New clsCalculoPinaculo3
    objPinaculo3.FechaNacimiento = pFechaNacimiento
    objPinaculo3.Calcular
    
    If pEdadActual >= objPinaculo3.EdadInicio And pEdadActual <= objPinaculo3.EdadFin Then
        ObtenerPinaculoActual = 3
        Set objPinaculo1 = Nothing
        Set objPinaculo2 = Nothing
        Set objPinaculo3 = Nothing
        Exit Function
    End If
    
    ' Si no está en 1, 2 ni 3, está en Pináculo 4
    ObtenerPinaculoActual = 4
    
    Set objPinaculo1 = Nothing
    Set objPinaculo2 = Nothing
    Set objPinaculo3 = Nothing
    
    Exit Function
    
ErrorHandler:
    Set objPinaculo1 = Nothing
    Set objPinaculo2 = Nothing
    Set objPinaculo3 = Nothing
    Err.Raise Err.Number, "clsCalculadorPeriodoActual.ObtenerPinaculoActual", Err.Description
End Function

Public Function ObtenerInformacionPinaculoActual() As String
    Dim numeroPinaculo As Integer
    Dim objPinaculo1 As clsCalculoPinaculo1
    Dim objPinaculo2 As clsCalculoPinaculo2
    Dim objPinaculo3 As clsCalculoPinaculo3
    Dim objPinaculo4 As clsCalculoPinaculo4
    Dim valorPinaculo As Integer
    Dim rangoEdades As String
    Dim info As String
    
    On Error GoTo ErrorHandler
    
    numeroPinaculo = ObtenerPinaculoActual()
    
    Select Case numeroPinaculo
        Case 1
            Set objPinaculo1 = New clsCalculoPinaculo1
            objPinaculo1.FechaNacimiento = pFechaNacimiento
            valorPinaculo = objPinaculo1.Calcular()
            rangoEdades = objPinaculo1.ObtenerRangoEdades()
            
        Case 2
            Set objPinaculo2 = New clsCalculoPinaculo2
            objPinaculo2.FechaNacimiento = pFechaNacimiento
            valorPinaculo = objPinaculo2.Calcular()
            rangoEdades = objPinaculo2.ObtenerRangoEdades()
            
        Case 3
            Set objPinaculo3 = New clsCalculoPinaculo3
            objPinaculo3.FechaNacimiento = pFechaNacimiento
            valorPinaculo = objPinaculo3.Calcular()
            rangoEdades = objPinaculo3.ObtenerRangoEdades()
            
        Case 4
            Set objPinaculo4 = New clsCalculoPinaculo4
            objPinaculo4.FechaNacimiento = pFechaNacimiento
            valorPinaculo = objPinaculo4.Calcular()
            rangoEdades = objPinaculo4.ObtenerRangoEdades()
    End Select
    
    info = "=== PINÁCULO ACTUAL ===" & vbCrLf
    info = info & "Edad actual: " & pEdadActual & " años" & vbCrLf
    info = info & "Pináculo activo: Pináculo " & numeroPinaculo & vbCrLf
    info = info & "Valor del pináculo: " & FormatearNumeroResultado(valorPinaculo, True) & vbCrLf
    info = info & "Rango de edades: " & rangoEdades
    
    ObtenerInformacionPinaculoActual = info
    
    Set objPinaculo1 = Nothing
    Set objPinaculo2 = Nothing
    Set objPinaculo3 = Nothing
    Set objPinaculo4 = Nothing
    
    Exit Function
    
ErrorHandler:
    Set objPinaculo1 = Nothing
    Set objPinaculo2 = Nothing
    Set objPinaculo3 = Nothing
    Set objPinaculo4 = Nothing
    Err.Raise Err.Number, "clsCalculadorPeriodoActual.ObtenerInformacionPinaculoActual", Err.Description
End Function

' ============================================================================
' MÉTODOS PÚBLICOS - DESAFÍOS
' ============================================================================

Public Function ObtenerDesafioActual() As Integer
    Dim objDesafio1 As clsCalculoDesafio1
    Dim objDesafio2 As clsCalculoDesafio2
    Dim objDesafio3 As clsCalculoDesafio3
    
    On Error GoTo ErrorHandler
    
    If pFechaNacimiento = 0 Then
        Err.Raise vbObjectError + 3004, "clsCalculadorPeriodoActual", ERR_FECHA_INVALIDA
    End If
    
    ' Verificar Desafío 1
    Set objDesafio1 = New clsCalculoDesafio1
    objDesafio1.FechaNacimiento = pFechaNacimiento
    objDesafio1.Calcular
    
    If pEdadActual <= objDesafio1.EdadFin Then
        ObtenerDesafioActual = 1
        Set objDesafio1 = Nothing
        Exit Function
    End If
    
    ' Verificar Desafío 2
    Set objDesafio2 = New clsCalculoDesafio2
    objDesafio2.FechaNacimiento = pFechaNacimiento
    objDesafio2.Calcular
    
    If pEdadActual >= objDesafio2.EdadInicio And pEdadActual <= objDesafio2.EdadFin Then
        ObtenerDesafioActual = 2
        Set objDesafio1 = Nothing
        Set objDesafio2 = Nothing
        Exit Function
    End If
    
    ' Verificar Desafío 3
    Set objDesafio3 = New clsCalculoDesafio3
    objDesafio3.FechaNacimiento = pFechaNacimiento
    objDesafio3.Calcular
    
    If pEdadActual >= objDesafio3.EdadInicio And pEdadActual <= objDesafio3.EdadFin Then
        ObtenerDesafioActual = 3
        Set objDesafio1 = Nothing
        Set objDesafio2 = Nothing
        Set objDesafio3 = Nothing
        Exit Function
    End If
    
    ' Si no está en 1, 2 ni 3, está en Desafío 4
    ObtenerDesafioActual = 4
    
    Set objDesafio1 = Nothing
    Set objDesafio2 = Nothing
    Set objDesafio3 = Nothing
    
    Exit Function
    
ErrorHandler:
    Set objDesafio1 = Nothing
    Set objDesafio2 = Nothing
    Set objDesafio3 = Nothing
    Err.Raise Err.Number, "clsCalculadorPeriodoActual.ObtenerDesafioActual", Err.Description
End Function

Public Function ObtenerInformacionDesafioActual() As String
    Dim numeroDesafio As Integer
    Dim objDesafio1 As clsCalculoDesafio1
    Dim objDesafio2 As clsCalculoDesafio2
    Dim objDesafio3 As clsCalculoDesafio3
    Dim objDesafio4 As clsCalculoDesafio4
    Dim valorDesafio As Integer
    Dim rangoEdades As String
    Dim info As String
    
    On Error GoTo ErrorHandler
    
    numeroDesafio = ObtenerDesafioActual()
    
    Select Case numeroDesafio
        Case 1
            Set objDesafio1 = New clsCalculoDesafio1
            objDesafio1.FechaNacimiento = pFechaNacimiento
            valorDesafio = objDesafio1.Calcular()
            rangoEdades = objDesafio1.ObtenerRangoEdades()
            
        Case 2
            Set objDesafio2 = New clsCalculoDesafio2
            objDesafio2.FechaNacimiento = pFechaNacimiento
            valorDesafio = objDesafio2.Calcular()
            rangoEdades = objDesafio2.ObtenerRangoEdades()
            
        Case 3
            Set objDesafio3 = New clsCalculoDesafio3
            objDesafio3.FechaNacimiento = pFechaNacimiento
            valorDesafio = objDesafio3.Calcular()
            rangoEdades = objDesafio3.ObtenerRangoEdades()
            
        Case 4
            Set objDesafio4 = New clsCalculoDesafio4
            objDesafio4.FechaNacimiento = pFechaNacimiento
            valorDesafio = objDesafio4.Calcular()
            rangoEdades = objDesafio4.ObtenerRangoEdades()
    End Select
    
    info = "=== DESAFÍO ACTUAL ===" & vbCrLf
    info = info & "Edad actual: " & pEdadActual & " años" & vbCrLf
    info = info & "Desafío activo: Desafío " & numeroDesafio & vbCrLf
    info = info & "Valor del desafío: " & valorDesafio & vbCrLf
    info = info & "Rango de edades: " & rangoEdades
    
    ObtenerInformacionDesafioActual = info
    
    Set objDesafio1 = Nothing
    Set objDesafio2 = Nothing
    Set objDesafio3 = Nothing
    Set objDesafio4 = Nothing
    
    Exit Function
    
ErrorHandler:
    Set objDesafio1 = Nothing
    Set objDesafio2 = Nothing
    Set objDesafio3 = Nothing
    Set objDesafio4 = Nothing
    Err.Raise Err.Number, "clsCalculadorPeriodoActual.ObtenerInformacionDesafioActual", Err.Description
End Function

' ============================================================================
' MÉTODO COMPLETO - TODOS LOS PERÍODOS ACTUALES
' ============================================================================

Public Function ObtenerResumenPeriodosActuales() As String
    Dim resumen As String
    
    On Error GoTo ErrorHandler
    
    resumen = "=== PERÍODOS ACTUALES ===" & vbCrLf & vbCrLf
    resumen = resumen & ObtenerInformacionCicloActual() & vbCrLf & vbCrLf
    resumen = resumen & ObtenerInformacionPinaculoActual() & vbCrLf & vbCrLf
    resumen = resumen & ObtenerInformacionDesafioActual()
    
    ObtenerResumenPeriodosActuales = resumen
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, "clsCalculadorPeriodoActual.ObtenerResumenPeriodosActuales", Err.Description
End Function

' ============================================================================
' MÉTODOS PRIVADOS
' ============================================================================

Private Function CalcularEdadActual() As Integer
    Dim fechaHoy As Date
    Dim anos As Integer
    
    fechaHoy = Date
    anos = Year(fechaHoy) - Year(pFechaNacimiento)
    
    ' Ajustar si aún no ha cumplido años este año
    If Month(fechaHoy) < Month(pFechaNacimiento) Or _
       (Month(fechaHoy) = Month(pFechaNacimiento) And Day(fechaHoy) < Day(pFechaNacimiento)) Then
        anos = anos - 1
    End If
    
    CalcularEdadActual = anos
End Function

' ============================================================================
' EVENTOS DE CLASE
' ============================================================================

Private Sub Class_Initialize()
    pFechaNacimiento = 0
    pEdadActual = 0
End Sub

Private Sub Class_Terminate()
    ' Limpiar
End Sub