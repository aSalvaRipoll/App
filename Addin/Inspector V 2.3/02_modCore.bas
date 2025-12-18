Attribute VB_Name = "02_modCore"

Option Compare Database
Option Explicit

'===============================================================
' 02_modCore – Núcleo del InspectorVBA
' Lógica central: inicialización, análisis, reparación, estado
'===============================================================


'---------------------------------------------------------------
' Inicialización del motor del Inspector
'---------------------------------------------------------------
Public Sub Core_Inicializar()
    Set gCatalogoInspector = New clsCatalogoInspector
    Set gResultadosInspector = Nothing
End Sub


'---------------------------------------------------------------
' Ejecutar análisis completo del proyecto
'---------------------------------------------------------------
Public Function Core_Analizar() As EstadoAnalisis

    On Error GoTo ErrHandler

    ' Inicializar motor
    Core_Inicializar

    ' Ejecutar análisis
    Set gCatalogoInspector = AnalizarProyecto
    Set gResultadosInspector = EjecutarReglas(gCatalogoInspector)

    Inspector_Log "Análisis completado."
    Core_Analizar = AnalisisEjecutado
    Exit Function

ErrHandler:
    Inspector_Log "Error durante el análisis: " & Err.Description
    Core_Analizar = AnalisisConErrores
End Function


'---------------------------------------------------------------
' Reparar proyecto según resultados del análisis
'---------------------------------------------------------------
Public Function Core_Reparar() As EstadoReparacion

    On Error GoTo ErrHandler

    If gResultadosInspector Is Nothing Then
        Inspector_Log "Reparación no ejecutada: no hay resultados."
        Core_Reparar = ReparacionNoEjecutada
        Exit Function
    End If

    RepararResultados gResultadosInspector

    Inspector_Log "Reparación ejecutada sobre los resultados actuales."
    Core_Reparar = ReparacionEjecutada
    Exit Function

ErrHandler:
    Inspector_Log "Error durante la reparación: " & Err.Description
    Core_Reparar = ReparacionConErrores
End Function


'---------------------------------------------------------------
' Obtener resumen textual del análisis
'---------------------------------------------------------------
Public Function Core_Resumen() As String
    If gResultadosInspector Is Nothing Then
        Core_Resumen = MensajeAnalisis(AnalisisNoEjecutado)
    Else
        Core_Resumen = GenerarResumen(gResultadosInspector)
    End If
End Function


'---------------------------------------------------------------
' Reinicializar estado del Inspector
'---------------------------------------------------------------
Public Sub Core_Reset(Optional reiniciarMotor As Boolean = False)

    Set gResultadosInspector = Nothing

    If reiniciarMotor Then
        Set gCatalogoInspector = New clsCatalogoInspector
    End If

    gUltimoFormato = 0
    gUltimaRuta = vbNullString
    gUltimoEstiloHtml = TemaClaro

    Inspector_Log "Reset del Inspector ejecutado. MotorReiniciado=" & reiniciarMotor
End Sub

