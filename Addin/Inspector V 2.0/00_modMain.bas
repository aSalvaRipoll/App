Attribute VB_Name = "00_modMain"

Option Compare Database
Option Explicit

'===============================================================
' 00_modMain - Módulo principal del Inspector
' Punto de entrada único para análisis, reparación y exportación
'===============================================================

'---------------------------------------------------------------
' Inicialización del motor del Inspector
'---------------------------------------------------------------
Public Sub Inspector_Inicializar()
    Set gCatalogoInspector = New clsCatalogoInspector
    Set gResultadosInspector = Nothing
End Sub

'---------------------------------------------------------------
' Ejecutar análisis completo del proyecto
'---------------------------------------------------------------
Public Function Inspector_Analizar() As EstadoAnalisis

    On Error GoTo ErrHandler

    ' Inicializar motor
    Inspector_Inicializar

    ' Ejecutar análisis
    Set gCatalogoInspector = AnalizarProyecto
    Set gResultadosInspector = EjecutarReglas(gCatalogoInspector)

    Inspector_Log "Análisis completado."
    Inspector_Analizar = AnalisisEjecutado
    Exit Function

ErrHandler:
    Inspector_Log "Error durante el análisis: " & Err.Description
    Inspector_Analizar = AnalisisConErrores
End Function

'---------------------------------------------------------------
' Reparar proyecto según resultados del análisis
'---------------------------------------------------------------
Public Function Inspector_Reparar() As EstadoReparacion

    On Error GoTo ErrHandler

    If gResultadosInspector Is Nothing Then
        Inspector_Log "Reparación no ejecutada: no hay resultados."
        Inspector_Reparar = ReparacionNoEjecutada
        Exit Function
    End If

    RepararResultados gResultadosInspector

    Inspector_Log "Reparación ejecutada sobre los resultados actuales."
    Inspector_Reparar = ReparacionEjecutada
    Exit Function

ErrHandler:
    Inspector_Log "Error durante la reparación: " & Err.Description
    Inspector_Reparar = ReparacionConErrores
End Function

'---------------------------------------------------------------
' Exportar resultados según formato seleccionado
'---------------------------------------------------------------
Public Function Inspector_Exportar( _
    formato As FormatoExportacion, _
    ruta As String, _
    Optional estilo As EstiloHtml = TemaClaro) As EstadoExportacion

    On Error GoTo ErrHandler

    Dim rutaFinal As String

    If gResultadosInspector Is Nothing Then
        Inspector_Log "Exportación no ejecutada: no hay resultados."
        Inspector_Exportar = ExportacionNoEjecutada
        Exit Function
    End If

    If InStr(1, ruta, ".", vbTextCompare) = 0 Then
        rutaFinal = ruta & ExtensionDeFormato(formato)
    Else
        rutaFinal = ruta
    End If

    gUltimoFormato = formato
    gUltimaRuta = rutaFinal
    gUltimoEstiloHtml = estilo

    Select Case formato
        Case ExpResultadosTXT
            ExportarResultadosAArchivo gResultadosInspector, rutaFinal
        Case ExpSimbolosTXT
            ExportarSimbolosNoUsadosTXT rutaFinal
        Case ExpTodoTXT
            ExportarTodoATXT gResultadosInspector, rutaFinal
        Case ExpResultadosExcel
            ExportarResultadosAExcel gResultadosInspector, rutaFinal
        Case ExpSimbolosExcel
            ExportarSimbolosNoUsadosExcel rutaFinal
        Case ExpTodoExcel
            ExportarTodoAExcel gResultadosInspector, rutaFinal
        Case ExpTodoHTML
            ExportarTodoAHTML gResultadosInspector, rutaFinal, estilo
        Case Else
            Inspector_Log "Formato no implementado: " & CStr(formato)
            Inspector_Exportar = ExportacionConErrores
            Exit Function
    End Select

    Inspector_Log "Exportación completada. Formato=" & CStr(formato) & _
                  "; Ruta='" & rutaFinal & "'; EstiloHtml=" & CStr(estilo)

    Inspector_Exportar = ExportacionEjecutada
    Exit Function

ErrHandler:
    Inspector_Log "Error durante la exportación: " & Err.Description
    Inspector_Exportar = ExportacionConErrores
End Function

'---------------------------------------------------------------
' Versión simplificada para la cinta u otras llamadas rápidas
'---------------------------------------------------------------
Public Sub Inspector_ExportarSimple(formato As FormatoExportacion, _
                                    Optional estilo As EstiloHtml = TemaClaro)

    Dim rutaBase As String

    If Len(gUltimaRuta) > 0 Then
        rutaBase = gUltimaRuta
    Else
        rutaBase = "InformeInspector"
    End If

    Inspector_Log "ExportarSimple invocado. Formato=" & CStr(formato) & _
                  "; RutaBase='" & rutaBase & "'; EstiloHtml=" & CStr(estilo)

    Inspector_Exportar formato, rutaBase, estilo
End Sub

'---------------------------------------------------------------
' Obtener resumen textual del análisis
'---------------------------------------------------------------
Public Function Inspector_Resumen() As String
    If gResultadosInspector Is Nothing Then
        Inspector_Resumen = MensajeAnalisis(AnalisisNoEjecutado)
    Else
        Inspector_Resumen = GenerarResumen(gResultadosInspector)   ' <-- corregido
    End If
End Function

'---------------------------------------------------------------
' Reinicializar estado del Inspector
'---------------------------------------------------------------
Public Sub Inspector_Reset(Optional reiniciarMotor As Boolean = False)

    Set gResultadosInspector = Nothing

    If reiniciarMotor Then
        Set gCatalogoInspector = New clsCatalogoInspector
    End If

    gUltimoFormato = 0
    gUltimaRuta = vbNullString
    gUltimoEstiloHtml = TemaClaro

    Inspector_Log "Reset del Inspector ejecutado. MotorReiniciado=" & reiniciarMotor
End Sub

