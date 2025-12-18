Attribute VB_Name = "modInspectorMain"

Option Compare Database
Option Explicit

'===============================================================
' Módulo principal del Inspector
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
    Set gResultadosInspector = gCatalogoInspector.AnalizarProyecto

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
'---------------------------------------------------------------
' Reparar proyecto según resultados del análisis
'   - devuelve EstadoReparacion
'   - registra en log
'   - no muestra MsgBox (UI lo hará)
'---------------------------------------------------------------
Public Function Inspector_Reparar() As EstadoReparacion

    On Error GoTo ErrHandler

    '-----------------------------------------------------------
    ' Validación: no hay resultados
    '-----------------------------------------------------------
    If gResultadosInspector Is Nothing Then
        Inspector_Log "Reparación no ejecutada: no hay resultados."
        Inspector_Reparar = ReparacionNoEjecutada
        Exit Function
    End If

    '-----------------------------------------------------------
    ' Ejecutar reparación
    '-----------------------------------------------------------
    gResultadosInspector.RepararProyecto

    Inspector_Log "Reparación ejecutada sobre los resultados actuales."
    Inspector_Reparar = ReparacionEjecutada
    Exit Function

'-----------------------------------------------------------
' Manejo de errores
'-----------------------------------------------------------
ErrHandler:
    Inspector_Log "Error durante la reparación: " & Err.Description
    Inspector_Reparar = ReparacionConErrores
End Function

'Public Sub Inspector_Reparar()
'
'    If gResultadosInspector Is Nothing Then
'        MsgBox "No hay resultados para reparar. Ejecuta primero el análisis.", vbExclamation
'        Exit Sub
'    End If
'
'    gResultadosInspector.RepararProyecto
'
'    Inspector_Log "Reparación ejecutada sobre los resultados actuales."
'    MsgBox "Reparación completada.", vbInformation
'End Sub

'---------------------------------------------------------------
' Exportar resultados según formato seleccionado
'   - ruta puede ser ruta base (sin extensión)
'   - la extensión se deriva del formato
'   - se guarda el último estado de exportación
'   - se registra la operación en el log
'---------------------------------------------------------------
'---------------------------------------------------------------
' Exportar resultados según formato seleccionado
'   - ruta puede ser ruta base (sin extensión)
'   - la extensión se deriva del formato
'   - se guarda el último estado de exportación
'   - se registra la operación en el log
'   - devuelve EstadoExportacion
'---------------------------------------------------------------
Public Function Inspector_Exportar( _
    formato As FormatoExportacion, _
    ruta As String, _
    Optional estilo As EstiloHtml = TemaClaro) As EstadoExportacion

    On Error GoTo ErrHandler

    Dim rutaFinal As String

    '-----------------------------------------------------------
    ' Validación: no hay resultados
    '-----------------------------------------------------------
    If gResultadosInspector Is Nothing Then
        Inspector_Log "Exportación no ejecutada: no hay resultados."
        Inspector_Exportar = ExportacionNoEjecutada
        Exit Function
    End If

    '-----------------------------------------------------------
    ' Determinar ruta final con extensión
    '-----------------------------------------------------------
    If InStr(1, ruta, ".", vbTextCompare) = 0 Then
        rutaFinal = ruta & ExtensionDeFormato(formato)
    Else
        rutaFinal = ruta
    End If

    '-----------------------------------------------------------
    ' Guardar estado de exportación
    '-----------------------------------------------------------
    gUltimoFormato = formato
    gUltimaRuta = rutaFinal
    gUltimoEstiloHtml = estilo

    '-----------------------------------------------------------
    ' Ejecutar exportación según formato
    '-----------------------------------------------------------
    Select Case formato

        '===========================
        ' TXT
        '===========================
        Case ExpResultadosTXT
            ExportarResultadosAArchivo gResultadosInspector, rutaFinal

        Case ExpSimbolosTXT
            ExportarSimbolosNoUsadosTXT rutaFinal

        Case ExpTodoTXT
            ExportarTodoATXT gResultadosInspector, rutaFinal


        '===========================
        ' EXCEL
        '===========================
        Case ExpResultadosExcel
            ExportarResultadosAExcel gResultadosInspector, rutaFinal

        Case ExpSimbolosExcel
            ExportarSimbolosNoUsadosExcel rutaFinal

        Case ExpTodoExcel
            ExportarTodoAExcel gResultadosInspector, rutaFinal


        '===========================
        ' HTML
        '===========================
        Case ExpTodoHTML
            ExportarTodoAHTML gResultadosInspector, rutaFinal, estilo


        '===========================
        ' FUTUROS FORMATOS
        '===========================
        Case Else
            Inspector_Log "Formato no implementado: " & CStr(formato)
            Inspector_Exportar = ExportacionConErrores
            Exit Function

    End Select

    '-----------------------------------------------------------
    ' Registro final
    '-----------------------------------------------------------
    Inspector_Log "Exportación completada. Formato=" & CStr(formato) & _
                  "; Ruta='" & rutaFinal & "'; EstiloHtml=" & CStr(estilo)

    Inspector_Exportar = ExportacionEjecutada
    Exit Function

'---------------------------------------------------------------
' Manejo de errores
'---------------------------------------------------------------
ErrHandler:
    Inspector_Log "Error durante la exportación: " & Err.Description
    Inspector_Exportar = ExportacionConErrores
End Function


'---------------------------------------------------------------
' Exportar resultados según formato seleccionado
'   - ruta puede ser ruta base (sin extensión)
'   - la extensión se deriva del formato
'   - se guarda el último estado de exportación
'   - se registra la operación en el log
'---------------------------------------------------------------
'Public Sub Inspector_Exportar(formato As FormatoExportacion, ruta As String, _
'                              Optional estilo As EstiloHtml = TemaClaro)
'
'    Dim rutaFinal As String
'
'    If gResultadosInspector Is Nothing Then
'        MsgBox "No hay resultados para exportar. Ejecuta primero el análisis.", vbExclamation
'        Exit Sub
'    End If
'
'    ' Añadir extensión si no está incluida
'    If InStr(1, ruta, ".", vbTextCompare) = 0 Then
'        rutaFinal = ruta & ExtensionDeFormato(formato)
'    Else
'        rutaFinal = ruta
'    End If
'
'    ' Guardar estado de la última exportación
'    gUltimoFormato = formato
'    gUltimaRuta = rutaFinal
'    gUltimoEstiloHtml = estilo
'
'    Select Case formato
'
'        '=======================================================
'        ' TXT
'        '=======================================================
'        Case ExpResultadosTXT
'            ExportarResultadosAArchivo gResultadosInspector, rutaFinal
'
'        Case ExpSimbolosTXT
'            ExportarSimbolosNoUsadosTXT rutaFinal
'
'        Case ExpTodoTXT
'            ExportarTodoATXT gResultadosInspector, rutaFinal
'
'
'        '=======================================================
'        ' EXCEL
'        '=======================================================
'        Case ExpResultadosExcel
'            ExportarResultadosAExcel gResultadosInspector, rutaFinal
'
'        Case ExpSimbolosExcel
'            ExportarSimbolosNoUsadosExcel rutaFinal
'
'        Case ExpTodoExcel
'            ExportarTodoAExcel gResultadosInspector, rutaFinal
'
'
'        '=======================================================
'        ' HTML
'        '=======================================================
'        Case ExpTodoHTML
'            ExportarTodoAHTML gResultadosInspector, rutaFinal, estilo
'
'
'        '=======================================================
'        ' FUTUROS FORMATOS
'        '=======================================================
'        Case Else
'            MsgBox "Este formato aún no está implementado.", vbInformation
'            Inspector_Log "Intento de exportación en formato no implementado. Formato=" & CStr(formato)
'            Exit Sub
'
'    End Select
'
'    Inspector_Log "Exportación completada. Formato=" & CStr(formato) & _
'                  "; Ruta='" & rutaFinal & "'; EstiloHtml=" & CStr(estilo)
'
'    MsgBox "Exportación completada correctamente.", vbInformation
'End Sub



'---------------------------------------------------------------
' Versión simplificada para la cinta u otras llamadas rápidas
'   - Usa la última ruta recordada o un nombre por defecto
'   - Delega en Inspector_Exportar
'---------------------------------------------------------------
Public Sub Inspector_ExportarSimple(formato As FormatoExportacion, _
                                    Optional estilo As EstiloHtml = TemaClaro)

    Dim rutaBase As String

    ' Si hay una última ruta conocida, reutilizarla como base
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
        Inspector_Resumen = gResultadosInspector.GenerarResumen
    End If
End Function



'---------------------------------------------------------------
' Registrar mensaje en el log del Inspector
'   - De momento: salida inmediata a la ventana de depuración
'   - Futuro: redirección a archivo de texto o tabla
'---------------------------------------------------------------
Public Sub Inspector_Log(mensaje As String)

    Dim texto As String
    texto = Format$(Now, "yyyy-mm-dd hh:nn:ss") & " - " & mensaje

    Debug.Print texto

    ' FUTURO:
    'Call AppendTextToFile(LogPath, texto)
    'o
    'CurrentDb.Execute "INSERT INTO T_InspectorLog(FechaHora, Mensaje) VALUES(#" & _
    '                  Format$(Now, "yyyy\/mm\/dd hh:nn:ss") & "#, " & Chr$(34) & mensaje & Chr$(34) & ")"

End Sub

'---------------------------------------------------------------
' Reinicializar estado del Inspector
'   - Limpia resultados
'   - Opcionalmente reinicia el motor
'   - Limpia estado de exportación
'   - Registra la operación
'---------------------------------------------------------------
Public Sub Inspector_Reset(Optional reiniciarMotor As Boolean = False)

    ' Limpiar resultados
    Set gResultadosInspector = Nothing

    ' Reiniciar motor si se solicita
    If reiniciarMotor Then
        Set gCatalogoInspector = New clsCatalogoInspector
    End If

    ' Limpiar estado de exportación
    gUltimoFormato = 0
    gUltimaRuta = vbNullString
    gUltimoEstiloHtml = TemaClaro

    Inspector_Log "Reset del Inspector ejecutado. MotorReiniciado=" & reiniciarMotor
End Sub

