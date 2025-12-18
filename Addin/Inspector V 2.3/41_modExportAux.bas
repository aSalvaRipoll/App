Attribute VB_Name = "41_modExportAux"

Option Compare Database
Option Explicit

'===============================================================
' Módulo: 41_modExportAux
' Subsistema de exportación del InspectorVBA
'===============================================================


'---------------------------------------------------------------
' EXPORTACIÓN PRINCIPAL
'---------------------------------------------------------------
Public Function Inspector_Exportar( _
    formato As FormatoExportacion, _
    ruta As String, _
    Optional estilo As EstiloHtml = TemaClaro) As EstadoExportacion

    On Error GoTo ErrHandler

    ' Validación previa
    If gResultadosInspector Is Nothing Then
        Inspector_Log "Exportación no ejecutada: no hay resultados."
        Inspector_Exportar = ExportacionNoEjecutada
        Exit Function
    End If

    ' La ruta YA VIENE validada por ResolverRutaExportacion
    Dim rutaFinal As String
    rutaFinal = ruta

    ' Guardar últimos valores
    gUltimoFormato = formato
    gUltimaRuta = rutaFinal
    gUltimoEstiloHtml = estilo

    ' Exportación según formato
    Select Case formato

        Case ExpResultadosTXT
            ExportarResultadosAArchivo gResultadosInspector, rutaFinal

        Case ExpSimbolosTXT
            ExportarSimbolosNoUsadosTXT gCatalogoInspector, rutaFinal

        Case ExpTodoTXT
            ExportarTodoATXT gCatalogoInspector, gResultadosInspector, rutaFinal

        Case ExpResultadosExcel
            ExportarResultadosAExcel gResultadosInspector, rutaFinal

        Case ExpSimbolosExcel
            ExportarSimbolosNoUsadosExcel gCatalogoInspector, rutaFinal

        Case ExpTodoExcel
            ExportarTodoAExcel gCatalogoInspector, gResultadosInspector, rutaFinal

        Case ExpTodoHTML
            ExportarTodoAHTML gCatalogoInspector, gResultadosInspector, rutaFinal, estilo

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
' EXPORTACIÓN SIMPLE (cinta, accesos rápidos)
'---------------------------------------------------------------
Public Sub Inspector_ExportarSimple( _
    formato As FormatoExportacion, _
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
' RESUMEN DEL ANÁLISIS
'---------------------------------------------------------------
Public Function Inspector_Resumen() As String
    If gResultadosInspector Is Nothing Then
        Inspector_Resumen = MensajeAnalisis(AnalisisNoEjecutado)
    Else
        Inspector_Resumen = GenerarResumen(gResultadosInspector)
    End If
End Function






'---------------------------------------------------------------
' GENERACIÓN DE RESUMEN
'---------------------------------------------------------------
Public Function GenerarResumen(resultados As Collection) As String
    Dim res As clsResultadoAnalisis
    Dim resumen As String

    If resultados Is Nothing Or resultados.Count = 0 Then
        GenerarResumen = "No hay resultados disponibles."
        Exit Function
    End If

    resumen = "Resumen del análisis:" & vbCrLf & String(30, "-") & vbCrLf

    For Each res In resultados
        resumen = resumen & res.Formatear & vbCrLf
    Next res

    GenerarResumen = resumen
End Function



'---------------------------------------------------------------
' EXTENSIÓN SEGÚN FORMATO
'---------------------------------------------------------------
Public Function ExtensionDeFormato(formato As FormatoExportacion) As String
    Select Case formato

        Case ExpResultadosTXT, ExpSimbolosTXT, ExpTodoTXT
            ExtensionDeFormato = "txt"

        Case ExpResultadosExcel, ExpSimbolosExcel, ExpTodoExcel
            ExtensionDeFormato = "xlsx"

        Case ExpTodoHTML
            ExtensionDeFormato = "html"

        Case Else
            ExtensionDeFormato = "txt"   ' valor por defecto
    End Select
End Function



'---------------------------------------------------------------
' UTILIDADES DE TEXTO
'---------------------------------------------------------------
Public Function SeveridadToText(sev As SeveridadInspector) As String
    Select Case sev
        Case sevInfo: SeveridadToText = "INFO"
        Case sevAviso: SeveridadToText = "AVISO"
        Case sevError: SeveridadToText = "ERROR"
        Case Else: SeveridadToText = "?"
    End Select
End Function

Public Function TipoElementoToText(t As TipoElementoInspector) As String
    Select Case t
        Case teProyecto:   TipoElementoToText = "Proyecto"
        Case teModulo:     TipoElementoToText = "Módulo"
        Case teClase:      TipoElementoToText = "Clase"
        Case teUserForm:   TipoElementoToText = "UserForm"
        Case teFormulario: TipoElementoToText = "Formulario"
        Case teInforme:    TipoElementoToText = "Informe"
        Case teMiembro:    TipoElementoToText = "Miembro"
        Case Else:         TipoElementoToText = "Elemento"
    End Select
End Function


