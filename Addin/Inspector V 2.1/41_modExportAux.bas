Attribute VB_Name = "41_modExportAux"

Option Compare Database
Option Explicit

'===============================================================
' Módulo: 41_modExportAux
' Generación de resumen textual del análisis
'===============================================================


Public Function Inspector_Exportar( _
    formato As FormatoExportacion, _
    ruta As String, _
    Optional estilo As EstiloHtml = TemaClaro) As EstadoExportacion

    On Error GoTo ErrHandler

    '-----------------------------------------------------------
    ' Validación previa
    '-----------------------------------------------------------
    If gResultadosInspector Is Nothing Then
        Inspector_Log "Exportación no ejecutada: no hay resultados."
        Inspector_Exportar = ExportacionNoEjecutada
        Exit Function
    End If

    '-----------------------------------------------------------
    ' La ruta YA VIENE VALIDADA por ResolverRutaExportacion
    ' No tocarla, no añadir extensión, no modificar nada
    '-----------------------------------------------------------
    Dim rutaFinal As String
    rutaFinal = ruta

    ' Guardar últimos valores
    gUltimoFormato = formato
    gUltimaRuta = rutaFinal
    gUltimoEstiloHtml = estilo

    '-----------------------------------------------------------
    ' Exportación según formato
    '-----------------------------------------------------------
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
' Genera un resumen textual a partir de los resultados del análisis
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
' Funciones auxiliares
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


'---------------------------------------------------------------
' Devuelve la extensión de archivo correspondiente al formato
'---------------------------------------------------------------
' Se deja comentada por si se utiliza en versiones posteriores.
'---------------------------------------------------------------

'---------------------------------------------------------------
' Devuelve la extensión de archivo según el formato de exportación
'---------------------------------------------------------------

Public Function ExtensionDeFormato(formato As FormatoExportacion) As String
    Select Case formato

'        '-------------------------
'        ' TXT
'        '-------------------------
        Case ExpResultadosTXT, ExpSimbolosTXT, ExpTodoTXT
            ExtensionDeFormato = "txt"

'        '-------------------------
'        ' Excel
'        '-------------------------
        Case ExpResultadosExcel, ExpSimbolosExcel, ExpTodoExcel
            ExtensionDeFormato = "xlsx"

'        '-------------------------
'        ' HTML
'        '-------------------------
        Case ExpTodoHTML
            ExtensionDeFormato = "html"

        '-------------------------
        ' Futuros formatos
        '-------------------------
        'Case ExpTodoMarkdown
        '    ExtensionDeFormato = ".md"

        'Case ExpTodoPDF
        '    ExtensionDeFormato = ".pdf"

        Case Else
            ExtensionDeFormato = ".txt"   ' valor por defecto
    End Select
End Function


