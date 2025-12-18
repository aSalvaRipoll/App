Attribute VB_Name = "40_modExportAux"

Option Compare Database
Option Explicit

'===============================================================
' Módulo: 40_modResumenInspector
' Generación de resumen textual del análisis
'===============================================================

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

        '-------------------------
        ' TXT
        '-------------------------
        Case ExpResultadosTXT, ExpSimbolosTXT, ExpTodoTXT
            ExtensionDeFormato = ".txt"

        '-------------------------
        ' Excel
        '-------------------------
        Case ExpResultadosExcel, ExpSimbolosExcel, ExpTodoExcel
            ExtensionDeFormato = ".xlsx"

        '-------------------------
        ' HTML
        '-------------------------
        Case ExpTodoHTML
            ExtensionDeFormato = ".html"

        '-------------------------
        ' Futuros formatos
        '-------------------------
        'Case ExpTodoMarkdown
        '    ExtensionDeFormato = ".md"

        'Case ExpTodoPDF
        '    ExtensionDeFormato = ".pdf"

        Case Else
            'ExtensionDeFormato = ".dat"   ' valor por defecto
            ExtensionDeFormato = ".txt"   ' valor por defecto
    End Select
End Function

