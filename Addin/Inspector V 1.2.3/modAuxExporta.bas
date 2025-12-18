Attribute VB_Name = "modAuxExporta"

Option Compare Database
Option Explicit

'===============================================================
' Módulo: modAuxExporta
' Funciones auxiliares comunes a la exportación
'===============================================================

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

'Public Function ExtensionDeFormato(formato As FormatoExportacion) As String
'    Select Case formato
'
'        '-------------------------
'        ' TXT
'        '-------------------------
'        Case ExpResultadosTXT, ExpSimbolosTXT, ExpTodoTXT
'            ExtensionDeFormato = ".txt"
'
'        '-------------------------
'        ' Excel
'        '-------------------------
'        Case ExpResultadosExcel, ExpSimbolosExcel, ExpTodoExcel
'            ExtensionDeFormato = ".xlsx"
'
'        '-------------------------
'        ' HTML
'        '-------------------------
'        Case ExpTodoHTML
'            ExtensionDeFormato = ".html"
'
'        '-------------------------
'        ' Futuros formatos
'        '-------------------------
'        'Case ExpTodoMarkdown
'        '    ExtensionDeFormato = ".md"
'
'        'Case ExpTodoPDF
'        '    ExtensionDeFormato = ".pdf"
'
'        Case Else
'            ExtensionDeFormato = ""
'    End Select
'End Function

