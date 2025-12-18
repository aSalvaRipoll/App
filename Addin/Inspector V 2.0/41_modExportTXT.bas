Attribute VB_Name = "41_modExportTXT"

Option Compare Database
Option Explicit

'===============================================================
' Módulo: modExportaTXT
' Exportación de resultados y símbolos del Inspector a TXT
'===============================================================


'---------------------------------------------------------------
' Exporta resultados del análisis a archivo de texto
'---------------------------------------------------------------
Public Sub ExportarResultadosAArchivo(resultados As Collection, ByVal ruta As String)
    Dim f As Integer
    Dim r As clsResultadoAnalisis

    If resultados Is Nothing Then Exit Sub
    If resultados.Count = 0 Then Exit Sub

    f = FreeFile
    Open ruta For Output As #f

    Print #f, "CodigoRegla | Severidad | Tipo | Elemento | Miembro | Linea | Descripcion | Detalles"

    For Each r In resultados
        Print #f, r.ToTextFile
    Next r

    Close #f
End Sub



'---------------------------------------------------------------
' Exporta símbolos no usados a archivo de texto
'---------------------------------------------------------------
Public Sub ExportarSimbolosNoUsadosTXT(ByVal ruta As String)
    Dim col As Collection
    Dim sim As clsSimbolo
    Dim f As Integer

    If gCatalogoSimbolos Is Nothing Then Exit Sub

    Set col = gCatalogoSimbolos.SimbolosNoUsados
    If col.Count = 0 Then Exit Sub

    f = FreeFile
    Open ruta For Output As #f

    Print #f, "Nombre | Categoria | Modulo | Miembro | Linea | Tipo | Usado"

    For Each sim In col
        Print #f, sim.nombre & " | " & _
                    sim.categoria & " | " & _
                    sim.modulo & " | " & _
                    sim.miembro & " | " & _
                    sim.LineaDeclaracion & " | " & _
                    sim.TipoTexto & " | No"
    Next sim

    Close #f
End Sub



'---------------------------------------------------------------
' Exporta TODO el análisis del Inspector a un único archivo TXT
'---------------------------------------------------------------
Public Sub ExportarTodoATXT(resultados As Collection, ByVal ruta As String)
    Dim f As Integer
    Dim r As clsResultadoAnalisis
    Dim sim As clsSimbolo
    Dim stats As Object

    If resultados Is Nothing Then Exit Sub
    If gCatalogoSimbolos Is Nothing Then Exit Sub

    f = FreeFile
    Open ruta For Output As #f

    '===========================================================
    ' CABECERA
    '===========================================================
    Print #f, "============================================================"
    Print #f, "INSPECTOR – INFORME COMPLETO"
    Print #f, "Fecha: " & Format(Now, "dd/mm/yyyy hh:nn:ss")
    Print #f, "============================================================"
    Print #f, ""

    '===========================================================
    ' SECCIÓN 1: RESULTADOS DEL ANÁLISIS
    '===========================================================
    Print #f, "[1] RESULTADOS DEL ANÁLISIS"
    Print #f, "------------------------------------------------------------"

    For Each r In resultados
        Print #f, r.codigoRegla & " | " & _
                   SeveridadToText(r.severidad) & " | " & _
                   TipoElementoToText(r.tipoElemento) & " | " & _
                   r.nombreElemento & " | " & _
                   r.nombreMiembro & " | Línea " & r.linea & " | " & _
                   r.descripcion & " | " & r.Detalles
    Next r

    Print #f, ""
    Print #f, ""

    '===========================================================
    ' SECCIÓN 2: SÍMBOLOS NO USADOS
    '===========================================================
    Print #f, "[2] SÍMBOLOS NO USADOS"
    Print #f, "------------------------------------------------------------"

    For Each sim In gCatalogoSimbolos.SimbolosNoUsados
        Print #f, sim.nombre & " | " & _
                   sim.categoria & " | " & _
                   sim.modulo & " | " & _
                   sim.miembro & " | Línea " & sim.LineaDeclaracion & " | " & _
                   sim.TipoTexto & " | No usado"
    Next sim

    Print #f, ""
    Print #f, ""

    '===========================================================
    ' SECCIÓN 3: ESTADÍSTICAS
    '===========================================================
    Print #f, "[3] ESTADÍSTICAS"
    Print #f, "------------------------------------------------------------"

    Set stats = gCatalogoSimbolos.Estadisticas

    Print #f, "Total símbolos: " & stats("Total")
    Print #f, "Usados: " & stats("Usados")
    Print #f, "No usados: " & stats("NoUsados")

    Print #f, ""
    Print #f, ""

    '===========================================================
    ' SECCIÓN 4: RESUMEN DEL PROYECTO
    '===========================================================
    Print #f, "[4] RESUMEN DEL PROYECTO"
    Print #f, "------------------------------------------------------------"

    Print #f, "Módulos estándar: " & gCatalogoInspector.Modulos.Count
    Print #f, "Clases: " & gCatalogoInspector.Clases.Count
    Print #f, "UserForms: " & gCatalogoInspector.UserForms.Count
    Print #f, "Formularios: " & gCatalogoInspector.Formularios.Count
    Print #f, "Informes: " & gCatalogoInspector.Informes.Count

    Print #f, ""
    Print #f, "============================================================"
    Print #f, "FIN DEL INFORME"
    Print #f, "============================================================"

    Close #f
End Sub


