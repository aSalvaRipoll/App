Attribute VB_Name = "modExportaHTML"

Option Compare Database
Option Explicit

'===============================================================
' Módulo: modExportaHTML
' Exportación de resultados y símbolos del Inspector en formato HTML
'===============================================================

'---------------------------------------------------------------
' Exporta TODO el análisis del Inspector a un único archivo HTML
'---------------------------------------------------------------
'Public Sub ExportarTodoAHTML(resultados As Collection, ByVal ruta As String, Optional modoOscuro As Boolean = False)
Public Sub ExportarTodoAHTML(resultados As Collection, ByVal ruta As String, estilo As EstiloHtml)

    Dim f As Integer
    Dim r As clsResultadoAnalisis
    Dim sim As clsSimbolo
    Dim stats As Object

    If resultados Is Nothing Then Exit Sub
    If gCatalogoSimbolos Is Nothing Then Exit Sub

    f = FreeFile
    Open ruta For Output As #f

    '===========================================================
    ' CABECERA HTML + CSS
    '===========================================================
    Print #f, "<!DOCTYPE html>"
    Print #f, "<html lang='es'>"
    Print #f, "<head>"
    Print #f, "<meta charset='UTF-8'>"
    Print #f, "<title>Informe Inspector</title>"
    Print #f, "<style>"


    Select Case estilo

        Case TemaClaro
            Print #f, "body { background: #f8f8f8; color: #222; font-family: Arial; margin: 20px; }"
            Print #f, "table { background: white; }"
            Print #f, "th { background: #333; color: white; }"
            Print #f, "tr:nth-child(even) { background: #f2f2f2; }"
            Print #f, ".sev-error { background: #ffcccc; }"
            Print #f, ".sev-aviso { background: #fff5cc; }"
            Print #f, ".sev-info  { background: #e6f2ff; }"
    
        Case TemaOscuro
            Print #f, "body { background: #1e1e1e; color: #ddd; font-family: Arial; margin: 20px; }"
            Print #f, "table { background: #2a2a2a; }"
            Print #f, "th { background: #444; color: #eee; }"
            Print #f, "tr:nth-child(even) { background: #333; }"
            Print #f, ".sev-error { background: #662222; }"
            Print #f, ".sev-aviso { background: #665522; }"
            Print #f, ".sev-info  { background: #224466; }"
    
        Case TemaSepia
            Print #f, "body { background: #f4ecd8; color: #5b4636; font-family: Georgia; margin: 20px; }"
            Print #f, "table { background: #fffaf0; }"
            Print #f, "th { background: #8b6f47; color: white; }"
            Print #f, "tr:nth-child(even) { background: #f0e4cc; }"
            Print #f, ".sev-error { background: #e6b8af; }"
            Print #f, ".sev-aviso { background: #f7e7c6; }"
            Print #f, ".sev-info  { background: #dce6f2; }"
    
        Case TemaContraste
            Print #f, "body { background: black; color: white; font-family: Arial; margin: 20px; }"
            Print #f, "table { background: black; }"
            Print #f, "th { background: yellow; color: black; }"
            Print #f, "tr:nth-child(even) { background: #222; }"
            Print #f, ".sev-error { background: red; color: white; }"
            Print #f, ".sev-aviso { background: orange; color: black; }"
            Print #f, ".sev-info  { background: cyan; color: black; }"
    
        Case TemaMinimalista
            Print #f, "body { background: white; color: black; font-family: 'Segoe UI'; margin: 20px; }"
            Print #f, "table { background: white; border: 1px solid #ddd; }"
            Print #f, "th { background: #fafafa; color: #333; border-bottom: 2px solid #ddd; }"
            Print #f, "tr:nth-child(even) { background: #fafafa; }"
            Print #f, ".sev-error { background: #fdd; }"
            Print #f, ".sev-aviso { background: #ffd; }"
            Print #f, ".sev-info  { background: #eef; }"
    
    End Select

    '------------------------------
    ' Estilos comunes
    '------------------------------
    Print #f, "h1 { text-align: center; }"
    Print #f, "h2 { margin-top: 40px; border-bottom: 2px solid #444; padding-bottom: 5px; }"
    Print #f, "table { width: 100%; border-collapse: collapse; margin-top: 10px; }"
    Print #f, "th, td { padding: 8px 10px; border: 1px solid #555; }"
    Print #f, "th { position: sticky; top: 0; }"

    Print #f, "</style>"
    Print #f, "</head>"

    '===========================================================
    ' BODY con clase de tema
    '===========================================================
    Print #f, "<body>"

    'Print #f, "<p><em>Tema: " & IIf(modoOscuro, "Oscuro", "Claro") & "</em></p>"
    
    Print #f, "<h1>Informe Completo del Inspector</h1>"
    Print #f, "<p><strong>Fecha:</strong> " & Format(Now, "dd/mm/yyyy hh:nn:ss") & "</p>"

    '===========================================================
    ' SECCIÓN 1: RESULTADOS
    '===========================================================
    Print #f, "<h2>1. Resultados del análisis</h2>"
    Print #f, "<table>"
    Print #f, "<tr><th>Código</th><th>Severidad</th><th>Tipo</th><th>Elemento</th><th>Miembro</th><th>Línea</th><th>Descripción</th><th>Detalles</th></tr>"

    For Each r In resultados
        Dim clase As String
        Select Case r.Severidad
            Case sevError: clase = "sev-error"
            Case sevAviso: clase = "sev-aviso"
            Case sevInfo:  clase = "sev-info"
        End Select

        Print #f, "<tr class='" & clase & "'>" & _
                  "<td>" & r.codigoRegla & "</td>" & _
                  "<td>" & SeveridadToText(r.Severidad) & "</td>" & _
                  "<td>" & TipoElementoToText(r.TipoElemento) & "</td>" & _
                  "<td>" & r.nombreElemento & "</td>" & _
                  "<td>" & r.nombreMiembro & "</td>" & _
                  "<td>" & r.linea & "</td>" & _
                  "<td>" & r.descripcion & "</td>" & _
                  "<td>" & r.Detalles & "</td>" & _
                  "</tr>"
    Next r

    Print #f, "</table>"

    '===========================================================
    ' SECCIÓN 2: SÍMBOLOS NO USADOS
    '===========================================================
    Print #f, "<h2>2. Símbolos no usados</h2>"
    Print #f, "<table>"
    Print #f, "<tr><th>Nombre</th><th>Categoría</th><th>Módulo</th><th>Miembro</th><th>Línea</th><th>Tipo</th></tr>"

    For Each sim In gCatalogoSimbolos.SimbolosNoUsados
        Print #f, "<tr>" & _
                  "<td>" & sim.nombre & "</td>" & _
                  "<td>" & sim.categoria & "</td>" & _
                  "<td>" & sim.modulo & "</td>" & _
                  "<td>" & sim.miembro & "</td>" & _
                  "<td>" & sim.LineaDeclaracion & "</td>" & _
                  "<td>" & sim.TipoTexto & "</td>" & _
                  "</tr>"
    Next sim

    Print #f, "</table>"

    '===========================================================
    ' SECCIÓN 3: ESTADÍSTICAS
    '===========================================================
    Print #f, "<h2>3. Estadísticas</h2>"
    Set stats = gCatalogoSimbolos.Estadisticas

    Print #f, "<table>"
    Print #f, "<tr><th>Concepto</th><th>Valor</th></tr>"
    Print #f, "<tr><td>Total símbolos</td><td>" & stats("Total") & "</td></tr>"
    Print #f, "<tr><td>Usados</td><td>" & stats("Usados") & "</td></tr>"
    Print #f, "<tr><td>No usados</td><td>" & stats("NoUsados") & "</td></tr>"
    Print #f, "</table>"

    '===========================================================
    ' SECCIÓN 4: RESUMEN DEL PROYECTO
    '===========================================================
    Print #f, "<h2>4. Resumen del proyecto</h2>"
    Print #f, "<table>"
    Print #f, "<tr><th>Elemento</th><th>Cantidad</th></tr>"
    Print #f, "<tr><td>Módulos estándar</td><td>" & gCatalogoInspector.Modulos.Count & "</td></tr>"
    Print #f, "<tr><td>Clases</td><td>" & gCatalogoInspector.Clases.Count & "</td></tr>"
    Print #f, "<tr><td>UserForms</td><td>" & gCatalogoInspector.UserForms.Count & "</td></tr>"
    Print #f, "<tr><td>Formularios</td><td>" & gCatalogoInspector.Formularios.Count & "</td></tr>"
    Print #f, "<tr><td>Informes</td><td>" & gCatalogoInspector.Informes.Count & "</td></tr>"
    Print #f, "</table>"

    '===========================================================
    ' PIE
    '===========================================================
    Print #f, "<p style='margin-top:40px; text-align:center;'>Informe generado automáticamente por el Inspector.</p>"
    Print #f, "</body></html>"

    Close #f
End Sub



