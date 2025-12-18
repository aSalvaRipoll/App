Attribute VB_Name = "modExportaInspector"

Option Compare Database
Option Explicit

'===============================================================
' Módulo: modExportaInspector
' Exportación de resultados y símbolos del Inspector
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
' Exporta resultados del análisis directamente a Excel
'---------------------------------------------------------------
Public Sub ExportarResultadosAExcel(resultados As Collection, ByVal ruta As String)
    Dim xl As Object, wb As Object, ws As Object
    Dim r As clsResultadoAnalisis
    Dim fila As Long

    If resultados Is Nothing Then Exit Sub
    If resultados.Count = 0 Then Exit Sub

    Set xl = CreateObject("Excel.Application")
    Set wb = xl.Workbooks.Add
    Set ws = wb.Sheets(1)

    ws.Name = "Resultados"

    ' Encabezados
    ws.Range("A1:H1").Value = Array("Código", "Severidad", "Tipo", "Elemento", "Miembro", "Línea", "Descripción", "Detalles")
    ws.Range("A1:H1").Font.Bold = True

    fila = 2

    For Each r In resultados
        ws.Cells(fila, 1).Value = r.codigoRegla
        ws.Cells(fila, 2).Value = SeveridadToText(r.Severidad)
        ws.Cells(fila, 3).Value = TipoElementoToText(r.TipoElemento)
        ws.Cells(fila, 4).Value = r.nombreElemento
        ws.Cells(fila, 5).Value = r.nombreMiembro
        ws.Cells(fila, 6).Value = r.linea
        ws.Cells(fila, 7).Value = r.descripcion
        ws.Cells(fila, 8).Value = r.Detalles

        ' Colorear según severidad
        Select Case r.Severidad
            Case sevError: ws.Rows(fila).Interior.Color = RGB(255, 200, 200)
            Case sevAviso: ws.Rows(fila).Interior.Color = RGB(255, 255, 200)
            Case sevInfo:  ws.Rows(fila).Interior.Color = RGB(220, 240, 255)
        End Select

        ' Zebra
        If fila Mod 2 = 0 Then
            ws.Rows(fila).Interior.Color = RGB(245, 245, 245)
        End If

        fila = fila + 1
    Next r

    ' Formato automático
    ws.Columns("A:H").AutoFit
    ws.Range("A1:H1").AutoFilter

    ' Bordes
    ws.Range("A1:H" & fila - 1).Borders.LineStyle = 1

    ' Congelar encabezado
    ws.Rows("2:2").Select
    xl.ActiveWindow.FreezePanes = True

    wb.SaveAs ruta
    wb.Close False
    xl.Quit
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
        Print #f, sim.nombre & " | " & sim.categoria & " | " & sim.modulo & _
                    " | " & sim.miembro & " | " & sim.LineaDeclaracion & _
                    " | " & sim.TipoTexto & " | No"
    Next sim

    Close #f
End Sub

'---------------------------------------------------------------
' Exporta símbolos no usados directamente a Excel
'---------------------------------------------------------------
Public Sub ExportarSimbolosNoUsadosExcel(ByVal ruta As String)
    Dim col As Collection
    Dim sim As clsSimbolo
    Dim xl As Object, wb As Object, ws As Object
    Dim fila As Long

    If gCatalogoSimbolos Is Nothing Then Exit Sub

    Set col = gCatalogoSimbolos.SimbolosNoUsados
    If col.Count = 0 Then Exit Sub

    Set xl = CreateObject("Excel.Application")
    Set wb = xl.Workbooks.Add
    Set ws = wb.Sheets(1)

    ws.Name = "SimbolosNoUsados"

    ws.Range("A1:G1").Value = Array("Nombre", "Categoría", "Módulo", "Miembro", "Línea", "Tipo", "Usado")
    ws.Range("A1:G1").Font.Bold = True

    fila = 2

    For Each sim In col
        ws.Cells(fila, 1).Value = sim.nombre
        ws.Cells(fila, 2).Value = sim.categoria
        ws.Cells(fila, 3).Value = sim.modulo
        ws.Cells(fila, 4).Value = sim.miembro
        ws.Cells(fila, 5).Value = sim.LineaDeclaracion
        ws.Cells(fila, 6).Value = sim.TipoTexto
        ws.Cells(fila, 7).Value = "No"

        ' Zebra
        If fila Mod 2 = 0 Then
            ws.Rows(fila).Interior.Color = RGB(245, 245, 245)
        End If

        fila = fila + 1
    Next sim

    ws.Columns("A:G").AutoFit
    ws.Range("A1:G1").AutoFilter
    ws.Range("A1:G" & fila - 1).Borders.LineStyle = 1

    ws.Rows("2:2").Select
    xl.ActiveWindow.FreezePanes = True

    wb.SaveAs ruta
    wb.Close False
    xl.Quit
End Sub

'---------------------------------------------------------------
' Exporta TODO el análisis del Inspector a un único libro Excel
'---------------------------------------------------------------
Public Sub ExportarTodoAExcel(resultados As Collection, ByVal ruta As String)
    Dim xl As Object, wb As Object, ws As Object
    Dim fila As Long
    Dim r As clsResultadoAnalisis
    Dim sim As clsSimbolo
    Dim stats As Object

    If resultados Is Nothing Then Exit Sub
    If gCatalogoSimbolos Is Nothing Then Exit Sub

    Set xl = CreateObject("Excel.Application")
    Set wb = xl.Workbooks.Add

    '===========================================================
    ' HOJA 1: RESULTADOS
    '===========================================================
    Set ws = wb.Sheets(1)
    ws.Name = "Resultados"

    ws.Range("A1:H1").Value = Array("Código", "Severidad", "Tipo", "Elemento", "Miembro", "Línea", "Descripción", "Detalles")
    ws.Range("A1:H1").Font.Bold = True

    fila = 2

    For Each r In resultados
        ws.Cells(fila, 1).Value = r.codigoRegla
        ws.Cells(fila, 2).Value = SeveridadToText(r.Severidad)
        ws.Cells(fila, 3).Value = TipoElementoToText(r.TipoElemento)
        ws.Cells(fila, 4).Value = r.nombreElemento
        ws.Cells(fila, 5).Value = r.nombreMiembro
        ws.Cells(fila, 6).Value = r.linea
        ws.Cells(fila, 7).Value = r.descripcion
        ws.Cells(fila, 8).Value = r.Detalles

        Select Case r.Severidad
            Case sevError: ws.Rows(fila).Interior.Color = RGB(255, 200, 200)
            Case sevAviso: ws.Rows(fila).Interior.Color = RGB(255, 255, 200)
            Case sevInfo:  ws.Rows(fila).Interior.Color = RGB(220, 240, 255)
        End Select

        If fila Mod 2 = 0 Then ws.Rows(fila).Interior.Color = RGB(245, 245, 245)

        fila = fila + 1
    Next r

    ws.Columns("A:H").AutoFit
    ws.Range("A1:H1").AutoFilter
    ws.Range("A1:H" & fila - 1).Borders.LineStyle = 1
    ws.Rows("2:2").Select
    xl.ActiveWindow.FreezePanes = True

    '===========================================================
    ' HOJA 2: SÍMBOLOS NO USADOS
    '===========================================================
    Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    ws.Name = "SimbolosNoUsados"

    ws.Range("A1:G1").Value = Array("Nombre", "Categoría", "Módulo", "Miembro", "Línea", "Tipo", "Usado")
    ws.Range("A1:G1").Font.Bold = True

    fila = 2

    For Each sim In gCatalogoSimbolos.SimbolosNoUsados
        ws.Cells(fila, 1).Value = sim.nombre
        ws.Cells(fila, 2).Value = sim.categoria
        ws.Cells(fila, 3).Value = sim.modulo
        ws.Cells(fila, 4).Value = sim.miembro
        ws.Cells(fila, 5).Value = sim.LineaDeclaracion
        ws.Cells(fila, 6).Value = sim.TipoTexto
        ws.Cells(fila, 7).Value = "No"

        If fila Mod 2 = 0 Then ws.Rows(fila).Interior.Color = RGB(245, 245, 245)

        fila = fila + 1
    Next sim

    ws.Columns("A:G").AutoFit
    ws.Range("A1:G1").AutoFilter
    ws.Range("A1:G" & fila - 1).Borders.LineStyle = 1
    ws.Rows("2:2").Select
    xl.ActiveWindow.FreezePanes = True

    '===========================================================
    ' HOJA 3: ESTADÍSTICAS
    '===========================================================
    Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    ws.Name = "Estadisticas"

    Set stats = gCatalogoSimbolos.Estadisticas

    ws.Range("A1:B1").Value = Array("Concepto", "Valor")
    ws.Range("A1:B1").Font.Bold = True

    ws.Cells(2, 1).Value = "Total símbolos"
    ws.Cells(2, 2).Value = stats("Total")

    ws.Cells(3, 1).Value = "Usados"
    ws.Cells(3, 2).Value = stats("Usados")

    ws.Cells(4, 1).Value = "No usados"
    ws.Cells(4, 2).Value = stats("NoUsados")

    ws.Columns("A:B").AutoFit
    ws.Range("A1:B4").Borders.LineStyle = 1

    '===========================================================
    ' HOJA 4: RESUMEN DEL PROYECTO
    '===========================================================
    Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    ws.Name = "ResumenProyecto"

    ws.Range("A1:B1").Value = Array("Elemento", "Cantidad")
    ws.Range("A1:B1").Font.Bold = True

    ws.Cells(2, 1).Value = "Módulos estándar"
    ws.Cells(2, 2).Value = gCatalogoInspector.Modulos.Count

    ws.Cells(3, 1).Value = "Clases"
    ws.Cells(3, 2).Value = gCatalogoInspector.Clases.Count

    ws.Cells(4, 1).Value = "UserForms"
    ws.Cells(4, 2).Value = gCatalogoInspector.UserForms.Count

    ws.Cells(5, 1).Value = "Formularios"
    ws.Cells(5, 2).Value = gCatalogoInspector.Formularios.Count

    ws.Cells(6, 1).Value = "Informes"
    ws.Cells(6, 2).Value = gCatalogoInspector.Informes.Count

    ws.Columns("A:B").AutoFit
    ws.Range("A1:B6").Borders.LineStyle = 1

    '===========================================================
    ' Guardar y cerrar
    '===========================================================
    wb.SaveAs ruta
    wb.Close False
    xl.Quit
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
                   SeveridadToText(r.Severidad) & " | " & _
                   TipoElementoToText(r.TipoElemento) & " | " & _
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





'Public Sub ExportarTodoAHTML(resultados As Collection, ByVal ruta As String, Optional modoOscuro As Boolean = False)
'    Dim f As Integer
'    Dim r As clsResultadoAnalisis
'    Dim sim As clsSimbolo
'    Dim stats As Object
'
'    If resultados Is Nothing Then Exit Sub
'    If gCatalogoSimbolos Is Nothing Then Exit Sub
'
'    f = FreeFile
'    Open ruta For Output As #f
'
'    '===========================================================
'    ' CABECERA HTML + CSS
'    '===========================================================
'    Print #f, "<!DOCTYPE html>"
'    Print #f, "<html lang='es'>"
'    Print #f, "<head>"
'    Print #f, "<meta charset='UTF-8'>"
'    Print #f, "<title>Informe Inspector</title>"
'    Print #f, "<style>"
'    Print #f, "body { font-family: Arial, sans-serif; margin: 20px; background: #f8f8f8; }"
'    Print #f, "h1 { text-align: center; }"
'    Print #f, "h2 { margin-top: 40px; border-bottom: 2px solid #444; padding-bottom: 5px; }"
'    Print #f, "table { width: 100%; border-collapse: collapse; margin-top: 10px; }"
'    Print #f, "th, td { padding: 8px 10px; border: 1px solid #ccc; }"
'    Print #f, "th { background: #333; color: white; position: sticky; top: 0; }"
'    Print #f, "tr:nth-child(even) { background: #f2f2f2; }"
'    Print #f, ".sev-error { background: #ffcccc; }"
'    Print #f, ".sev-aviso { background: #fff5cc; }"
'    Print #f, ".sev-info { background: #e6f2ff; }"
'    Print #f, "</style>"
'    Print #f, "</head>"
'    Print #f, "<body>"
'
'    Print #f, "<h1>Informe Completo del Inspector</h1>"
'    Print #f, "<p><strong>Fecha:</strong> " & Format(Now, "dd/mm/yyyy hh:nn:ss") & "</p>"
'
'    '===========================================================
'    ' SECCIÓN 1: RESULTADOS DEL ANÁLISIS
'    '===========================================================
'    Print #f, "<h2>1. Resultados del análisis</h2>"
'    Print #f, "<table>"
'    Print #f, "<tr><th>Código</th><th>Severidad</th><th>Tipo</th><th>Elemento</th><th>Miembro</th><th>Línea</th><th>Descripción</th><th>Detalles</th></tr>"
'
'    For Each r In resultados
'        Dim clase As String
'        Select Case r.Severidad
'            Case sevError: clase = "sev-error"
'            Case sevAviso: clase = "sev-aviso"
'            Case sevInfo:  clase = "sev-info"
'        End Select
'
'        Print #f, "<tr class='" & clase & "'>" & _
'                  "<td>" & r.codigoRegla & "</td>" & _
'                  "<td>" & SeveridadToText(r.Severidad) & "</td>" & _
'                  "<td>" & TipoElementoToText(r.TipoElemento) & "</td>" & _
'                  "<td>" & r.nombreElemento & "</td>" & _
'                  "<td>" & r.nombreMiembro & "</td>" & _
'                  "<td>" & r.linea & "</td>" & _
'                  "<td>" & r.descripcion & "</td>" & _
'                  "<td>" & r.Detalles & "</td>" & _
'                  "</tr>"
'    Next r
'
'    Print #f, "</table>"
'
'    '===========================================================
'    ' SECCIÓN 2: SÍMBOLOS NO USADOS
'    '===========================================================
'    Print #f, "<h2>2. Símbolos no usados</h2>"
'    Print #f, "<table>"
'    Print #f, "<tr><th>Nombre</th><th>Categoría</th><th>Módulo</th><th>Miembro</th><th>Línea</th><th>Tipo</th></tr>"
'
'    For Each sim In gCatalogoSimbolos.SimbolosNoUsados
'        Print #f, "<tr>" & _
'                  "<td>" & sim.nombre & "</td>" & _
'                  "<td>" & sim.categoria & "</td>" & _
'                  "<td>" & sim.modulo & "</td>" & _
'                  "<td>" & sim.miembro & "</td>" & _
'                  "<td>" & sim.LineaDeclaracion & "</td>" & _
'                  "<td>" & sim.TipoTexto & "</td>" & _
'                  "</tr>"
'    Next sim
'
'    Print #f, "</table>"
'
'    '===========================================================
'    ' SECCIÓN 3: ESTADÍSTICAS
'    '===========================================================
'    Print #f, "<h2>3. Estadísticas</h2>"
'    Set stats = gCatalogoSimbolos.Estadisticas
'
'    Print #f, "<table>"
'    Print #f, "<tr><th>Concepto</th><th>Valor</th></tr>"
'    Print #f, "<tr><td>Total símbolos</td><td>" & stats("Total") & "</td></tr>"
'    Print #f, "<tr><td>Usados</td><td>" & stats("Usados") & "</td></tr>"
'    Print #f, "<tr><td>No usados</td><td>" & stats("NoUsados") & "</td></tr>"
'    Print #f, "</table>"
'
'    '===========================================================
'    ' SECCIÓN 4: RESUMEN DEL PROYECTO
'    '===========================================================
'    Print #f, "<h2>4. Resumen del proyecto</h2>"
'    Print #f, "<table>"
'    Print #f, "<tr><th>Elemento</th><th>Cantidad</th></tr>"
'    Print #f, "<tr><td>Módulos estándar</td><td>" & gCatalogoInspector.Modulos.Count & "</td></tr>"
'    Print #f, "<tr><td>Clases</td><td>" & gCatalogoInspector.Clases.Count & "</td></tr>"
'    Print #f, "<tr><td>UserForms</td><td>" & gCatalogoInspector.UserForms.Count & "</td></tr>"
'    Print #f, "<tr><td>Formularios</td><td>" & gCatalogoInspector.Formularios.Count & "</td></tr>"
'    Print #f, "<tr><td>Informes</td><td>" & gCatalogoInspector.Informes.Count & "</td></tr>"
'    Print #f, "</table>"
'
'    '===========================================================
'    ' PIE
'    '===========================================================
'    Print #f, "<p style='margin-top:40px; text-align:center;'>Informe generado automáticamente por el Inspector.</p>"
'    Print #f, "</body></html>"
'
'    Close #f
'End Sub



'---------------------------------------------------------------
' Funciones auxiliares
'---------------------------------------------------------------
Private Function SeveridadToText(sev As SeveridadInspector) As String
    Select Case sev
        Case sevInfo: SeveridadToText = "INFO"
        Case sevAviso: SeveridadToText = "AVISO"
        Case sevError: SeveridadToText = "ERROR"
    End Select
End Function

Private Function TipoElementoToText(t As TipoElementoInspector) As String
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


