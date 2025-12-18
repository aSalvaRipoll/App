Attribute VB_Name = "modExportaInspector_v2"
Option Compare Database

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
    Dim xl As Object
    Dim wb As Object
    Dim ws As Object
    Dim r As clsResultadoAnalisis
    Dim fila As Long

    If resultados Is Nothing Then Exit Sub
    If resultados.Count = 0 Then Exit Sub

    Set xl = CreateObject("Excel.Application")
    Set wb = xl.Workbooks.Add
    Set ws = wb.Sheets(1)

    ' Encabezados
    ws.Cells(1, 1).Value = "Código"
    ws.Cells(1, 2).Value = "Severidad"
    ws.Cells(1, 3).Value = "Tipo"
    ws.Cells(1, 4).Value = "Elemento"
    ws.Cells(1, 5).Value = "Miembro"
    ws.Cells(1, 6).Value = "Línea"
    ws.Cells(1, 7).Value = "Descripción"
    ws.Cells(1, 8).Value = "Detalles"

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
        fila = fila + 1
    Next r

    wb.SaveAs ruta
    wb.Close False
    xl.Quit
End Sub

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
        Print #f, _
            sim.nombre & " | " & _
            sim.categoria & " | " & _
            sim.modulo & " | " & _
            sim.miembro & " | " & _
            sim.LineaDeclaracion & " | " & _
            sim.TipoTexto & " | " & _
            IIf(sim.Usado, "Sí", "No")
    Next sim

    Close #f
End Sub

'---------------------------------------------------------------
' Exporta símbolos no usados directamente a Excel
'---------------------------------------------------------------
Public Sub ExportarSimbolosNoUsadosExcel(ByVal ruta As String)
    Dim col As Collection
    Dim sim As clsSimbolo
    Dim xl As Object
    Dim wb As Object
    Dim ws As Object
    Dim fila As Long

    If gCatalogoSimbolos Is Nothing Then Exit Sub

    Set col = gCatalogoSimbolos.SimbolosNoUsados
    If col.Count = 0 Then Exit Sub

    Set xl = CreateObject("Excel.Application")
    Set wb = xl.Workbooks.Add
    Set ws = wb.Sheets(1)

    ws.Cells(1, 1).Value = "Nombre"
    ws.Cells(1, 2).Value = "Categoría"
    ws.Cells(1, 3).Value = "Módulo"
    ws.Cells(1, 4).Value = "Miembro"
    ws.Cells(1, 5).Value = "Línea"
    ws.Cells(1, 6).Value = "Tipo"
    ws.Cells(1, 7).Value = "Usado"

    fila = 2

    For Each sim In col
        ws.Cells(fila, 1).Value = sim.nombre
        ws.Cells(fila, 2).Value = sim.categoria
        ws.Cells(fila, 3).Value = sim.modulo
        ws.Cells(fila, 4).Value = sim.miembro
        ws.Cells(fila, 5).Value = sim.LineaDeclaracion
        ws.Cells(fila, 6).Value = sim.TipoTexto
        ws.Cells(fila, 7).Value = IIf(sim.Usado, "Sí", "No")
        fila = fila + 1
    Next sim

    wb.SaveAs ruta
    wb.Close False
    xl.Quit
End Sub

'---------------------------------------------------------------
' Exportación genérica para cualquier colección con ToTextFile
'---------------------------------------------------------------
Public Sub ExportarColeccionTexto(col As Collection, ByVal ruta As String)
    Dim f As Integer
    Dim o As Object

    If col Is Nothing Then Exit Sub
    If col.Count = 0 Then Exit Sub

    f = FreeFile
    Open ruta For Output As #f

    For Each o In col
        If HasMethod(o, "ToTextFile") Then
            Print #f, o.ToTextFile
        Else
            Print #f, CStr(o)
        End If
    Next o

    Close #f
End Sub

Private Function HasMethod(obj As Object, methodName As String) As Boolean
    On Error Resume Next
    Dim t As Object
    Set t = CallByName(obj, methodName, VbGet)
    HasMethod = (Err.Number = 0)
    Err.Clear
End Function


