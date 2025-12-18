Attribute VB_Name = "modExportaInspector"
Option Compare Database

Option Compare Database
Option Explicit

'===============================================================
' Módulo: modExportaInspector
' Exportación de resultados y símbolos a archivos de texto
'===============================================================

'---------------------------------------------------------------
' Exporta una colección de clsResultadoAnalisis a un archivo
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
' Exporta símbolos no usados
'---------------------------------------------------------------
Public Sub ExportarSimbolosNoUsados(ByVal ruta As String)
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
' Exporta cualquier colección de objetos con método ToTextFile
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

'---------------------------------------------------------------
' Detecta si un objeto tiene un método concreto
'---------------------------------------------------------------
Private Function HasMethod(obj As Object, methodName As String) As Boolean
    On Error Resume Next
    Dim t As Object
    Set t = CallByName(obj, methodName, VbGet)
    HasMethod = (Err.Number = 0)
    Err.Clear
End Function


