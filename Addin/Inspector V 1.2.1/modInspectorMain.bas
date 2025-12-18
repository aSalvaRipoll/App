Attribute VB_Name = "modInspectorMain"

Option Compare Database
Option Explicit

'===============================================================
' Módulo: modInspectorMain
' Punto de entrada moderno del Inspector VBA
'===============================================================

Private mUltimosResultados As Collection   ' Para RepararProyecto

'---------------------------------------------------------------
' Acceso opcional desde otros módulos
'---------------------------------------------------------------
Public Property Get UltimosResultadosInspector() As Collection
    Set UltimosResultadosInspector = mUltimosResultados
End Property

Public Property Set UltimosResultadosInspector(ByVal Value As Collection)
    Set mUltimosResultados = Value
End Property

'---------------------------------------------------------------
' Ejecutar el Inspector completo
'---------------------------------------------------------------
Public Sub EjecutarInspector()
    Dim cat As clsCatalogoInspector
    Dim resultados As Collection
    Dim r As clsResultadoAnalisis
    Dim t0 As Single

    If Not AsegurarReferenciaVBIDE() Then
        Debug.Print "No se puede ejecutar el Inspector sin VBIDE."
        Exit Sub
    End If

    Debug.Print
    Debug.Print "==============================================="
    Debug.Print "   INSPECTOR DE PROYECTO VBA - INICIO"
    Debug.Print "==============================================="
    Debug.Print

    t0 = Timer

    ' 1. Analizar proyecto
    Set cat = AnalizarProyecto()
    If cat Is Nothing Then
        Debug.Print "Error: no se pudo analizar el proyecto."
        Exit Sub
    End If

    ' 2. Ejecutar reglas
    Set resultados = EjecutarReglas(cat)
    If resultados Is Nothing Then
        Debug.Print "No se generaron resultados."
        Exit Sub
    End If

    ' Guardar para reparaciones posteriores
    Set mUltimosResultados = resultados

    Debug.Print "Resultados encontrados: "; resultados.Count
    Debug.Print String(60, "-")

    ' 3. Mostrar resultados
    For Each r In resultados
        Debug.Print r.Formatear
    Next r

    Debug.Print String(60, "-")
    Debug.Print "Tiempo total: "; Format(Timer - t0, "0.000") & " s"
    Debug.Print "==============================================="
    Debug.Print "   FIN DEL ANÁLISIS"
    Debug.Print "==============================================="
End Sub

'---------------------------------------------------------------
' Reparar el proyecto usando los últimos resultados generados
'---------------------------------------------------------------
Public Sub RepararProyecto()
    If mUltimosResultados Is Nothing Then
        Debug.Print "No hay resultados previos. Ejecuta el Inspector primero."
        Exit Sub
    End If

    RepararResultados mUltimosResultados
End Sub

