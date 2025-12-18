Attribute VB_Name = "modInspectorDebug_v2"

Option Compare Database
Option Explicit

'===============================================================
' Módulo: modInspectorDebug
' Punto de entrada manual para depuración del Inspector
'===============================================================

Public Sub EjecutarInspectorDebug()
    Dim cat As clsCatalogoInspector
    Dim resultados As Collection
    Dim r As clsResultadoAnalisis
    Dim t0 As Single

    Debug.Print
    Debug.Print "==============================================="
    Debug.Print "   INSPECTOR DE PROYECTO VBA (DEBUG)"
    Debug.Print "==============================================="
    Debug.Print

    ' 1. Asegurar VBIDE
    If Not AsegurarReferenciaVBIDE() Then
        Debug.Print "No se puede ejecutar el Inspector sin VBIDE."
        Exit Sub
    End If

    t0 = Timer

    ' 2. Analizar proyecto
    Set cat = AnalizarProyecto()
    If cat Is Nothing Then
        Debug.Print "Error: no se pudo analizar el proyecto."
        Exit Sub
    End If

    ' 3. Ejecutar reglas
    Set resultados = EjecutarReglas(cat)

    ' 4. Mostrar resultados
    Debug.Print "Resultados encontrados: "; resultados.Count
    Debug.Print String(60, "-")

    For Each r In resultados
        Debug.Print r.Formatear
    Next r

    Debug.Print String(60, "-")
    Debug.Print "Tiempo total: "; Format(Timer - t0, "0.000") & " s"
    Debug.Print "==============================================="
    Debug.Print "   FIN DEL ANÁLISIS (DEBUG)"
    Debug.Print "==============================================="
End Sub

