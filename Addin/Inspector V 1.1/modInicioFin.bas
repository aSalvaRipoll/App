Attribute VB_Name = "modInicioFin"

Option Compare Database
Option Explicit

'===============================================================
' Módulo: modInicioFin
' Ciclo de vida del Inspector VBA como Add-In ACCDA
'===============================================================

'---------------------------------------------------------------
' Punto de entrada automático del ACCDA (AutoExec)
'---------------------------------------------------------------
Public Sub InicioInspector()
    Debug.Print "Iniciando Inspector VBA..."

    ' 1. Asegurar referencia VBIDE
    If Not AsegurarReferenciaVBIDE() Then
        Debug.Print "No se pudo asegurar VBIDE. El Inspector no funcionará."
        Exit Sub
    End If

    ' 2. Crear menú del IDE
    CrearMenuInspectorVBE

    Debug.Print "Inspector VBA listo."
End Sub

'---------------------------------------------------------------
' Punto de salida opcional (si quieres llamarlo desde un macro)
'---------------------------------------------------------------
Public Sub FinInspector()
    Debug.Print "Cerrando Inspector VBA..."
    EliminarMenuInspectorVBE
End Sub

