Attribute VB_Name = "modInicioFin"
Option Compare Database
Option Explicit

'===============================================================
' Módulo: modInicioFin
' Ciclo de vida del Inspector VBA como Add-In ACCDA
'===============================================================

Private iniciado As Boolean

'---------------------------------------------------------------
' Inicialización del Inspector (llamado desde frmInicio)
'---------------------------------------------------------------
Public Sub InicioInspector()
    If iniciado Then Exit Sub
    Debug.Print "Iniciando Inspector VBA..."

    ' 0. Comprobar versión mínima recomendada
    If Not VersionCompatible() Then
        Debug.Print "Versión de Access incompatible. Cancelando inicio."
        Exit Sub
    End If

    ' 1. Comprobar acceso al VBIDE
    If Not AsegurarReferenciaVBIDE() Then
        Debug.Print "No hay acceso al VBIDE. Cancelando inicio."
        Exit Sub
    End If

    ' 2. Crear menú del IDE (idempotente)
    On Error Resume Next
    CrearMenuInspectorVBE
    If Err.Number <> 0 Then
        Debug.Print "Error creando menú del VBE: " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0

    iniciado = True
    Debug.Print "Inspector VBA listo."
End Sub

'---------------------------------------------------------------
' Finalización del Inspector (llamado desde frmInicio al cerrar)
'---------------------------------------------------------------
Public Sub FinInspector()
    If Not iniciado Then Exit Sub
    Debug.Print "Cerrando Inspector VBA..."

    ' 1. Eliminar menú del IDE (idempotente)
    On Error Resume Next
    EliminarMenuInspectorVBE
    If Err.Number <> 0 Then
        Debug.Print "Error eliminando menú del VBE: " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0

    iniciado = False
End Sub


'---------------------------------------------------------------
' Comprobar versión mínima recomendada de Access
'---------------------------------------------------------------
Public Function VersionCompatible() As Boolean
    Dim v As Double
    v = Val(Application.Version)

    ' Access 2016 = 16.0
    ' Access 2013 = 15.0
    ' Access 2010 = 14.0

    If v < 16 Then
        MsgBox "Esta versión de Access es antigua o incompatible." & vbCrLf & _
               "El Inspector VBA requiere Access 2016 o superior.", _
               vbExclamation, "Inspector VBA"
        VersionCompatible = False
    Else
        VersionCompatible = True
    End If
End Function


