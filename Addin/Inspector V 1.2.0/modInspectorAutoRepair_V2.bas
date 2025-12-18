Attribute VB_Name = "modInspectorAutoRepair_V2"

Option Compare Database
Option Explicit

'===============================================================
' Módulo: modInspectorAutoRepair
' Verificación segura del acceso al VBIDE para el Inspector
'===============================================================

'---------------------------------------------------------------
' Función principal: asegurar que VBIDE está disponible
'---------------------------------------------------------------
Public Function AsegurarReferenciaVBIDE() As Boolean
    On Error GoTo ErrHandler

    ' 1. Comprobar si el acceso al modelo de objetos VBA está habilitado
    If Not AccesoVBIDEHabilitado() Then
        Debug.Print "Inspector: Acceso al modelo de objetos VBA deshabilitado."
        AsegurarReferenciaVBIDE = False
        Exit Function
    End If

    ' 2. Comprobar si Application.VBE responde correctamente
    Dim vbProj As Object

    ' Nota: ActiveVBProject puede fallar si el proyecto está protegido
    Set vbProj = Application.VBE.ActiveVBProject

    ' Si llegamos aquí, todo está bien
    AsegurarReferenciaVBIDE = True
    Exit Function

ErrHandler:
    Debug.Print "Inspector: No se pudo acceder al VBIDE (" & Err.Number & "): " & Err.Description
    AsegurarReferenciaVBIDE = False
End Function

'---------------------------------------------------------------
' ¿Está habilitado el acceso al modelo de objetos VBA?
'---------------------------------------------------------------
Private Function AccesoVBIDEHabilitado() As Boolean
    On Error Resume Next

    ' Si falla, el acceso está deshabilitado en:
    ' Archivo ? Opciones ? Centro de confianza ? Configuración del Centro de confianza ?
    ' Configuración de macros ? “Confiar en el acceso al modelo de objetos de proyecto VBA”
    Dim test As Object
    Set test = Application.VBE

    AccesoVBIDEHabilitado = (Err.Number = 0)
End Function

