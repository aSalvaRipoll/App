Attribute VB_Name = "31_modAutoRepair"

Option Compare Database
Option Explicit

'===============================================================
' Módulo: 31_modAutoRepair
' Verificación segura del acceso al VBIDE para el Inspector
'===============================================================

' Estado global del VBIDE (se actualiza al arrancar el Inspector)
' Debe declararse en 00_modGlobal:
' Public gVBIDEDisponible As Boolean

'---------------------------------------------------------------
' Función principal: asegurar que VBIDE está disponible
'---------------------------------------------------------------
Public Function VerificaVBIDEinterno() As Boolean
    On Error GoTo ErrHandler

    ' 1. Comprobar si el acceso al modelo de objetos VBA está habilitado
    If Not AccesoVBIDEHabilitado() Then
        Debug.Print "Inspector: Acceso al modelo de objetos VBA deshabilitado."
        gVBIDEDisponible = False
        VerificaVBIDEinterno = False
        Exit Function
    End If

    ' 2. Comprobar si Application.VBE responde correctamente
    Dim vbProj As Object
    Set vbProj = Application.VBE.ActiveVBProject

    ' Si llegamos aquí, todo está bien
    gVBIDEDisponible = True
    VerificaVBIDEinterno = True
    Exit Function

ErrHandler:
    Debug.Print "Inspector: Error accediendo al VBIDE (" & Err.Number & "): " & Err.Description

    ' Diagnóstico adicional
    Select Case Err.Number
        Case 1004, 91
            Debug.Print "Inspector: El proyecto activo podría estar protegido o no disponible."
        Case Else
            Debug.Print "Inspector: Error inesperado al acceder al IDE."
    End Select

    gVBIDEDisponible = False
    VerificaVBIDEinterno = False
End Function


'---------------------------------------------------------------
' ¿Está habilitado el acceso al modelo de objetos VBA?
'---------------------------------------------------------------
Private Function AccesoVBIDEHabilitado() As Boolean
    On Error Resume Next

    Dim test As Object
    Set test = Application.VBE

    AccesoVBIDEHabilitado = (Err.Number = 0)
End Function

