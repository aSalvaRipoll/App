Attribute VB_Name = "modReferenciasVBIDE"

Option Compare Database
Option Explicit

'=====================================================
' Comprobación de la referencia VBIDE Extensibility
'=====================================================

Public Function ReferenciaExtensibilidadActiva() As Boolean
    Dim ref As Reference
    On Error Resume Next

    For Each ref In Application.VBE.ActiveVBProject.References
        If ref.Name = "VBIDE" Then
            ReferenciaExtensibilidadActiva = True
            Exit Function
        End If
    Next ref
End Function

'=====================================================
' Activar referencia VBIDE Extensibility 5.3
'=====================================================

Public Sub ActivarReferenciaExtensibilidad()
    Dim vbProj As VBIDE.VBProject
    Dim ruta As String

    On Error GoTo ErrHandler

    Set vbProj = Application.VBE.ActiveVBProject

    ' Si ya está activa, no hacemos nada
    If ReferenciaExtensibilidadActiva() Then
        Debug.Print "? La referencia VBIDE ya está activa."
        Exit Sub
    End If

    ' Rutas típicas de VBIDE.dll según versión de Office
    ruta = Environ$("CommonProgramFiles") & "\Microsoft Shared\VBA\VBA6\VBIDE.dll"
    If Dir(ruta) = "" Then
        ruta = Environ$("CommonProgramFiles") & "\Microsoft Shared\VBA\VBA7\VBIDE.dll"
    End If

    ' Si no existe, no podemos continuar
    If Dir(ruta) = "" Then
        Debug.Print "? No se encontró VBIDE.dll en las rutas estándar."
        GoTo ErrHandler
    End If

    vbProj.References.AddFromFile ruta
    Debug.Print "? Referencia VBIDE activada correctamente."
    Exit Sub

ErrHandler:
    Debug.Print "? No se pudo activar la referencia VBIDE automáticamente."
    Debug.Print "   Actívala manualmente desde:"
    Debug.Print "   Herramientas ? Referencias ?"
    Debug.Print "   Microsoft Visual Basic for Applications Extensibility 5.3"
End Sub

