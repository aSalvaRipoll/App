Attribute VB_Name = "modReferenciasVBIDE_old"

Option Compare Database
Option Explicit

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

Public Sub ActivarReferenciaExtensibilidad()
    Dim vbProj As VBIDE.VBProject
    Dim ref As Reference
    Dim ruta As String
    
    On Error GoTo ErrHandler
    
    Set vbProj = Application.VBE.ActiveVBProject
    
    ' Si ya está activa, no hacemos nada
    If ReferenciaExtensibilidadActiva() Then
        Debug.Print "? La referencia ya está activa."
        Exit Sub
    End If
    
    ' Ruta típica de VBIDE.dll
    ruta = Environ$("CommonProgramFiles") & "\Microsoft Shared\VBA\VBA6\VBIDE.dll"
    
    vbProj.References.AddFromFile ruta
    
    Debug.Print "? Referencia activada correctamente."
    Exit Sub

ErrHandler:
    Debug.Print "? No se pudo activar la referencia automáticamente."
    Debug.Print "   Actívala manualmente desde:"
    Debug.Print "   Herramientas ? Referencias ? Microsoft Visual Basic for Applications Extensibility 5.3"
End Sub

