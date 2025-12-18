Attribute VB_Name = "modInspectorAutoRepair"

Option Compare Database
Option Explicit

'===============================================================
' Módulo: modInspectorAutoRepair
' Autoreparación de la referencia VBIDE para el Inspector
'===============================================================

Private Const NOMBRE_REF As String = "VBIDE"

'---------------------------------------------------------------
' Función principal: asegurar que VBIDE está activa
'---------------------------------------------------------------
Public Function AsegurarReferenciaVBIDE() As Boolean

    If ReferenciaVBIDEActiva() Then
        AsegurarReferenciaVBIDE = True
        Exit Function
    End If

    If RepararReferenciaVBIDE() Then
        AsegurarReferenciaVBIDE = True
        Exit Function
    End If

    Debug.Print "No se pudo reparar la referencia VBIDE."
    Debug.Print "Actívala manualmente desde:"
    Debug.Print "Herramientas ? Referencias ? Microsoft Visual Basic for Applications Extensibility 5.3"

    AsegurarReferenciaVBIDE = False
End Function

'---------------------------------------------------------------
' ¿La referencia VBIDE está activa y no rota?
'---------------------------------------------------------------
Public Function ReferenciaVBIDEActiva() As Boolean
    Dim ref As Reference

    On Error Resume Next
    For Each ref In Application.VBE.ActiveVBProject.References
        If ref.Name = NOMBRE_REF And Not ref.IsBroken Then
            ReferenciaVBIDEActiva = True
            Exit Function
        End If
    Next ref
End Function

'---------------------------------------------------------------
' Intentar reparar la referencia VBIDE
'---------------------------------------------------------------
Private Function RepararReferenciaVBIDE() As Boolean
    Dim vbProj As VBIDE.VBProject
    Dim ref As Reference
    Dim ruta As String
    Dim rutas() As String
    Dim i As Long

    On Error Resume Next
    Set vbProj = Application.VBE.ActiveVBProject

    ' Eliminar referencia rota
    For Each ref In vbProj.References
        If ref.Name = NOMBRE_REF And ref.IsBroken Then
            vbProj.References.Remove ref
        End If
    Next ref

    ' Probar rutas conocidas
    rutas = RutasVBIDE()

    For i = LBound(rutas) To UBound(rutas)
        ruta = rutas(i)
        If Len(Dir(ruta)) > 0 Then
            vbProj.References.AddFromFile ruta
            If Err.Number = 0 Then
                RepararReferenciaVBIDE = True
                Exit Function
            End If
        End If
    Next i

    RepararReferenciaVBIDE = False
End Function

'---------------------------------------------------------------
' Rutas posibles de VBIDE.dll
'---------------------------------------------------------------
Private Function RutasVBIDE() As String()
    Dim rutas(1 To 6) As String

    rutas(1) = Environ$("CommonProgramFiles") & "\Microsoft Shared\VBA\VBA6\VBIDE.dll"
    rutas(2) = Environ$("CommonProgramFiles(x86)") & "\Microsoft Shared\VBA\VBA6\VBIDE.dll"
    rutas(3) = "C:\Program Files\Common Files\Microsoft Shared\VBA\VBA6\VBIDE.dll"
    rutas(4) = "C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBIDE.dll"
    rutas(5) = Environ$("ProgramFiles") & "\Common Files\Microsoft Shared\VBA\VBA6\VBIDE.dll"
    rutas(6) = Environ$("ProgramFiles(x86)") & "\Common Files\Microsoft Shared\VBA\VBA6\VBIDE.dll"

    RutasVBIDE = rutas
End Function

