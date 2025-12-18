Attribute VB_Name = "modInspectorAutoRepair"

Option Compare Database
Option Explicit

Private Const NOMBRE_REF As String = "VBIDE"

Public Function AsegurarReferenciaVBIDE() As Boolean
    Debug.Print "Comprobando referencia VBIDE..."
    
    If ReferenciaVBIDEActiva() Then
        Debug.Print "? VBIDE está activa."
        AsegurarReferenciaVBIDE = True
        Exit Function
    End If
    
    Debug.Print "?? VBIDE no está activa. Intentando repararla..."
    
    If RepararReferenciaVBIDE() Then
        Debug.Print "? VBIDE reparada correctamente."
        AsegurarReferenciaVBIDE = True
        Exit Function
    End If
    
    Debug.Print "? No se pudo reparar automáticamente."
    Debug.Print "   Actívala manualmente desde:"
    Debug.Print "   Herramientas ? Referencias ? Microsoft Visual Basic for Applications Extensibility 5.3"
    
    AsegurarReferenciaVBIDE = False
End Function

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

Private Function RepararReferenciaVBIDE() As Boolean
    Dim vbProj As VBIDE.VBProject
    Dim ref As Reference
    Dim ruta As String
    
    On Error Resume Next
    Set vbProj = Application.VBE.ActiveVBProject
    
    For Each ref In vbProj.References
        If ref.Name = NOMBRE_REF And ref.IsBroken Then
            Debug.Print "   - Eliminando referencia rota..."
            vbProj.References.Remove ref
        End If
    Next ref
    
    Dim rutas() As String
    rutas = RutasVBIDE()
    
    Dim i As Long
    For i = LBound(rutas) To UBound(rutas)
        ruta = rutas(i)
        If Len(Dir(ruta)) > 0 Then
            Debug.Print "   - Probando ruta: " & ruta
            On Error Resume Next
            vbProj.References.AddFromFile ruta
            If Err.Number = 0 Then
                RepararReferenciaVBIDE = True
                Exit Function
            End If
        End If
    Next i
    
    RepararReferenciaVBIDE = False
End Function

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



'Option Compare Database
'Option Explicit
'
'' ============================================================
''   SISTEMA DE AUTOREPARACIÓN DE LA REFERENCIA VBIDE
'' ============================================================
'' Este módulo:
''   - Detecta si la referencia VBIDE está activa
''   - Detecta si está rota
''   - Intenta repararla automáticamente
''   - Busca VBIDE.dll en varias rutas
''   - Evita que el Inspector falle
'' ============================================================
'
'' Nombre oficial de la referencia
'Private Const NOMBRE_REF As String = "VBIDE"
'
'' ============================================================
''   FUNCIÓN PRINCIPAL: GARANTIZAR REFERENCIA
'' ============================================================
'Public Function AsegurarReferenciaVBIDE() As Boolean
'    Debug.Print "Comprobando referencia VBIDE…"
'
'    If ReferenciaVBIDEActiva() Then
'        Debug.Print "? VBIDE está activa."
'        AsegurarReferenciaVBIDE = True
'        Exit Function
'    End If
'
'    Debug.Print "?? VBIDE no está activa. Intentando repararla…"
'
'    If RepararReferenciaVBIDE() Then
'        Debug.Print "? VBIDE reparada correctamente."
'        AsegurarReferenciaVBIDE = True
'        Exit Function
'    End If
'
'    Debug.Print "? No se pudo reparar automáticamente."
'    Debug.Print "   Actívala manualmente desde:"
'    Debug.Print "   Herramientas ? Referencias ? Microsoft Visual Basic for Applications Extensibility 5.3"
'
'    AsegurarReferenciaVBIDE = False
'End Function
'
'' ============================================================
''   DETECTAR SI LA REFERENCIA ESTÁ ACTIVA
'' ============================================================
'Public Function ReferenciaVBIDEActiva() As Boolean
'    Dim ref As Reference
'    On Error Resume Next
'
'    For Each ref In Application.VBE.ActiveVBProject.References
'        If ref.Name = NOMBRE_REF And Not ref.IsBroken Then
'            ReferenciaVBIDEActiva = True
'            Exit Function
'        End If
'    Next ref
'End Function
'
'' ============================================================
''   REPARAR REFERENCIA
'' ============================================================
'Private Function RepararReferenciaVBIDE() As Boolean
'    Dim vbProj As VBIDE.VBProject
'    Dim ref As Reference
'    Dim ruta As String
'
'    On Error Resume Next
'    Set vbProj = Application.VBE.ActiveVBProject
'
'    ' 1. Si existe pero está rota, eliminarla
'    For Each ref In vbProj.References
'        If ref.Name = NOMBRE_REF And ref.IsBroken Then
'            Debug.Print "   - Eliminando referencia rota…"
'            vbProj.References.Remove ref
'        End If
'    Next ref
'
'    ' 2. Intentar agregar desde rutas conocidas
'    Dim rutas() As String
'    rutas = RutasVBIDE()
'
'    Dim i As Long
'    For i = LBound(rutas) To UBound(rutas)
'        ruta = rutas(i)
'        If Len(Dir(ruta)) > 0 Then
'            Debug.Print "   - Probando ruta: " & ruta
'            On Error Resume Next
'            vbProj.References.AddFromFile ruta
'            If Err.Number = 0 Then
'                RepararReferenciaVBIDE = True
'                Exit Function
'            End If
'        End If
'    Next i
'
'    RepararReferenciaVBIDE = False
'End Function
'
'' ============================================================
''   RUTAS POSIBLES DE VBIDE.DLL
'' ============================================================
'Private Function RutasVBIDE() As String()
'    Dim rutas(1 To 6) As String
'
'    rutas(1) = Environ$("CommonProgramFiles") & "\Microsoft Shared\VBA\VBA6\VBIDE.dll"
'    rutas(2) = Environ$("CommonProgramFiles(x86)") & "\Microsoft Shared\VBA\VBA6\VBIDE.dll"
'    rutas(3) = "C:\Program Files\Common Files\Microsoft Shared\VBA\VBA6\VBIDE.dll"
'    rutas(4) = "C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBIDE.dll"
'    rutas(5) = Environ$("ProgramFiles") & "\Common Files\Microsoft Shared\VBA\VBA6\VBIDE.dll"
'    rutas(6) = Environ$("ProgramFiles(x86)") & "\Common Files\Microsoft Shared\VBA\VBA6\VBIDE.dll"
'
'    RutasVBIDE = rutas
'End Function
'
