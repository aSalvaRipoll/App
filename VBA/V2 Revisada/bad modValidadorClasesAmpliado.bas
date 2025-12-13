Attribute VB_Name = "bad modValidadorClasesAmpliado"
'
'Option Compare Database
'Option Explicit
'
'' ============================================================
'' VERIFICADOR DE CLASES PARA ACCESS 2019
'' ============================================================
'' Revisa:
''   - Que todas las clases sean realmente módulos de clase
''   - Que el nombre interno VB_Name coincida
''   - Que no haya duplicados
''   - Que no haya clases privadas
''   - Que todas compilen
''   - Que todas puedan instanciarse (si procede)
'' ============================================================
'
'Public Sub VerificarClases()
'    Dim comp As VBIDE.VBComponent
'    Dim dictNombres As Object
'    Dim nombre As String
'    Dim tipo As String
'    Dim errores As String
'    Dim puedeInstanciar As Boolean
'
'    Set dictNombres = CreateObject("Scripting.Dictionary")
'    errores = ""
'
'    Debug.Print "=== VERIFICACIÓN DE CLASES (Access 2019) ==="
'
'    For Each comp In Application.VBE.ActiveVBProject.VBComponents
'
'        Select Case comp.Type
'
'            Case vbext_ct_ClassModule
'                tipo = "Clase"
'
'                ' Nombre interno
'                nombre = ObtenerVBName(comp)
'
'                If Len(nombre) = 0 Then
'                    errores = errores & "ERROR: Clase sin VB_Name ? " & comp.Name & vbCrLf
'                End If
'
'                ' Duplicados
'                If dictNombres.Exists(nombre) Then
'                    errores = errores & "ERROR: Nombre de clase duplicado ? " & nombre & vbCrLf
'                Else
'                    dictNombres.Add nombre, True
'                End If
'
'                ' Verificar si es PublicNotCreatable
'                If Not EsPublicNotCreatable(comp) Then
'                    errores = errores & "ADVERTENCIA: Clase no es PublicNotCreatable ? " & nombre & vbCrLf
'                End If
'
'                ' Intentar instanciar (si no tiene PredeclaredId)
'                puedeInstanciar = PuedeInstanciarClase(nombre)
'                If Not puedeInstanciar Then
'                    errores = errores & "ERROR: No se puede instanciar ? " & nombre & vbCrLf
'                End If
'
'                Debug.Print "Clase OK ? "; nombre
'
'            Case vbext_ct_StdModule
'                tipo = "Módulo estándar"
'                Debug.Print "Módulo estándar ? "; comp.Name
'
'            Case Else
'                ' Formularios, informes, etc.
'        End Select
'
'    Next comp
'
'    Debug.Print "=== FIN DE VERIFICACIÓN ==="
'
'    If Len(errores) > 0 Then
'        Debug.Print vbCrLf & "=== ERRORES DETECTADOS ==="
'        Debug.Print errores
'        MsgBox "Se detectaron problemas. Revisa la ventana Inmediato.", vbExclamation
'    Else
'        MsgBox "Todas las clases están correctas.", vbInformation
'    End If
'End Sub
'
'' ============================================================
'' OBTENER NOMBRE INTERNO VB_Name
'' ============================================================
'Private Function ObtenerVBName(comp As VBIDE.VBComponent) As String
'    Dim linea As String
'    Dim i As Long
'
'    For i = 1 To comp.CodeModule.CountOfLines
'        linea = comp.CodeModule.Lines(i, 1)
'        If InStr(1, linea, "Attribute VB_Name", vbTextCompare) > 0 Then
'            ObtenerVBName = ExtraerNombreVB(linea)
'            Exit Function
'        End If
'    Next i
'End Function
'
'Private Function ExtraerNombreVB(linea As String) As String
'    Dim p As Long
'    p = InStr(linea, "=")
'    If p > 0 Then
'        ExtraerNombreVB = Trim(Replace(Mid(linea, p + 1), """", ""))
'    End If
'End Function
'
'' ============================================================
'' VERIFICAR SI ES PUBLIC NOT CREATABLE
'' ============================================================
'Private Function EsPublicNotCreatable(comp As VBIDE.VBComponent) As Boolean
'    Dim linea As String
'    Dim i As Long
'
'    For i = 1 To comp.CodeModule.CountOfLines
'        linea = comp.CodeModule.Lines(i, 1)
'
'        If InStr(1, linea, "MultiUse", vbTextCompare) > 0 Then
'            If InStr(1, linea, "-1", vbTextCompare) > 0 Then
'                EsPublicNotCreatable = True
'            Else
'                EsPublicNotCreatable = False
'            End If
'            Exit Function
'        End If
'    Next i
'
'    EsPublicNotCreatable = False
'End Function
'
'' ============================================================
'' PROBAR SI SE PUEDE INSTANCIAR UNA CLASE
'' ============================================================
'Private Function PuedeInstanciarClase(nombreClase As String) As Boolean
'    On Error GoTo ErrHandler
'
'    Dim obj As Object
'    Set obj = CreateObject("Access.Application").VBE.ActiveVBProject.VBComponents(nombreClase)
'
'    ' Si llega aquí, la clase existe en el proyecto
'    ' Ahora probamos instanciación real
'    Set obj = Nothing
'    Set obj = NewInstance(nombreClase)
'
'    PuedeInstanciarClase = True
'    Exit Function
'
'ErrHandler:
'    PuedeInstanciarClase = False
'End Function
'
'Private Function NewInstance(nombreClase As String) As Object
'    ' Instancia dinámica
'    Set NewInstance = VBA.CreateObject("Access." & nombreClase)
'End Function
'
