Attribute VB_Name = "__basDetectarHuerfanos_2"

'Option Compare Database
'Option Explicit
'
'Private gProcedimientos As Object
'Private gLlamados As Object
'Private gModulos As Object
'
'' ============================================================
'' ENTRADA PRINCIPAL
'' ============================================================
'Public Sub DetectarProcedimientosHuerfanosPro()
'
'    Set gProcedimientos = Nothing
'    Set gLlamados = Nothing
'    Set gModulos = Nothing
'
'    Set gProcedimientos = CreateObject("Scripting.Dictionary")
'    Set gLlamados = CreateObject("Scripting.Dictionary")
'    Set gModulos = CreateObject("Scripting.Dictionary")
'
'    RegistrarLlamadasEnCodigo
'    RegistrarLlamadasEnMacros
'    RegistrarLlamadasEnFormularios
'
'    Dim f As Integer
'    f = FreeFile
'    Open CurrentProject.Path & "\procedimientos_huerfanos.txt" For Output As #f
'
'    Print #f, "Procedimientos huérfanos (no llamados por ningún otro elemento)"
'    Print #f, String(70, "-")
'
'    Dim proc As Variant
'    For Each proc In gProcedimientos.Keys
'
'        If EsEvento(CStr(proc)) Then GoTo siguiente
'        If EsAPI(CStr(proc)) Then GoTo siguiente
'        If gLlamados.Exists(proc) Then GoTo siguiente
'
'        Print #f, proc & "   (Módulo: " & gModulos(proc) & ")"
'
'siguiente:
'    Next proc
'
'Dim total As Long, usados As Long
'total = gProcedimientos.count
'usados = gLlamados.count
'
'Print #f, "Total procedimientos: " & total
'Print #f, "Marcados como usados: " & usados
'Print #f, ""
'
'    Close #f
'    MsgBox "Análisis completado: procedimientos_huerfanos.txt"
'
'End Sub
'
'' ============================================================
'' REGISTRAR LLAMADAS EN CÓDIGO
'' ============================================================
'
'Private Sub RegistrarLlamadasEnCodigo()
'
'    Dim comp As VBIDE.VBComponent
'    Dim cm As VBIDE.CodeModule
'    Dim nombre As String
'    Dim i As Long
'
'    For Each comp In Application.VBE.ActiveVBProject.VBComponents
'        Set cm = comp.CodeModule
'        i = 1
'
'        Do While i <= cm.CountOfLines
'            nombre = cm.ProcOfLine(i, vbext_pk_Proc)
'
'            If Len(nombre) > 0 Then
'                RegistrarLlamadasEnProcedimiento cm, nombre
'                i = i + cm.ProcCountLines(nombre, vbext_pk_Proc)
'            Else
'                i = i + 1
'            End If
'        Loop
'    Next comp
'
'End Sub
'
'Private Sub RegistrarLlamadasEnProcedimiento(cm As VBIDE.CodeModule, proc As String)
'
'    Dim inicio As Long, fin As Long
'    inicio = cm.ProcStartLine(proc, vbext_pk_Proc)
'    fin = inicio + cm.ProcCountLines(proc, vbext_pk_Proc) - 1
'
'    Dim i As Long
'    For i = inicio To fin
'        DetectarLlamadasReales cm.Lines(i, 1), proc
'    Next i
'
'End Sub
'
'' ============================================================
'' DETECTOR DE LLAMADAS REALES
'' ============================================================
'
'Private Sub DetectarLlamadasReales(linea As String, ProcActual As String)
'
'    Dim texto As String
'    texto = Trim$(linea)
'
'    If InStr(texto, "'") > 0 Then texto = Left$(texto, InStr(texto, "'") - 1)
'    If Len(texto) = 0 Then Exit Sub
'
'    Do While InStr(texto, """") > 0
'        Dim p1 As Long, p2 As Long
'        p1 = InStr(texto, """")
'        p2 = InStr(p1 + 1, texto, """")
'        If p2 = 0 Then Exit Do
'        texto = Left$(texto, p1 - 1) & Mid$(texto, p2 + 1)
'    Loop
'
'    Dim nombre As Variant
'    For Each nombre In gProcedimientos.Keys
'
'        If nombre = ProcActual Then GoTo siguiente
'
'        If LCase$(texto) Like "*call " & LCase$(nombre) & "*" Then
'            gLlamados(nombre) = True
'            GoTo siguiente
'        End If
'
'        If InStr(1, texto, nombre & "(", vbTextCompare) > 0 Then
'            gLlamados(nombre) = True
'            GoTo siguiente
'        End If
'
'        Dim pos As Long
'        pos = InStr(1, texto, nombre & " ", vbTextCompare)
'        If pos > 0 Then
'            Dim resto As String
'            resto = Trim$(Mid$(texto, pos + Len(nombre)))
'            If Left$(resto, 1) <> "=" Then
'                gLlamados(nombre) = True
'                GoTo siguiente
'            End If
'        End If
'
'siguiente:
'    Next nombre
'
'End Sub
'
'' ============================================================
'' REGISTRAR LLAMADAS EN MACROS
'' ============================================================
'
'Private Sub RegistrarLlamadasEnMacros()
'
'    Dim m As AccessObject
'    For Each m In Application.CurrentProject.AllMacros
'
'        Dim ruta As String
'        ruta = Environ$("TEMP") & "\macro_tmp.txt"
'
'        Application.SaveAsText acMacro, m.Name, ruta
'
'        Dim f As Integer
'        f = FreeFile
'        Open ruta For Input As #f
'
'        Dim linea As String
'        Do While Not EOF(f)
'            Line Input #f, linea
'
'            Dim proc As Variant
'            For Each proc In gProcedimientos.Keys
'                If InStr(1, linea, proc, vbTextCompare) > 0 Then
'                    gLlamados(proc) = True
'                End If
'            Next proc
'
'        Loop
'
'        Close #f
'        Kill ruta
'
'    Next m
'
'End Sub
'
'' ============================================================
'' REGISTRAR LLAMADAS EN FORMULARIOS
'' ============================================================
'
'Private Sub RegistrarLlamadasEnFormularios()
'
'    Dim frm As AccessObject
'    For Each frm In Application.CurrentProject.AllForms
'
'        Dim ruta As String
'        ruta = Environ$("TEMP") & "\form_tmp.txt"
'
'        Application.SaveAsText acForm, frm.Name, ruta
'
'        Dim f As Integer
'        f = FreeFile
'        Open ruta For Input As #f
'
'        Dim linea As String
'        Do While Not EOF(f)
'            Line Input #f, linea
'
'            Dim proc As Variant
'            For Each proc In gProcedimientos.Keys
'
'                If InStr(1, linea, "=" & proc & "(", vbTextCompare) > 0 Then
'                    gLlamados(proc) = True
'                End If
'
'                If InStr(1, linea, "RunCode", vbTextCompare) > 0 _
'                And InStr(1, linea, proc, vbTextCompare) > 0 Then
'                    gLlamados(proc) = True
'                End If
'
'            Next proc
'
'        Loop
'
'        Close #f
'        Kill ruta
'
'    Next frm
'
'End Sub
'
'' ============================================================
'' OBTENER TODOS LOS PROCEDIMIENTOS
'' ============================================================
'Private Function ObtenerTodosLosProcedimientos() As Object
'
'    ' Asegurar inicialización
'    If gModulos Is Nothing Then Set gModulos = CreateObject("Scripting.Dictionary")
'
'    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
'
'    Dim comp As VBIDE.VBComponent
'    Dim cm As VBIDE.CodeModule
'    Dim i As Long, nombre As String
'
'    For Each comp In Application.VBE.ActiveVBProject.VBComponents
'        Set cm = comp.CodeModule
'        i = 1
'
'        Do While i <= cm.CountOfLines
'            nombre = cm.ProcOfLine(i, vbext_pk_Proc)
'
'            If Len(nombre) > 0 Then
'
'                If Not d.Exists(nombre) Then
'                    d.Add nombre, True
'
'                    ' Guardar módulo
'                    If Not gModulos.Exists(nombre) Then
'                        gModulos.Add nombre, comp.Name
'                    End If
'
'                End If
'
'                i = i + cm.ProcCountLines(nombre, vbext_pk_Proc)
'            Else
'                i = i + 1
'            End If
'
'        Loop
'    Next comp
'
'    Set ObtenerTodosLosProcedimientos = d
'
'End Function
'
''Private Function ObtenerTodosLosProcedimientos() As Object
''
''    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
''
''    Dim comp As VBIDE.VBComponent
''    Dim cm As VBIDE.CodeModule
''    Dim i As Long, nombre As String
''
''    If gModulos Is Nothing Then Set gModulos = CreateObject("Scripting.Dictionary")
''
''    For Each comp In Application.VBE.ActiveVBProject.VBComponents
''        Set cm = comp.CodeModule
''        i = 1
''
''        Do While i <= cm.CountOfLines
''            nombre = cm.ProcOfLine(i, vbext_pk_Proc)
''            If Len(nombre) > 0 Then
''                If Not d.Exists(nombre) Then
''                    d.Add nombre, True
''                    gModulos(nombre) = comp.Name
''                End If
''                i = i + cm.ProcCountLines(nombre, vbext_pk_Proc)
''            Else
''                i = i + 1
''            End If
''        Loop
''    Next comp
''
''    Set ObtenerTodosLosProcedimientos = d
''
''End Function
'
'' ============================================================
'' EXCLUSIONES
'' ============================================================
'
'Private Function EsEvento(nombre As String) As Boolean
'    EsEvento = (InStr(1, nombre, "_") > 0)
'End Function
'
'Private Function EsAPI(nombre As String) As Boolean
'
'    Dim comp As VBIDE.VBComponent
'    Dim cm As VBIDE.CodeModule
'    Dim i As Long
'
'    For Each comp In Application.VBE.ActiveVBProject.VBComponents
'        Set cm = comp.CodeModule
'
'        For i = 1 To cm.CountOfLines
'            Dim linea As String
'            linea = Trim$(cm.Lines(i, 1))
'
'            If LCase$(linea) Like "declare *" Then
'                If InStr(1, linea, nombre, vbTextCompare) > 0 Then
'                    EsAPI = True
'                    Exit Function
'                End If
'            End If
'        Next i
'    Next comp
'
'End Function


