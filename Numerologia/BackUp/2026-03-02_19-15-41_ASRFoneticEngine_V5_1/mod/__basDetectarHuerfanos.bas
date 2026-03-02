Attribute VB_Name = "__basDetectarHuerfanos"

Option Compare Database
Option Explicit

'' Diccionarios globales
'Private gProcedimientos As Object
'Private gLlamados As Object
'
'' ============================================================
'' ENTRADA PRINCIPAL
'' ============================================================
'
'Public Sub DetectarProcedimientosHuerfanos()
'
'    Set gProcedimientos = ObtenerTodosLosProcedimientos()
'    Set gLlamados = CreateObject("Scripting.Dictionary")
'
'    Dim comp As VBIDE.VBComponent
'    Dim cm As VBIDE.CodeModule
'    Dim nombre As String
'    Dim i As Long
'
'    ' Recorrer todo el proyecto y registrar llamadas
'    For Each comp In Application.VBE.ActiveVBProject.VBComponents
'        Set cm = comp.CodeModule
'        i = 1
'
'        Do While i <= cm.CountOfLines
'            nombre = cm.ProcOfLine(i, vbext_pk_Proc)
'
'            If Len(nombre) > 0 Then
'                RegistrarLlamadas cm, nombre
'                i = i + cm.ProcCountLines(nombre, vbext_pk_Proc)
'            Else
'                i = i + 1
'            End If
'        Loop
'    Next comp
'
'    ' Generar archivo de salida
'    Dim f As Integer
'    f = FreeFile
'    Open CurrentProject.Path & "\procedimientos_huerfanos.txt" For Output As #f
'
'    Print #f, "Procedimientos huérfanos (no llamados por ningún otro procedimiento)"
'    Print #f, String(60, "-")
'
'    Dim proc As Variant
'    For Each proc In gProcedimientos.Keys
'        If Not gLlamados.Exists(proc) Then
'            Print #f, proc
'        End If
'    Next proc
'
'    Close #f
'    MsgBox "Análisis completado: procedimientos_huerfanos.txt"
'
'End Sub
'
'' ============================================================
'' REGISTRAR LLAMADAS
'' ============================================================
'
'Private Sub RegistrarLlamadas(cm As VBIDE.CodeModule, proc As String)
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
'' DETECTOR DE LLAMADAS REALES (MISMO QUE EL DEL ÁRBOL)
'' ============================================================
'
'Private Sub DetectarLlamadasReales(linea As String, ProcActual As String)
'
'    Dim texto As String
'    texto = Trim$(linea)
'
'    ' Quitar comentarios
'    If InStr(texto, "'") > 0 Then texto = Left$(texto, InStr(texto, "'") - 1)
'    If Len(texto) = 0 Then Exit Sub
'
'    ' Quitar cadenas
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
'        ' 1) Call nombre(...)
'        If LCase$(texto) Like "*call " & LCase$(nombre) & "*" Then
'            gLlamados(nombre) = True
'            GoTo siguiente
'        End If
'
'        ' 2) nombre(...)
'        If InStr(1, texto, nombre & "(", vbTextCompare) > 0 Then
'            gLlamados(nombre) = True
'            GoTo siguiente
'        End If
'
'        ' 3) nombre arg1, arg2
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
'' OBTENER TODOS LOS PROCEDIMIENTOS
'' ============================================================
'
'Private Function ObtenerTodosLosProcedimientos() As Object
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
'            If Len(nombre) > 0 Then
'                If Not d.Exists(nombre) Then d.Add nombre, True
'                i = i + cm.ProcCountLines(nombre, vbext_pk_Proc)
'            Else
'                i = i + 1
'            End If
'        Loop
'    Next comp
'
'    Set ObtenerTodosLosProcedimientos = d
'
'End Function
'
