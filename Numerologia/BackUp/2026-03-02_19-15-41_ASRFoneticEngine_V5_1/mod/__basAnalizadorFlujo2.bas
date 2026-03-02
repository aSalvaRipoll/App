Attribute VB_Name = "__basAnalizadorFlujo2"

Option Compare Database
Option Explicit

' Diccionarios globales
Private gProcedimientos As Object
Private gVisitados As Object

' ============================================================
' ENTRADA PRINCIPAL
' ============================================================

Public Sub GenerarArbolFlujo(Procedimiento As String)

    Set gProcedimientos = ObtenerTodosLosProcedimientos()
    Set gVisitados = CreateObject("Scripting.Dictionary")

    Dim f As Integer
    f = FreeFile
    Open CurrentProject.Path & "\arbol_flujo.txt" For Output As #f

    Print #f, "Árbol de flujo desde: " & Procedimiento
    Print #f, String(60, "-")

    DibujarNodo Procedimiento, "", True, f

    Close #f
    MsgBox "Árbol generado."

End Sub

' ============================================================
' ÁRBOL ASCII
' ============================================================

Private Sub DibujarNodo(nombre As String, Prefijo As String, Ultimo As Boolean, f As Integer)

    Dim Modulo As String
    Dim cm As VBIDE.CodeModule
    Set cm = BuscarModulo(nombre, Modulo)

    Dim rama As String
    rama = IIf(Ultimo, Prefijo & "+-- ", Prefijo & "+-- ")

    If cm Is Nothing Then
        Print #f, rama & nombre & " (NO ENCONTRADO)"
        Exit Sub
    End If

    If gVisitados.Exists(nombre) Then
        Print #f, rama & nombre & "   [Módulo: " & Modulo & "] (ya visitado)"
        Exit Sub
    End If

    gVisitados.Add nombre, True
    Print #f, rama & nombre & "   [Módulo: " & Modulo & "]"

    Dim llamadas As Collection
    Set llamadas = ObtenerLlamadas(cm, nombre)

    Dim i As Long
    For i = 1 To llamadas.count
        Dim nuevoPrefijo As String
        nuevoPrefijo = IIf(Ultimo, Prefijo & "    ", Prefijo & "¦   ")

        DibujarNodo llamadas(i), nuevoPrefijo, (i = llamadas.count), f
    Next i

End Sub

' ============================================================
' OBTENER TODOS LOS PROCEDIMIENTOS
' ============================================================

Private Function ObtenerTodosLosProcedimientos() As Object

    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")

    Dim comp As VBIDE.VBComponent
    Dim cm As VBIDE.CodeModule
    Dim i As Long, nombre As String

    For Each comp In Application.VBE.ActiveVBProject.VBComponents
        Set cm = comp.CodeModule
        i = 1

        Do While i <= cm.CountOfLines
            nombre = cm.ProcOfLine(i, vbext_pk_Proc)
            If Len(nombre) > 0 Then
                If Not d.Exists(nombre) Then d.Add nombre, True
                i = i + cm.ProcCountLines(nombre, vbext_pk_Proc)
            Else
                i = i + 1
            End If
        Loop
    Next comp

    Set ObtenerTodosLosProcedimientos = d

End Function

' ============================================================
' BUSCAR MÓDULO DE UN PROCEDIMIENTO
' ============================================================

Private Function BuscarModulo(nombre As String, ByRef Modulo As String) As VBIDE.CodeModule

    Dim comp As VBIDE.VBComponent
    Dim cm As VBIDE.CodeModule

    If Not gProcedimientos.Exists(nombre) Then Exit Function

    For Each comp In Application.VBE.ActiveVBProject.VBComponents
        Set cm = comp.CodeModule

        On Error Resume Next
        Dim linea As Long
        linea = cm.ProcStartLine(nombre, vbext_pk_Proc)
        On Error GoTo 0

        If linea > 0 Then
            Modulo = comp.Name
            Set BuscarModulo = cm
            Exit Function
        End If
    Next comp

End Function

' ============================================================
' OBTENER LLAMADAS REALES
' ============================================================

Private Function ObtenerLlamadas(cm As VBIDE.CodeModule, proc As String) As Collection

    Dim col As New Collection
    Dim inicio As Long, fin As Long

    inicio = cm.ProcStartLine(proc, vbext_pk_Proc)
    fin = inicio + cm.ProcCountLines(proc, vbext_pk_Proc) - 1

    Dim i As Long
    For i = inicio To fin
        DetectarLlamadasReales cm.Lines(i, 1), col, proc
    Next i

    Set ObtenerLlamadas = col

End Function

' ============================================================
' DETECTOR DE LLAMADAS REALES (VERSIÓN DEFINITIVA)
' ============================================================

Private Sub DetectarLlamadasReales(linea As String, ByRef col As Collection, ProcActual As String)

    Dim texto As String
    texto = Trim$(linea)

    ' Quitar comentarios
    If InStr(texto, "'") > 0 Then texto = Left$(texto, InStr(texto, "'") - 1)
    If Len(texto) = 0 Then Exit Sub

    ' Quitar cadenas
    Do While InStr(texto, """") > 0
        Dim p1 As Long, p2 As Long
        p1 = InStr(texto, """")
        p2 = InStr(p1 + 1, texto, """")
        If p2 = 0 Then Exit Do
        texto = Left$(texto, p1 - 1) & Mid$(texto, p2 + 1)
    Loop

    Dim nombre As Variant
    For Each nombre In gProcedimientos.Keys

        If nombre = ProcActual Then GoTo siguiente

        ' 1) Call nombre(...)
        If LCase$(texto) Like "*call " & LCase$(nombre) & "*" Then
            col.Add nombre
            GoTo siguiente
        End If

        ' 2) nombre(...)
        If InStr(1, texto, nombre & "(", vbTextCompare) > 0 Then
            col.Add nombre
            GoTo siguiente
        End If

        ' 3) nombre arg1, arg2
        Dim pos As Long
        pos = InStr(1, texto, nombre & " ", vbTextCompare)
        If pos > 0 Then
            Dim resto As String
            resto = Trim$(Mid$(texto, pos + Len(nombre)))
            If Left$(resto, 1) <> "=" Then
                col.Add nombre
                GoTo siguiente
            End If
        End If

siguiente:
    Next nombre

End Sub

