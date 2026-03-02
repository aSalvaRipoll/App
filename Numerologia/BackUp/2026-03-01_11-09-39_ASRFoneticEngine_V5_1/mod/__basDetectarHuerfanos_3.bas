Attribute VB_Name = "__basDetectarHuerfanos_3"

Option Compare Database
Option Explicit

Private gProcedimientos As Object
Private gLlamados As Object
Private gModulos As Object

' ============================================================
' ENTRADA PRINCIPAL
' ============================================================

Public Sub DetectarProcedimientosHuerfanosPro()

    ' Inicialización segura
    Set gProcedimientos = CreateObject("Scripting.Dictionary")
    Set gLlamados = CreateObject("Scripting.Dictionary")
    Set gModulos = CreateObject("Scripting.Dictionary")

    ' 1. Obtener todos los procedimientos
    Set gProcedimientos = ObtenerTodosLosProcedimientos()

    ' 2. Registrar llamadas reales en código
    RegistrarLlamadasEnCodigo

    ' 3. Registrar llamadas desde macros
    RegistrarLlamadasEnMacros

    ' 4. Registrar llamadas desde formularios
    RegistrarLlamadasEnFormularios

    ' 5. Generar salida
    Dim f As Integer
    f = FreeFile
    Open CurrentProject.Path & "\procedimientos_huerfanos.txt" For Output As #f

    Print #f, "Procedimientos huérfanos (no llamados por ningún otro elemento)"
    Print #f, String(70, "-")

    Dim proc As Variant
    For Each proc In gProcedimientos.Keys

        If EsEvento(CStr(proc)) Then GoTo siguiente
        If EsAPI(CStr(proc)) Then GoTo siguiente
        If gLlamados.Exists(proc) Then GoTo siguiente

        Print #f, proc & "   (Módulo: " & gModulos(proc) & ")"

siguiente:
    Next proc

    Close #f
    MsgBox "Análisis completado: procedimientos_huerfanos.txt"

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

                If Not d.Exists(nombre) Then
                    d.Add nombre, True
                    gModulos(nombre) = comp.Name
                End If

                i = i + cm.ProcCountLines(nombre, vbext_pk_Proc)
            Else
                i = i + 1
            End If

        Loop
    Next comp

    Set ObtenerTodosLosProcedimientos = d

End Function

' ============================================================
' REGISTRAR LLAMADAS EN CÓDIGO
' ============================================================

Private Sub RegistrarLlamadasEnCodigo()

    Dim comp As VBIDE.VBComponent
    Dim cm As VBIDE.CodeModule
    Dim nombre As String
    Dim i As Long

    For Each comp In Application.VBE.ActiveVBProject.VBComponents
        Set cm = comp.CodeModule
        i = 1

        Do While i <= cm.CountOfLines
            nombre = cm.ProcOfLine(i, vbext_pk_Proc)

            If Len(nombre) > 0 Then
                RegistrarLlamadasEnProcedimiento cm, nombre
                i = i + cm.ProcCountLines(nombre, vbext_pk_Proc)
            Else
                i = i + 1
            End If
        Loop
    Next comp

End Sub

Private Sub RegistrarLlamadasEnProcedimiento(cm As VBIDE.CodeModule, proc As String)

    Dim inicio As Long, fin As Long
    inicio = cm.ProcStartLine(proc, vbext_pk_Proc)
    fin = inicio + cm.ProcCountLines(proc, vbext_pk_Proc) - 1

    Dim i As Long
    For i = inicio To fin
        DetectarLlamadasReales cm.Lines(i, 1), proc
    Next i

End Sub

' ============================================================
' DETECTOR DE LLAMADAS REALES (BLINDADO)
' ============================================================
Private Sub DetectarLlamadasReales(linea As String, ProcActual As String)

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

        ' Evitar métodos de objetos: obj.Nombre
        'If texto Like "*." & nombre & "*" Then GoTo siguiente
        
        ' Evitar métodos de objetos externos, pero permitir métodos de clases propias
        If texto Like "*." & nombre & "*" Then
            ' Si el nombre está en nuestros procedimientos, es un método de clase ? permitirlo
            ' Si no, ignorarlo
            If Not gProcedimientos.Exists(nombre) Then GoTo siguiente
        End If


        ' Evitar comparaciones: Nombre = x
        If texto Like "* " & nombre & "=" Then GoTo siguiente
        If texto Like "*=" & nombre & "*" Then GoTo siguiente

        ' --- DETECCIÓN ROBUSTA DE LLAMADAS ---

        ' 1) Call Nombre(...)
        If LCase$(texto) Like "*call " & LCase$(nombre) & "(*" Then
            gLlamados(nombre) = True
            GoTo siguiente
        End If

        ' 2) Nombre(...)
        If texto Like "*" & nombre & "(*" Then
            gLlamados(nombre) = True
            GoTo siguiente
        End If

        ' 3) Nombre arg1, arg2
        If texto Like "*" & nombre & " *" Then
            gLlamados(nombre) = True
            GoTo siguiente
        End If

        ' 4) If Nombre(...) Then
        If texto Like "*If *" & nombre & "(*" Then
            gLlamados(nombre) = True
            GoTo siguiente
        End If

        ' 5) If Not Nombre(...) Then
        If texto Like "*If Not *" & nombre & "(*" Then
            gLlamados(nombre) = True
            GoTo siguiente
        End If

        ' 6) x = Nombre(...)
        If texto Like "*=* " & nombre & "(*" Then
            gLlamados(nombre) = True
            GoTo siguiente
        End If

        ' 7) Set x = Nombre(...)
        If texto Like "*Set * = " & nombre & "(*" Then
            gLlamados(nombre) = True
            GoTo siguiente
        End If

        ' 8) ElseIf Nombre(...) Then
        If texto Like "*ElseIf *" & nombre & "(*" Then
            gLlamados(nombre) = True
            GoTo siguiente
        End If

        ' 9) Debug.Print Nombre(...)
        If texto Like "*Debug.Print *" & nombre & "(*" Then
            gLlamados(nombre) = True
            GoTo siguiente
        End If

        ' 10) For i = Nombre(...)
        If texto Like "*For * = " & nombre & "(*" Then
            gLlamados(nombre) = True
            GoTo siguiente
        End If

        ' Nuevos
        ' 11) x = Nombre
        If texto Like "*=* " & nombre Then
            gLlamados(nombre) = True
            GoTo siguiente
        End If
        
        ' 12) If Nombre Then
        If texto Like "*If *" & nombre & " Then*" Then
            gLlamados(nombre) = True
            GoTo siguiente
        End If
        
        ' 13) Set x = Nombre
        If texto Like "*Set * = " & nombre Then
            gLlamados(nombre) = True
            GoTo siguiente
        End If

        ' Nuevo 2
        ' 14) Llamadas dentro de expresiones concatenadas
        If texto Like "*& *" & nombre & "(*" Then
            gLlamados(nombre) = True
            GoTo siguiente
        End If
        
        If texto Like "*& *" & nombre & " *" Then
            gLlamados(nombre) = True
            GoTo siguiente
        End If
        
        ' 15) Llamadas dentro de expresiones compuestas
        If texto Like "* " & nombre & "(*" Then
            gLlamados(nombre) = True
            GoTo siguiente
        End If
        
        If texto Like "* " & nombre & " *" Then
            gLlamados(nombre) = True
            GoTo siguiente
        End If


siguiente:
    Next nombre

End Sub


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
'        ' Evitar métodos de objetos
'        If texto Like "*." & nombre & "*" Then GoTo siguiente
'
'        ' Evitar comparaciones
'        If texto Like "* " & nombre & "=" Then GoTo siguiente
'        If texto Like "*=" & nombre & "*" Then GoTo siguiente
'
'        ' 1) Call nombre(...)
'        If LCase$(texto) Like "*call " & LCase$(nombre) & "(*" Then
'            gLlamados(nombre) = True
'            GoTo siguiente
'        End If
'
'        ' 2) nombre(...)
'        If texto Like nombre & "(*" Then
'            gLlamados(nombre) = True
'            GoTo siguiente
'        End If
'
'        ' 3) nombre arg1, arg2
'        If texto Like nombre & " *" Then
'            gLlamados(nombre) = True
'            GoTo siguiente
'        End If
'
'siguiente:
'    Next nombre
'
'End Sub

' ============================================================
' REGISTRAR LLAMADAS EN MACROS (BLINDADO)
' ============================================================

Private Sub RegistrarLlamadasEnMacros()

    Dim m As AccessObject
    For Each m In Application.CurrentProject.AllMacros

        Dim ruta As String
        ruta = Environ$("TEMP") & "\macro_tmp.txt"

        Application.SaveAsText acMacro, m.Name, ruta

        Dim f As Integer
        f = FreeFile
        Open ruta For Input As #f

        Dim linea As String
        Do While Not EOF(f)
            Line Input #f, linea

            Dim proc As Variant
            For Each proc In gProcedimientos.Keys

                ' Solo detectar llamadas reales
                If linea Like "*Function Name=""" & proc & """*" Then
                    gLlamados(proc) = True
                End If

            Next proc

        Loop

        Close #f
        Kill ruta

    Next m

End Sub

' ============================================================
' REGISTRAR LLAMADAS EN FORMULARIOS (BLINDADO)
' ============================================================
Private Sub RegistrarLlamadasEnFormularios()

    Dim frm As AccessObject
    For Each frm In Application.CurrentProject.AllForms

        Dim ruta As String
        ruta = Environ$("TEMP") & "\form_tmp.txt"

        Application.SaveAsText acForm, frm.Name, ruta

        Dim f As Integer
        f = FreeFile
        Open ruta For Input As #f

        Dim linea As String
        Do While Not EOF(f)
            Line Input #f, linea

            ' Ignorar eventos automáticos
            If InStr(linea, "[Event Procedure]") > 0 Then GoTo siguienteLinea

            Dim proc As Variant
            For Each proc In gProcedimientos.Keys

                ' --- 1) RunCode en macros incrustadas ---
                ' <Action Name="RunCode">
                '   <Argument Name="FunctionName">MiFuncion</Argument>
                ' </Action>
                If linea Like "*FunctionName*" & proc & "*" Then
                    gLlamados(proc) = True
                End If

                ' --- 2) SetValue con expresiones ---
                ' <Argument Name="Item">=MiFuncion([Campo])</Argument>
                If linea Like "*=*" & proc & "(*" Then
                    gLlamados(proc) = True
                End If

                ' --- 3) WhereCondition con funciones ---
                ' <Argument Name="WhereCondition">=MiFuncion([ID])</Argument>
                If linea Like "*WhereCondition*" And linea Like "*=*" & proc & "(*" Then
                    gLlamados(proc) = True
                End If

                ' --- 4) Condition en macros incrustadas ---
                ' <Condition>=MiFuncion([Campo])</Condition>
                If linea Like "*<Condition>*" & proc & "(*" Then
                    gLlamados(proc) = True
                End If

                ' --- 5) Expresiones de evento ---
                ' OnClick = =MiFuncion([Campo])
                If linea Like "*=" & proc & "(*" Then
                    gLlamados(proc) = True
                End If

            Next proc

siguienteLinea:
        Loop

        Close #f
        Kill ruta

    Next frm

End Sub

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
'            ' Ignorar eventos automáticos
'            If InStr(linea, "[Event Procedure]") > 0 Then GoTo siguienteLinea
'
'            Dim proc As Variant
'            For Each proc In gProcedimientos.Keys
'
'                ' Llamadas tipo =MiFuncion(...)
'                If linea Like "*=" & proc & "(*" Then
'                    gLlamados(proc) = True
'                End If
'
'                ' Llamadas tipo RunCode
'                If linea Like "*RunCode*" And linea Like "*" & proc & "*" Then
'                    gLlamados(proc) = True
'                End If
'
'            Next proc
'
'siguienteLinea:
'        Loop
'
'        Close #f
'        Kill ruta
'
'    Next frm
'
'End Sub

' ============================================================
' EXCLUSIONES
' ============================================================

Private Function EsEvento(nombre As String) As Boolean
    EsEvento = (InStr(1, nombre, "_") > 0)
End Function

Private Function EsAPI(nombre As String) As Boolean

    Dim comp As VBIDE.VBComponent
    Dim cm As VBIDE.CodeModule
    Dim i As Long

    For Each comp In Application.VBE.ActiveVBProject.VBComponents
        Set cm = comp.CodeModule

        For i = 1 To cm.CountOfLines
            Dim linea As String
            linea = Trim$(cm.Lines(i, 1))

            If LCase$(linea) Like "declare *" Then
                If InStr(1, linea, nombre, vbTextCompare) > 0 Then
                    EsAPI = True
                    Exit Function
                End If
            End If
        Next i
    Next comp

End Function

