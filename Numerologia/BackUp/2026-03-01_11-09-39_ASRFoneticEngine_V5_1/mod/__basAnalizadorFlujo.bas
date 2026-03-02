Attribute VB_Name = "__basAnalizadorFlujo"

Option Compare Database
Option Explicit

'' ===========================
''  ANALIZADOR DE FLUJO VBA
''  (Access 2013/2019)
'' ===========================
'
'' Referencias necesarias:
'' - Microsoft Visual Basic for Applications Extensibility 5.3
'' - Activar: "Confiar en el acceso al modelo de objetos de proyecto VBA"
'
'' Variables globales internas del módulo
'Private gTodosProcedimientos As Object   ' Diccionario: nombreProc -> True
'Private gVisitados As Object             ' Para evitar bucles en el árbol
'
'
'' ===========================
''  PÚBLICO: ÁRBOL DE FLUJO
'' ===========================
'
'Public Sub GenerarArbolFlujo(ByVal ProcedimientoInicial As String)
'    Set gTodosProcedimientos = ObtenerTodosLosProcedimientos()
'    Set gVisitados = CreateObject("Scripting.Dictionary")
'
'    Dim Archivo As Integer
'    Archivo = FreeFile
'    Open CurrentProject.Path & "\arbol_flujo.txt" For Output As #Archivo
'
'    Print #Archivo, "Árbol de flujo desde: " & ProcedimientoInicial
'    Print #Archivo, String(60, "-")
'
'    DibujarNodo ProcedimientoInicial, "", True, Archivo
'
'    Close #Archivo
'    MsgBox "Árbol generado: arbol_flujo.txt"
'End Sub


'' ===========================
''  PÚBLICO: HUÉRFANOS
'' ===========================
'
'Public Sub DetectarProcedimientosHuerfanos()
'    Dim Todos As Object
'    Set Todos = ObtenerTodosLosProcedimientos()
'
'    Dim Llamados As Object
'    Set Llamados = CreateObject("Scripting.Dictionary")
'
'    Dim comp As VBIDE.VBComponent
'    Dim cm As VBIDE.CodeModule
'    Dim i As Long, nombre As String
'
'    ' Recorrer todo el proyecto y registrar llamadas
'    For Each comp In Application.VBE.ActiveVBProject.VBComponents
'        Set cm = comp.CodeModule
'        i = 1
'
'        Do While i < cm.CountOfLines
'            nombre = cm.ProcOfLine(i, vbext_pk_Proc)
'
'            If Len(nombre) > 0 Then
'                Dim llamadas As Collection
'                Set llamadas = ObtenerLlamadasDeProcedimiento(cm, nombre, Todos)
'
'                Dim l As Variant
'                For Each l In llamadas
'                    If Not Llamados.Exists(l) Then Llamados.Add l, True
'                Next l
'
'                i = i + cm.ProcCountLines(nombre, vbext_pk_Proc)
'            Else
'                i = i + 1
'            End If
'        Loop
'    Next comp
'
'    ' Archivo de salida
'    Dim Archivo As Integer
'    Archivo = FreeFile
'    Open CurrentProject.Path & "\procedimientos_huerfanos.txt" For Output As #Archivo
'
'    Print #Archivo, "Procedimientos huérfanos (no llamados por ningún otro procedimiento)"
'    Print #Archivo, String(60, "-")
'
'    Dim Proc As Variant
'    For Each Proc In Todos.Keys
'        If Not Llamados.Exists(Proc) Then
'            Print #Archivo, Proc
'        End If
'    Next Proc
'
'    Close #Archivo
'    MsgBox "Análisis completado: procedimientos_huerfanos.txt"
'End Sub
'
'
'' ===========================
''  ÁRBOL ASCII
'' ===========================
'
'Private Sub DibujarNodo(ByVal NombreProc As String, _
'                        ByVal Prefijo As String, _
'                        ByVal EsUltimo As Boolean, _
'                        ByVal Archivo As Integer)
'
'    Dim ModuloNombre As String
'    Dim cm As VBIDE.CodeModule
'    Set cm = BuscarModuloDeProcedimiento(NombreProc, ModuloNombre)
'
'    Dim rama As String
'    If EsUltimo Then
'        rama = Prefijo & "+-- "
'    Else
'        rama = Prefijo & "+-- "
'    End If
'
'    If cm Is Nothing Then
'        Print #Archivo, rama & NombreProc & "  (NO ENCONTRADO)"
'        Exit Sub
'    End If
'
'    If gVisitados Is Nothing Then Set gVisitados = CreateObject("Scripting.Dictionary")
'
'    If gVisitados.Exists(NombreProc) Then
'        Print #Archivo, rama & NombreProc & "   [Módulo: " & ModuloNombre & "] (ya visitado)"
'        Exit Sub
'    End If
'
'    gVisitados.Add NombreProc, True
'    Print #Archivo, rama & NombreProc & "   [Módulo: " & ModuloNombre & "]"
'
'    Dim llamadas As Collection
'    Set llamadas = ObtenerLlamadasDeProcedimiento(cm, NombreProc, gTodosProcedimientos)
'
'    If llamadas.count = 0 Then Exit Sub
'
'    Dim i As Long
'    For i = 1 To llamadas.count
'        Dim nuevoPrefijo As String
'        If EsUltimo Then
'            nuevoPrefijo = Prefijo & "    "
'        Else
'            nuevoPrefijo = Prefijo & "¦   "
'        End If
'
'        DibujarNodo llamadas(i), nuevoPrefijo, (i = llamadas.count), Archivo
'    Next i
'End Sub
'
'
'' ===========================
''  OBTENER TODOS LOS PROCEDIMIENTOS
'' ===========================
'
'Private Function ObtenerTodosLosProcedimientos() As Object
'    Dim dict As Object
'    Set dict = CreateObject("Scripting.Dictionary")
'
'    Dim comp As VBIDE.VBComponent
'    Dim cm As VBIDE.CodeModule
'    Dim i As Long, nombre As String
'
'    For Each comp In Application.VBE.ActiveVBProject.VBComponents
'        Set cm = comp.CodeModule
'        i = 1
'
'        Do While i < cm.CountOfLines
'            nombre = cm.ProcOfLine(i, vbext_pk_Proc)
'            If Len(nombre) > 0 Then
'                If Not dict.Exists(nombre) Then dict.Add nombre, True
'                i = i + cm.ProcCountLines(nombre, vbext_pk_Proc)
'            Else
'                i = i + 1
'            End If
'        Loop
'    Next comp
'
'    Set ObtenerTodosLosProcedimientos = dict
'End Function
'
'
'' ===========================
''  BUSCAR MÓDULO DE UN PROCEDIMIENTO
'' ===========================
'
'Private Function BuscarModuloDeProcedimiento(ByVal NombreProc As String, _
'                                             ByRef ModuloNombre As String) _
'                                             As VBIDE.CodeModule
'    Dim comp As VBIDE.VBComponent
'    Dim cm As VBIDE.CodeModule
'
'    ' Si tenemos diccionario global, filtramos primero
'    If Not gTodosProcedimientos Is Nothing Then
'        If Not gTodosProcedimientos.Exists(NombreProc) Then
'            Exit Function
'        End If
'    End If
'
'    For Each comp In Application.VBE.ActiveVBProject.VBComponents
'        Set cm = comp.CodeModule
'
'        On Error Resume Next
'        Dim linea As Long
'        linea = cm.ProcStartLine(NombreProc, vbext_pk_Proc)
'        On Error GoTo 0
'
'        If linea > 0 Then
'            ModuloNombre = comp.Name
'            Set BuscarModuloDeProcedimiento = cm
'            Exit Function
'        End If
'    Next comp
'End Function
'
'
'' ===========================
''  RANGO DE UN PROCEDIMIENTO
'' ===========================
'
'Private Function ObtenerRangoProcedimiento(cm As VBIDE.CodeModule, _
'                                           ByVal NombreProc As String, _
'                                           ByRef inicio As Long, _
'                                           ByRef fin As Long) As Boolean
'    On Error GoTo ErrHandler
'
'    inicio = cm.ProcStartLine(NombreProc, vbext_pk_Proc)
'    fin = inicio + cm.ProcCountLines(NombreProc, vbext_pk_Proc) - 1
'    ObtenerRangoProcedimiento = True
'    Exit Function
'
'ErrHandler:
'    ObtenerRangoProcedimiento = False
'End Function
'
'
'' ===========================
''  OBTENER LLAMADAS DE UN PROCEDIMIENTO
'' ===========================
'
'Private Function ObtenerLlamadasDeProcedimiento(cm As VBIDE.CodeModule, _
'                                                ByVal NombreProc As String, _
'                                                ByVal Todos As Object) As Collection
'    Dim inicio As Long, fin As Long
'    Dim linea As Long
'    Dim texto As String
'
'    Set ObtenerLlamadasDeProcedimiento = New Collection
'
'    If Not ObtenerRangoProcedimiento(cm, NombreProc, inicio, fin) Then Exit Function
'
'    For linea = inicio To fin
'        texto = cm.Lines(linea, 1)
'
'        Dim llamadas As New Collection
'        DetectarLlamadasAvanzado texto, llamadas, Todos, NombreProc  ' <<-- le pasamos el nombre actual
'
'        Dim l As Variant
'        For Each l In llamadas
'            On Error Resume Next
'            ObtenerLlamadasDeProcedimiento.Add l
'            On Error GoTo 0
'        Next l
'    Next linea
'End Function
'
''Private Function ObtenerLlamadasDeProcedimiento(cm As VBIDE.CodeModule, _
''                                                ByVal NombreProc As String, _
''                                                ByVal Todos As Object) As Collection
''    Dim Inicio As Long, Fin As Long
''    Dim linea As Long
''    Dim texto As String
''
''    Set ObtenerLlamadasDeProcedimiento = New Collection
''
''    If Not ObtenerRangoProcedimiento(cm, NombreProc, Inicio, Fin) Then Exit Function
''
''    For linea = Inicio To Fin
''        texto = cm.Lines(linea, 1)
''
''        Dim llamadas As New Collection
''        DetectarLlamadasAvanzado texto, llamadas, Todos
''
''        Dim l As Variant
''        For Each l In llamadas
''            On Error Resume Next
''            ObtenerLlamadasDeProcedimiento.Add l
''            On Error GoTo 0
''        Next l
''    Next linea
''End Function
'
'
'' ===========================
''  DETECTOR AVANZADO DE LLAMADAS
'' ===========================
'Private Sub DetectarLlamadasAvanzado(ByVal linea As String, _
'                                     ByRef Lista As Collection, _
'                                     ByVal Todos As Object, _
'                                     ByVal ProcActual As String)
'
'    Dim texto As String
'    texto = Trim$(linea)
'
'    ' Quitar comentarios
'    If InStr(texto, "'") > 0 Then
'        texto = Left$(texto, InStr(texto, "'") - 1)
'    End If
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
'    For Each nombre In Todos.Keys
'
'        If nombre = ProcActual Then GoTo siguiente
'
'        ' 1) Detectar "Call nombre"
'        If LCase$(texto) Like "*call " & LCase$(nombre) & "*" Then
'            On Error Resume Next
'            Lista.Add nombre
'            On Error GoTo 0
'            GoTo siguiente
'        End If
'
'        ' 2) Detectar "nombre("  ? llamada con paréntesis
'        If InStr(1, texto, nombre & "(", vbTextCompare) > 0 Then
'            On Error Resume Next
'            Lista.Add nombre
'            On Error GoTo 0
'            GoTo siguiente
'        End If
'
'        ' 3) Detectar "nombre " seguido de algo que no sea "="
'        Dim pos As Long
'        pos = InStr(1, texto, nombre & " ", vbTextCompare)
'        If pos > 0 Then
'            Dim resto As String
'            resto = Trim$(Mid$(texto, pos + Len(nombre)))
'            If Left$(resto, 1) <> "=" Then
'                On Error Resume Next
'                Lista.Add nombre
'                On Error GoTo 0
'                GoTo siguiente
'            End If
'        End If
'
'siguiente:
'    Next nombre
'End Sub
'
''Private Sub DetectarLlamadasAvanzado(ByVal linea As String, _
''                                     ByRef Lista As Collection, _
''                                     ByVal Todos As Object, _
''                                     ByVal ProcActual As String)
''
''    Dim texto As String
''    texto = Trim$(linea)
''
''    ' Quitar comentarios
''    If InStr(texto, "'") > 0 Then
''        texto = Left$(texto, InStr(texto, "'") - 1)
''    End If
''    If Len(texto) = 0 Then Exit Sub
''
''    ' Quitar cadenas
''    Do While InStr(texto, """") > 0
''        Dim p1 As Long, p2 As Long
''        p1 = InStr(texto, """")
''        p2 = InStr(p1 + 1, texto, """")
''        If p2 = 0 Then Exit Do
''        texto = Left$(texto, p1 - 1) & Mid$(texto, p2 + 1)
''    Loop
''
''    Dim nombre As Variant
''    For Each nombre In Todos.Keys
''
''        If nombre = ProcActual Then GoTo siguiente
''
''        Dim pos As Long
''        pos = 1
''
''        Do
''            pos = InStr(pos, texto, nombre, vbTextCompare)
''            If pos = 0 Then Exit Do
''
''            Dim antes As String, despues As String
''
''            'antes = IIf(pos > 1, Mid$(texto, pos - 1, 1), " ")
''            If pos > 1 Then
''                antes = Mid$(texto, pos - 1, 1)
''            Else
''                antes = " "
''            End If
''
''            'despues = IIf(pos + Len(nombre) <= Len(texto), Mid$(texto, pos + Len(nombre), 1), " ")
''            If pos + Len(nombre) <= Len(texto) Then
''                despues = Mid$(texto, pos + Len(nombre), 1)
''            Else
''                despues = " "
''            End If
''
''
''            ' Debe estar delimitado por no-letras
''            If Not (antes Like "[A-Za-z0-9_]" Or despues Like "[A-Za-z0-9_]") Then
''
''                Dim resto As String
''                resto = LCase$(Mid$(texto, pos + Len(nombre)))
''                resto = Trim$(resto)
''
''                ' No es declaración
''                If LCase$(Left$(texto, pos - 1)) Like "*sub" _
''                Or LCase$(Left$(texto, pos - 1)) Like "*function" _
''                Or LCase$(Left$(texto, pos - 1)) Like "*dim" _
''                Or LCase$(Left$(texto, pos - 1)) Like "*public" _
''                Or LCase$(Left$(texto, pos - 1)) Like "*private" Then
''                    GoTo continuar
''                End If
''
''                ' No es asignación
''                If Left$(resto, 1) = "=" Then GoTo continuar
''
''                ' No es método de objeto
''                If antes = "." Then GoTo continuar
''
''                ' No es etiqueta
''                If despues = ":" Then GoTo continuar
''
''                ' Si llega aquí, es llamada real
''                On Error Resume Next
''                Lista.Add nombre
''                On Error GoTo 0
''
''            End If
''
''continuar:
''            pos = pos + Len(nombre)
''        Loop
''
''siguiente:
''    Next nombre
''End Sub
'
''Private Sub DetectarLlamadasAvanzado(ByVal linea As String, _
''                                     ByRef ListaLlamadas As Collection, _
''                                     ByVal Todos As Object, _
''                                     ByVal ProcActual As String)
''
''    Dim texto As String
''    texto = Trim$(linea)
''
''    ' Quitar comentarios
''    If InStr(texto, "'") > 0 Then
''        texto = Left$(texto, InStr(texto, "'") - 1)
''    End If
''    If Len(texto) = 0 Then Exit Sub
''
''    ' Quitar cadenas entre comillas
''    Do While InStr(texto, """") > 0
''        Dim p1 As Long, p2 As Long
''        p1 = InStr(texto, """")
''        p2 = InStr(p1 + 1, texto, """")
''        If p2 = 0 Then Exit Do
''        texto = Left$(texto, p1 - 1) & Mid$(texto, p2 + 1)
''    Loop
''
''    Dim nombre As Variant
''    For Each nombre In Todos.Keys
''
''        ' No queremos que se detecte a sí mismo
''        If nombre = ProcActual Then GoTo siguienteNombre
''
''        Dim pos As Long
''        pos = 1
''
''        Do
''            pos = InStr(pos, texto, nombre, vbTextCompare)
''            If pos = 0 Then Exit Do
''
''            Dim cAntes As String, cDespues As String
''
''            If pos > 1 Then
''                cAntes = Mid$(texto, pos - 1, 1)
''            Else
''                cAntes = " "
''            End If
''
''            If pos + Len(nombre) <= Len(texto) Then
''                cDespues = Mid$(texto, pos + Len(nombre), 1)
''            Else
''                cDespues = " "
''            End If
''
''            ' Debe estar delimitado por no-letras
''            If Not (cAntes Like "[A-Za-z0-9_]" Or cDespues Like "[A-Za-z0-9_]") Then
''
''                ' Ignorar declaraciones: Sub/Function/Dim/Public/Private nombre
''                Dim cabecera As String
''                'cabecera = LCase$(Left$(Application.Trim$(texto), pos - 1))
''                cabecera = LCase$(Left$(Trim$(texto), pos - 1))
''                If cabecera Like "*sub" Or cabecera Like "*function" _
''                   Or cabecera Like "*dim" Or cabecera Like "*public" _
''                   Or cabecera Like "*private" Then
''                    ' no es llamada
''                Else
''                    ' Ignorar asignaciones: nombre = ...
''                    Dim resto As String
''                    resto = Mid$(texto, pos + Len(nombre))
''                    'resto = LCase$(Application.Trim$(resto))
''                    resto = LCase$(Trim$(resto))
''                    If Left$(resto, 1) <> "=" Then
''                        ' Aquí sí lo consideramos llamada
''                        On Error Resume Next
''                        ListaLlamadas.Add CStr(nombre)
''                        On Error GoTo 0
''                    End If
''                End If
''            End If
''
''            pos = pos + Len(nombre)
''        Loop
''
''siguienteNombre:
''    Next nombre
''End Sub
'
''Private Sub DetectarLlamadasAvanzado(ByVal linea As String, _
''                                     ByRef ListaLlamadas As Collection, _
''                                     ByVal Todos As Object)
''
''    Dim texto As String
''    texto = Trim(linea)
''
''    ' Quitar comentarios
''    If InStr(texto, "'") > 0 Then
''        texto = Left(texto, InStr(texto, "'") - 1)
''    End If
''    If Len(texto) = 0 Then Exit Sub
''
''    ' Quitar cadenas entre comillas
''    Do While InStr(texto, """") > 0
''        Dim p1 As Long, p2 As Long
''        p1 = InStr(texto, """")
''        p2 = InStr(p1 + 1, texto, """")
''        If p2 = 0 Then Exit Do
''        texto = Left(texto, p1 - 1) & Mid(texto, p2 + 1)
''    Loop
''
''    ' Normalizar
''    texto = Replace(texto, "(", " ( ")
''    texto = Replace(texto, ")", " ) ")
''    texto = Replace(texto, ",", " , ")
''    texto = Trim(texto) 'Application.Trim(texto)
''
''    Dim palabras() As String
''    palabras = Split(texto, " ")
''
''    Dim i As Long
''    For i = 0 To UBound(palabras)
''        Dim token As String
''        token = Trim(palabras(i))
''
''        If Len(token) = 0 Then GoTo siguiente
''        If EsPalabraReservada(token) Then GoTo siguiente
''
''        ' Detectar "Call X"
''        If LCase(token) = "call" Then
''            If i < UBound(palabras) Then
''                AgregarSiEsProcedimiento palabras(i + 1), ListaLlamadas, Todos
''            End If
''            GoTo siguiente
''        End If
''
''        ' Solo considerar tokens EXACTOS
''        If Todos.Exists(token) Then
''
''            ' Ignorar asignaciones: X = Proc
''            If i > 0 Then
''                If palabras(i - 1) = "=" Then GoTo siguiente
''            End If
''
''            ' Ignorar comparaciones: If Proc = ...
''            If i < UBound(palabras) Then
''                If palabras(i + 1) = "=" Then GoTo siguiente
''            End If
''
''            ' Ignorar variables con el mismo nombre
''            If i > 0 Then
''                If LCase(palabras(i - 1)) = "dim" Or _
''                   LCase(palabras(i - 1)) = "private" Or _
''                   LCase(palabras(i - 1)) = "public" Then GoTo siguiente
''            End If
''
''            ' Ignorar propiedades de objetos: obj.Proc
''            If InStr(token, ".") > 0 Then GoTo siguiente
''
''            ' Si llega aquí, es una llamada real
''            AgregarSiEsProcedimiento token, ListaLlamadas, Todos
''        End If
''
''siguiente:
''    Next i
''End Sub
'
''Private Sub DetectarLlamadasAvanzado(ByVal linea As String, _
''                                     ByRef ListaLlamadas As Collection, _
''                                     ByVal TodosLosProcedimientos As Object)
''
''    Dim texto As String
''    texto = Trim(linea)
''
''    ' Quitar comentarios
''    If InStr(texto, "'") > 0 Then
''        texto = Left(texto, InStr(texto, "'") - 1)
''    End If
''    If Len(texto) = 0 Then Exit Sub
''
''    ' Quitar cadenas entre comillas
''    Do While InStr(texto, """") > 0
''        Dim p1 As Long, p2 As Long
''        p1 = InStr(texto, """")
''        p2 = InStr(p1 + 1, texto, """")
''        If p2 = 0 Then Exit Do
''        texto = Left(texto, p1 - 1) & Mid(texto, p2 + 1)
''    Loop
''
''    ' Normalizar espacios
''    texto = Replace(texto, "(", " ( ")
''    texto = Replace(texto, ")", " ) ")
''    texto = Replace(texto, ",", " , ")
''    texto = Trim(texto) 'Application.Trim(texto)
''
''    Dim palabras() As String
''    palabras = Split(texto, " ")
''
''    Dim i As Long
''    For i = 0 To UBound(palabras)
''        Dim token As String
''        token = Trim(palabras(i))
''        If Len(token) = 0 Then GoTo siguiente
''
''        ' Palabras reservadas
''        If EsPalabraReservada(token) Then GoTo siguiente
''
''        ' "Call X"
''        If LCase(token) = "call" Then
''            If i < UBound(palabras) Then
''                AgregarSiEsProcedimiento palabras(i + 1), ListaLlamadas, TodosLosProcedimientos
''            End If
''            GoTo siguiente
''        End If
''
''        ' Si el token es un nombre de procedimiento conocido
''        If TodosLosProcedimientos.Exists(token) Then
''            ' Evitar asignaciones: x = Funcion(...)
''            If i > 0 Then
''                If palabras(i - 1) = "=" Then GoTo siguiente
''            End If
''
''            ' Evitar propiedades de objetos: obj.Metodo
''            If InStr(token, ".") > 0 Then
''                Dim parte As String
''                parte = Split(token, ".")(UBound(Split(token, ".")))
''                If TodosLosProcedimientos.Exists(parte) Then
''                    AgregarSiEsProcedimiento parte, ListaLlamadas, TodosLosProcedimientos
''                End If
''            Else
''                AgregarSiEsProcedimiento token, ListaLlamadas, TodosLosProcedimientos
''            End If
''        End If
''
''siguiente:
''    Next i
''End Sub
'
'
'Private Function EsPalabraReservada(ByVal p As String) As Boolean
'    Dim r As Variant
'    Dim reservadas As Variant
'
'    reservadas = Array("if", "then", "else", "elseif", "end", "for", "next", _
'                       "while", "wend", "do", "loop", "select", "case", _
'                       "set", "let", "with", "byval", "byref", "as", _
'                       "dim", "private", "public", "function", "sub", _
'                       "property", "get", "call", "exit", "goto", _
'                       "option", "on", "error", "resume")
'
'    For Each r In reservadas
'        If LCase(p) = r Then
'            EsPalabraReservada = True
'            Exit Function
'        End If
'    Next r
'End Function
'
'
'Private Sub AgregarSiEsProcedimiento(ByVal nombre As String, _
'                                     ByRef Lista As Collection, _
'                                     ByVal Todos As Object)
'
'    nombre = Replace(nombre, "(", "")
'    nombre = Replace(nombre, ")", "")
'    nombre = Trim(nombre)
'
'    If Len(nombre) = 0 Then Exit Sub
'    If Not Todos.Exists(nombre) Then Exit Sub
'
'    On Error Resume Next
'    Lista.Add nombre
'    On Error GoTo 0
'End Sub


