Attribute VB_Name = "___basInfoModulos"


Option Compare Database
Option Explicit

Public Sub ExportarInfoModulos()
    Dim vbProj As VBIDE.VBProject
    Dim vbComp As VBIDE.VBComponent
    Dim vbMod As VBIDE.CodeModule
    
    Dim ruta As String
    Dim f As Integer
    Dim linea As Long
    Dim NombreProc As String
    Dim tipoProc As VBIDE.vbext_ProcKind
    Dim visib As String
    
    Set vbProj = Application.VBE.ActiveVBProject
    
    ruta = CurrentProject.Path & "\ListadoModulos_" & vbProj.Name & ".txt"
    f = FreeFile
    
    Open ruta For Output As #f
    
    ' --- Nombre del proyecto ---
    Print #f, ""
    Print #f, String(40, "=")
    Print #f, "PROYECTO: " & vbProj.Name
    Print #f, String(40, "=")
    Print #f, ""
    
    For Each vbComp In vbProj.VBComponents
        Set vbMod = vbComp.CodeModule
        
        Print #f, "Módulo: " & vbComp.Name
        Print #f, "Tipo:   " & TipoModulo(vbComp.Type)
        Print #f, "Procedimientos:"
        
        linea = 1
        Do While linea < vbMod.CountOfLines
            NombreProc = vbMod.ProcOfLine(linea, tipoProc)
            If NombreProc <> "" Then
                
                ' Obtener visibilidad del procedimiento
                visib = VisibilidadProcedimiento(vbMod, NombreProc, tipoProc)
                
                'Print #f, "   - " & nombreProc & "   (" & visib & ")"
                Print #f, "    (" & visib & ")" & vbTab & "- " & NombreProc; ""
                
                linea = vbMod.ProcStartLine(NombreProc, tipoProc) + _
                        vbMod.ProcCountLines(NombreProc, tipoProc)
            Else
                linea = linea + 1
            End If
        Loop
        
        Print #f, ""
    Next vbComp
    
    Close #f
    
    MsgBox "Fichero generado en: " & ruta, vbInformation
End Sub

Public Sub ExportarInfoMotores()
    Dim vbProj As VBIDE.VBProject
    Dim vbComp As VBIDE.VBComponent
    Dim vbMod As VBIDE.CodeModule
    
    Dim ruta As String
    Dim f As Integer
    Dim linea As Long
    Dim NombreProc As String
    Dim tipoProc As VBIDE.vbext_ProcKind
    Dim visib As String
    
    Set vbProj = Application.VBE.ActiveVBProject
    
    ruta = CurrentProject.Path & "\ListadoMotores_" & vbProj.Name & ".txt"
    f = FreeFile
    
    Open ruta For Output As #f
    
    ' --- Nombre del proyecto ---
    Print #f, ""
    Print #f, String(40, "=")
    Print #f, "PROYECTO: " & vbProj.Name
    Print #f, String(40, "=")
    Print #f, ""
    
    Print #f, String(40, "-")
    Print #f, "Lista módulos motores"
    Print #f, String(40, "-")
    Print #f, ""
    
    For Each vbComp In vbProj.VBComponents
        If Left(vbComp.Name, 10) = "bas_Motor_" Then
            Set vbMod = vbComp.CodeModule
            
            Print #f, "Módulo: " & vbComp.Name
            Print #f, "Tipo:   " & TipoModulo(vbComp.Type)
            Print #f, "Procedimientos:"
            
            linea = 1
            Do While linea < vbMod.CountOfLines
                NombreProc = vbMod.ProcOfLine(linea, tipoProc)
                If NombreProc <> "" Then
                    
                    ' Obtener visibilidad del procedimiento
                    visib = VisibilidadProcedimiento(vbMod, NombreProc, tipoProc)
                    
                    'Print #f, "   - " & nombreProc & "   (" & visib & ")"
                    Print #f, "    (" & visib & ")" & vbTab & "- " & NombreProc; ""
                    
                    linea = vbMod.ProcStartLine(NombreProc, tipoProc) + _
                            vbMod.ProcCountLines(NombreProc, tipoProc)
                Else
                    linea = linea + 1
                End If
            Loop
            
            Print #f, ""
            End If
    Next vbComp
    
    Close #f
    
    MsgBox "Fichero generado en: " & ruta, vbInformation
End Sub

Private Function TipoModulo(tipo As VBIDE.vbext_ComponentType) As String
    Select Case tipo
        Case vbext_ct_StdModule: TipoModulo = "Módulo estándar"
        Case vbext_ct_ClassModule: TipoModulo = "Módulo de clase"
        Case vbext_ct_MSForm: TipoModulo = "Formulario"
        Case vbext_ct_Document: TipoModulo = "Módulo de formulario/informe"
        Case Else: TipoModulo = "Desconocido"
    End Select
End Function

Private Function VisibilidadProcedimiento(cm As VBIDE.CodeModule, _
                                          nombre As String, _
                                          tipo As VBIDE.vbext_ProcKind) As String
    Dim lineaInicio As Long
    Dim texto As String
    
    lineaInicio = cm.ProcStartLine(nombre, tipo)
    
    texto = Trim$(cm.Lines(lineaInicio, 1))
    While InStr(texto, nombre) = 0
        lineaInicio = lineaInicio + 1
        texto = Trim$(cm.Lines(lineaInicio, 1))
        If Left(texto, 1) = "'" Then texto = "'"
    Wend
    
    'texto = Trim$(cm.Lines(lineaInicio + 1, 1))
    
    If LCase$(Left$(texto, 6)) = "public" Then
        VisibilidadProcedimiento = "Public"
    ElseIf LCase$(Left$(texto, 7)) = "private" Then
        VisibilidadProcedimiento = "Private"
    Else
        ' Si no especifica, en módulos estándar es público por defecto
        VisibilidadProcedimiento = "Public (implícito)"
    End If
End Function
'===========================================================================================


''Option Compare Database
''Option Explicit
'
'' Diccionario para evitar bucles infinitos
'Dim Visitados As Object
'Dim Archivo As Integer
'
'Public Sub AnalizarFlujo(ProcedimientoInicial As String)
'    Set Visitados = CreateObject("Scripting.Dictionary")
'    Archivo = FreeFile
'
'    Open CurrentProject.Path & "\flujo.txt" For Output As #Archivo
'
'    Print #Archivo, "Flujo de ejecución desde: " & ProcedimientoInicial
'    Print #Archivo, String(60, "-")
'
'    Call SeguirProcedimiento(ProcedimientoInicial, 0)
'
'    Close #Archivo
'    MsgBox "Análisis completado. Archivo generado en la carpeta del proyecto."
'End Sub
'
'Private Sub SeguirProcedimiento(NombreProc As String, Nivel As Integer)
'    Dim Modulo As Object 'CodeModule
'    Dim linea As Long, Ultima As Long
'    Dim texto As String
'    Dim ModuloNombre As String
'
'    If Visitados.Exists(NombreProc) Then Exit Sub
'    Visitados.Add NombreProc, True
'
'    ' Buscar el módulo donde está el procedimiento
'    Set Modulo = BuscarModuloDeProcedimiento(NombreProc, ModuloNombre)
'
'    If Modulo Is Nothing Then
'        Print #Archivo, Space(Nivel * 4) & NombreProc & "  (NO ENCONTRADO)"
'        Exit Sub
'    End If
'
'    Print #Archivo, Space(Nivel * 4) & NombreProc & "   [Módulo: " & ModuloNombre & "]"
'
'    ' Obtener rango del procedimiento
'    Dim Inicio As Long, Fin As Long
'    If Not ObtenerRangoProcedimiento(Modulo, NombreProc, Inicio, Fin) Then Exit Sub
'
'    ' Recorrer líneas del procedimiento
'    For linea = Inicio To Fin
'        texto = Trim(Modulo.Lines(linea, 1))
'
'        ' Detectar llamadas a otros procedimientos
'        Dim Llamado As String
'        Llamado = DetectarLlamada(texto)
'
'        If Llamado <> "" Then
'            Call SeguirProcedimiento(Llamado, Nivel + 1)
'        End If
'    Next linea
'End Sub
'Private Function BuscarModuloDeProcedimiento(NombreProc As String, _
'                                             ByRef ModuloNombre As String) _
'                                             As VBIDE.CodeModule
'
'    Dim comp As VBIDE.VBComponent
'    Dim cm As VBIDE.CodeModule
'
'    ' Si el nombre no existe en el diccionario global, salir
'    If Not TodosLosProcedimientos.Exists(NombreProc) Then
'        Exit Function
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
''Private Function BuscarModuloDeProcedimiento(NombreProc As String, _
''                                             ByRef ModuloNombre As String) _
''                                             As VBIDE.CodeModule
''
''    Dim comp As VBIDE.VBComponent
''    Dim cm As VBIDE.CodeModule
''
''    For Each comp In Application.VBE.ActiveVBProject.VBComponents
''        Set cm = comp.CodeModule
''
''        If cm.ProcStartLine(NombreProc, vbext_pk_Proc) > 0 Then
''            ModuloNombre = comp.Name
''            Set BuscarModuloDeProcedimiento = cm
''            Exit Function
''        End If
''    Next comp
''End Function
'
''Private Function BuscarModuloDeProcedimiento(NombreProc As String, ByRef ModuloNombre As String) As CodeModule
''    Dim comp As VBIDE.VBComponent
''    Set comp = Application.VBE.ActiveVBProject.VBComponents("Modulo1")
''    Dim cm As VBIDE.CodeModule
''    'Set cm = comp.CodeModule   ' ?? Esto sí funciona
''
''    Dim i As Long
''
''    For Each cm In Application.VBE.ActiveVBProject.VBComponents
''        If cm.CodeModule.ProcStartLine(NombreProc, vbext_pk_Proc) > 0 Then
''            ModuloNombre = cm.Name
''            Set BuscarModuloDeProcedimiento = cm.CodeModule
''            Exit Function
''        End If
''    Next cm
''End Function
'
'Private Function ObtenerRangoProcedimiento(cm As CodeModule, NombreProc As String, _
'                                           ByRef Inicio As Long, ByRef Fin As Long) As Boolean
'    On Error GoTo ErrHandler
'
'    Inicio = cm.ProcStartLine(NombreProc, vbext_pk_Proc)
'    Fin = Inicio + cm.ProcCountLines(NombreProc, vbext_pk_Proc) - 1
'    ObtenerRangoProcedimiento = True
'    Exit Function
'
'ErrHandler:
'    ObtenerRangoProcedimiento = False
'End Function
'
'Private Function DetectarCllamada(linea As String) As String
'    Dim palabras() As String
'    Dim p As String
'
'    If linea Like "Call *" Then
'        DetectarCllamada = Trim(Mid(linea, 6))
'        Exit Function
'    End If
'
'    palabras = Split(linea, " ")
'
'    ' Si la primera palabra es un nombre de procedimiento
'    p = palabras(0)
'
'    ' Excluir palabras reservadas
'    Select Case UCase(p)
'        Case "IF", "FOR", "WHILE", "DO", "SELECT", "SET", "LET", "WITH", "ELSE", "END"
'            Exit Function
'    End Select
'
'    ' Si tiene paréntesis, probablemente es una llamada
'    If InStr(p, "(") > 0 Then
'        DetectarCllamada = Left(p, InStr(p, "(") - 1)
'    End If
'End Function
'
'Private Function DetectarLlamadasAvanzado(linea As String, _
'                                          ByRef ListaLlamadas As Collection, _
'                                          ByRef TodosLosProcedimientos As Object)
'
'    Dim texto As String
'    texto = Trim(linea)
'
'    ' Quitar comentarios
'    If InStr(texto, "'") > 0 Then
'        texto = Left(texto, InStr(texto, "'") - 1)
'    End If
'    If Len(texto) = 0 Then Exit Function
'
'    ' Quitar cadenas entre comillas
'    Do While InStr(texto, """") > 0
'        Dim p1 As Long, p2 As Long
'        p1 = InStr(texto, """")
'        p2 = InStr(p1 + 1, texto, """")
'        If p2 = 0 Then Exit Do
'        texto = Left(texto, p1 - 1) & Mid(texto, p2 + 1)
'    Loop
'
'    ' Normalizar espacios
'    texto = Replace(texto, "(", " ( ")
'    texto = Replace(texto, ")", " ) ")
'    texto = Replace(texto, ",", " , ")
'    texto = Application.Trim(texto)
'
'    Dim palabras() As String
'    palabras = Split(texto, " ")
'
'    Dim i As Long
'    For i = 0 To UBound(palabras)
'        Dim token As String
'        token = palabras(i)
'
'        ' Saltar tokens vacíos
'        If Len(token) = 0 Then GoTo siguiente
'
'        ' Saltar palabras reservadas
'        If EsPalabraReservada(token) Then GoTo siguiente
'
'        ' Si es "Call", el siguiente token es el procedimiento
'        If LCase(token) = "call" Then
'            If i < UBound(palabras) Then
'                AgregarSiEsProcedimiento palabras(i + 1), ListaLlamadas, TodosLosProcedimientos
'            End If
'            GoTo siguiente
'        End If
'
'        ' Si el token es un nombre de procedimiento conocido
'        If TodosLosProcedimientos.Exists(token) Then
'            ' Evitar confundir asignaciones: x = Funcion(...)
'            If i > 0 And palabras(i - 1) = "=" Then GoTo siguiente
'
'            ' Evitar confundir propiedades de objetos: obj.Metodo
'            If InStr(token, ".") > 0 Then
'                Dim parte As String
'                parte = Split(token, ".")(UBound(Split(token, ".")))
'                If TodosLosProcedimientos.Exists(parte) Then
'                    AgregarSiEsProcedimiento parte, ListaLlamadas, TodosLosProcedimientos
'                End If
'            Else
'                AgregarSiEsProcedimiento token, ListaLlamadas, TodosLosProcedimientos
'            End If
'        End If
'
'siguiente:
'    Next i
'
'End Function
'
'Private Function EsPalabraReservada(p As String) As Boolean
'    Dim r As Variant
'    Dim reservadas As Variant
'
'    reservadas = Array("if", "then", "else", "elseif", "end", "for", "next", _
'                       "while", "wend", "do", "loop", "select", "case", _
'                       "set", "let", "with", "byval", "byref", "as", _
'                       "dim", "private", "public", "function", "sub", _
'                       "property", "get", "set", "call", "exit", "goto")
'
'    For Each r In reservadas
'        If LCase(p) = r Then
'            EsPalabraReservada = True
'            Exit Function
'        End If
'    Next r
'End Function
'
'Private Sub AgregarSiEsProcedimiento(nombre As String, _
'                                     ByRef Lista As Collection, _
'                                     ByRef Todos As Object)
'
'    nombre = Replace(nombre, "(", "")
'    nombre = Replace(nombre, ")", "")
'    nombre = Trim(nombre)
'
'    If Todos.Exists(nombre) Then
'        On Error Resume Next
'        Lista.Add nombre
'        On Error GoTo 0
'    End If
'End Sub
'
'Private Function ObtenerTodosLosProcedimientos() As Object
'    Dim dict As Object
'    Set dict = CreateObject("Scripting.Dictionary")
'
'    Dim comp As VBIDE.VBComponent
'    Dim cm As CodeModule
'    Dim i As Long, nombre As String
'
'    For Each comp In Application.VBE.ActiveVBProject.VBComponents
'        Set cm = comp.CodeModule
'
'        i = 1
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
''---------------------------------------------------------------------------------
'
'Public Sub GenerarArbolFlujo(ProcedimientoInicial As String)
'    Dim Visitados As Object
'    Set Visitados = CreateObject("Scripting.Dictionary")
'
'    Dim Todos As Object
'    Set Todos = ObtenerTodosLosProcedimientos()
'
'    Dim Archivo As Integer
'    Archivo = FreeFile
'    Open CurrentProject.Path & "\arbol_flujo.txt" For Output As #Archivo
'
'    Print #Archivo, "Árbol de flujo desde: " & ProcedimientoInicial
'    Print #Archivo, String(60, "-")
'
'    Call DibujarNodo(ProcedimientoInicial, "", True, Visitados, Todos, Archivo)
'
'    Close #Archivo
'    MsgBox "Árbol generado en arbol_flujo.txt"
'End Sub
'
'Private Sub DibujarNodo(NombreProc As String, Prefijo As String, _
'                        EsUltimo As Boolean, Visitados As Object, _
'                        Todos As Object, Archivo As Integer)
'
'    Dim ModuloNombre As String
'    Dim cm As Object 'CodeModule
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
'    Print #Archivo, rama & NombreProc & "   [Módulo: " & ModuloNombre & "]"
'
'    If Visitados.Exists(NombreProc) Then Exit Sub
'    Visitados.Add NombreProc, True
'
'    ' Obtener llamadas internas
'    Dim llamadas As Collection
'    Set llamadas = ObtenerLlamadasDeProcedimiento(cm, NombreProc, Todos)
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
'        Call DibujarNodo(llamadas(i), nuevoPrefijo, (i = llamadas.count), Visitados, Todos, Archivo)
'    Next i
'End Sub
'
'Private Function ObtenerLlamadasDeProcedimiento(cm As CodeModule, NombreProc As String, _
'                                                Todos As Object) As Collection
'    Dim Inicio As Long, Fin As Long
'    Dim linea As Long
'    Dim texto As String
'
'    Set ObtenerLlamadasDeProcedimiento = New Collection
'
'    If Not ObtenerRangoProcedimiento(cm, NombreProc, Inicio, Fin) Then Exit Function
'
'    For linea = Inicio To Fin
'        texto = cm.Lines(linea, 1)
'
'        Dim llamadas As New Collection
'        Call DetectarLlamadasAvanzado(texto, llamadas, Todos)
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
''------------------------------------------------------------------------------------------
'
'Public Sub DetectarProcedimientosHuerfanos()
'    Dim Todos As Object
'    Set Todos = ObtenerTodosLosProcedimientos()
'
'    Dim Llamados As Object
'    Set Llamados = CreateObject("Scripting.Dictionary")
'
'    Dim comp As VBIDE.VBComponent
'    Dim cm As CodeModule
'    Dim proc As Variant
'
'    ' Analizar llamadas en todo el proyecto
'    For Each comp In Application.VBE.ActiveVBProject.VBComponents
'        Set cm = comp.CodeModule
'
'        Dim i As Long, nombre As String
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
'    ' Crear archivo de salida
'    Dim Archivo As Integer
'    Archivo = FreeFile
'    Open CurrentProject.Path & "\procedimientos_huerfanos.txt" For Output As #Archivo
'
'    Print #Archivo, "Procedimientos huérfanos (no llamados por nadie)"
'    Print #Archivo, String(60, "-")
'
'    ' Buscar huérfanos
'    For Each proc In Todos.Keys
'        If Not Llamados.Exists(proc) Then
'            Print #Archivo, proc
'        End If
'    Next proc
'
'    Close #Archivo
'
'    MsgBox "Análisis completado. Archivo generado: procedimientos_huerfanos.txt"
'End Sub
'

