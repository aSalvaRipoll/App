Attribute VB_Name = "modAuxAnalisis"

Option Compare Database
Option Explicit

'===============================================================
' Módulo: modAuxAnalisis
' Motor moderno de análisis del proyecto VBA
'===============================================================

'---------------------------------------------------------------
' Analiza el proyecto completo y devuelve un catálogo
'---------------------------------------------------------------
Public Function AnalizarProyecto() As clsCatalogoInspector
    Dim cat As New clsCatalogoInspector
    Dim vbComp As VBIDE.VBComponent
    Dim tipo As VBIDE.vbext_ComponentType

    For Each vbComp In Application.VBE.ActiveVBProject.VBComponents
        tipo = vbComp.Type

        Select Case tipo

            Case vbext_ct_StdModule
                cat.AgregarModulo AnalizarModulo(vbComp)

            Case vbext_ct_ClassModule
                cat.AgregarClase AnalizarClase(vbComp)

            Case vbext_ct_MSForm
                cat.AgregarUserForm AnalizarModulo(vbComp)

            Case vbext_ct_Document
                Select Case TipoModuloObjeto(vbComp)
                    Case "Formulario": cat.AgregarFormulario AnalizarModulo(vbComp)
                    Case "Informe":    cat.AgregarInforme AnalizarModulo(vbComp)
                    Case Else:         cat.AgregarOtro AnalizarModulo(vbComp)
                End Select

        End Select
    Next vbComp

    Set AnalizarProyecto = cat
End Function

'---------------------------------------------------------------
' Determina si un módulo de documento es Formulario/Informe/Otro
'---------------------------------------------------------------
Private Function TipoModuloObjeto(vbComp As VBIDE.VBComponent) As String
    Dim nombre As String
    nombre = vbComp.Name

    On Error Resume Next

    If Not CurrentProject.AllForms(nombre) Is Nothing Then
        TipoModuloObjeto = "Formulario"
        Exit Function
    End If

    If Not CurrentProject.AllReports(nombre) Is Nothing Then
        TipoModuloObjeto = "Informe"
        Exit Function
    End If

    TipoModuloObjeto = "Otro"
End Function

'---------------------------------------------------------------
' Analiza un módulo estándar o formulario
'---------------------------------------------------------------
Public Function AnalizarModulo(vbComp As VBIDE.VBComponent) As clsModulo
    Dim m As New clsModulo
    Dim code As VBIDE.CodeModule
    Dim total As Long

    m.nombre = vbComp.Name
    m.tipo = TipoModuloInspectorStd(vbComp)

    Set code = vbComp.CodeModule
    total = code.CountOfLines

    If total > 0 Then
        m.Lineas = Split(code.Lines(1, total), vbCrLf)
    End If

    m.CalcularMetricasModulo
    Set m.Miembros = DetectarMiembros(code)

    Set AnalizarModulo = m
End Function

'---------------------------------------------------------------
' Analiza una clase
'---------------------------------------------------------------
Public Function AnalizarClase(vbComp As VBIDE.VBComponent) As clsClase
    Dim c As New clsClase
    Dim code As VBIDE.CodeModule
    Dim total As Long

    c.nombre = vbComp.Name

    Set code = vbComp.CodeModule
    total = code.CountOfLines

    If total > 0 Then
        c.Lineas = Split(code.Lines(1, total), vbCrLf)
    End If

    c.CalcularMetricasClase
    Set c.Miembros = DetectarMiembros(code)

    Set AnalizarClase = c
End Function

'---------------------------------------------------------------
' Determina el tipo de módulo estándar
'---------------------------------------------------------------
Private Function TipoModuloInspectorStd(vbComp As VBIDE.VBComponent) As TipoModuloInspector
    Select Case vbComp.Type
        Case vbext_ct_StdModule: TipoModuloInspectorStd = tmStdModule
        Case vbext_ct_MSForm:    TipoModuloInspectorStd = tmUserForm
        Case vbext_ct_Document:  TipoModuloInspectorStd = tmUnknownModule
        Case Else:               TipoModuloInspectorStd = tmUnknownModule
    End Select
End Function

'---------------------------------------------------------------
' Detecta miembros (Sub/Function/Property/Event/Declare)
'---------------------------------------------------------------
Public Function DetectarMiembros(code As VBIDE.CodeModule) As Collection
    Dim col As New Collection
    Dim i As Long
    Dim linea As String
    Dim firma As String
    Dim nombre As String
    Dim kind As VBIDE.vbext_ProcKind
    Dim m As clsMiembro

    For i = 1 To code.CountOfLines
        linea = Trim$(code.Lines(i, 1))

        If EsInicioMiembro(linea) Then
            firma = linea

            Set m = New clsMiembro
            m.nombre = ExtraerNombreMiembro(firma)
            m.EstablecerTipoDesdeTexto firma
            m.EstablecerAmbitoDesdeTexto firma

            kind = code.ProcKind(m.nombre)
            m.LineaInicio = code.ProcStartLine(m.nombre, kind)
            m.NumLineas = code.ProcCountLines(m.nombre, kind)
            m.LineaFin = m.LineaInicio + m.NumLineas - 1

            CalcularMetricasMiembro code, m

            col.Add m
        End If
    Next i

    Set DetectarMiembros = col
End Function

'---------------------------------------------------------------
' ¿Es una línea que inicia un miembro?
'---------------------------------------------------------------
Private Function EsInicioMiembro(ByVal linea As String) As Boolean
    Dim t As String
    t = LCase$(Trim$(linea))

    EsInicioMiembro = _
           Left$(t, 3) = "sub" _
        Or Left$(t, 8) = "function" _
        Or Left$(t, 12) = "property get" _
        Or Left$(t, 12) = "property let" _
        Or Left$(t, 12) = "property set" _
        Or Left$(t, 5) = "event" _
        Or Left$(t, 7) = "declare" _
        Or InStr(t, " sub ") > 0 _
        Or InStr(t, " function ") > 0 _
        Or InStr(t, " property ") > 0
End Function

'---------------------------------------------------------------
' Extrae el nombre del miembro desde la firma
'---------------------------------------------------------------
Private Function ExtraerNombreMiembro(ByVal firma As String) As String
    Dim tmp As String
    Dim p As Long

    tmp = Replace(firma, "(", " ")
    tmp = Replace(tmp, ")", " ")
    tmp = Replace(tmp, "As ", " As ")
    tmp = Trim$(tmp)

    Dim parts() As String
    parts = Split(tmp, " ")

    ' El nombre suele ser la última palabra antes de "(" o "As"
    If UBound(parts) >= 1 Then
        ExtraerNombreMiembro = parts(UBound(parts) - 1)
    Else
        ExtraerNombreMiembro = parts(0)
    End If
End Function

'---------------------------------------------------------------
' Calcula métricas de un miembro
'---------------------------------------------------------------
Private Sub CalcularMetricasMiembro(code As VBIDE.CodeModule, m As clsMiembro)
    Dim i As Long
    Dim linea As String
    Dim tipo As String

    For i = m.LineaInicio To m.LineaFin
        linea = code.Lines(i, 1)
        tipo = ClasificarLinea(linea)
        m.IncrementarMetricas tipo
    Next i
End Sub

'---------------------------------------------------------------
' Clasifica una línea
'---------------------------------------------------------------
Private Function ClasificarLinea(ByVal linea As String) As String
    Dim t As String
    t = Trim$(linea)

    If t = "" Then
        ClasificarLinea = "Vacia"
    ElseIf Left$(t, 1) = "'" Or LCase$(Left$(t, 3)) = "rem" Then
        ClasificarLinea = "Comentario"
    Else
        ClasificarLinea = "Codigo"
    End If
End Function


