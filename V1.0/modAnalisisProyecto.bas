Attribute VB_Name = "modAnalisisProyecto"

Option Compare Database
Option Explicit

'=====================================================
' Módulo: modAnalisisProyecto
' Motor principal del análisis del proyecto VBA
'=====================================================

Public Function AnalizarProyecto() As clsCatalogoInspector
    Dim cat As New clsCatalogoInspector
    Dim vbComp As VBIDE.VBComponent

    For Each vbComp In Application.VBE.ActiveVBProject.VBComponents
        Select Case vbComp.Type

            Case vbext_ct_StdModule
                cat.Modulos.Add AnalizarModulo(vbComp)

            Case vbext_ct_ClassModule
                cat.Clases.Add AnalizarClase(vbComp)

            Case vbext_ct_MSForm
                cat.UserForms.Add AnalizarModulo(vbComp)

            Case vbext_ct_Document
                Select Case TipoModuloObjeto(vbComp)
                    Case "Formulario"
                        cat.Formularios.Add AnalizarModulo(vbComp)
                    Case "Informe"
                        cat.Informes.Add AnalizarModulo(vbComp)
                    Case Else
                        cat.Otros.Add AnalizarModulo(vbComp)
                End Select

        End Select
    Next vbComp

    Set AnalizarProyecto = cat
End Function


Public Function TipoModuloObjeto(vbComp As VBIDE.VBComponent) As String
    Dim nombre As String
    nombre = vbComp.Name

    ' Prefijos estándar
    If Left$(nombre, 5) = "Form_" Then
        TipoModuloObjeto = "Formulario"
        Exit Function
    End If

    If Left$(nombre, 7) = "Report_" Then
        TipoModuloObjeto = "Informe"
        Exit Function
    End If

    ' Comprobación en colecciones Access
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


Public Function AnalizarModulo(vbComp As VBIDE.VBComponent) As clsModulo
    Dim m As New clsModulo
    Dim code As VBIDE.CodeModule
    Dim total As Long

    m.nombre = vbComp.Name
    m.Tipo = TipoVBComponent(vbComp)

    Set code = vbComp.CodeModule
    total = code.CountOfLines

    If total > 0 Then
        m.Lineas = Split(code.Lines(1, total), vbCrLf)
    End If

    ' Detectar miembros
    Set m.Miembros = DetectarMiembros(code)

    Set AnalizarModulo = m
End Function


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

    Set c.Miembros = DetectarMiembros(code)

    Set AnalizarClase = c
End Function


Public Function DetectarMiembros(code As VBIDE.CodeModule) As Collection
    Dim col As New Collection
    Dim i As Long
    Dim linea As String
    Dim m As clsMiembro

    For i = 1 To code.CountOfLines
        linea = Trim(code.Lines(i, 1))

        If EsInicioMiembro(linea) Then
            Set m = New clsMiembro
            m.nombre = ExtraerNombreMiembro(linea)
            m.Tipo = ExtraerTipoMiembro(linea)
            m.Ambito = ExtraerAmbitoMiembro(linea)
            m.lineaInicio = i
            m.LineaFin = code.ProcBodyLine(m.nombre, code.ProcKind(m.nombre)) _
                         + code.ProcCountLines(m.nombre, code.ProcKind(m.nombre)) - 1
            m.NumLineas = m.LineaFin - m.lineaInicio + 1
            col.Add m
        End If
    Next i

    Set DetectarMiembros = col
End Function


Private Function EsInicioMiembro(linea As String) As Boolean
    linea = LCase$(linea)
    EsInicioMiembro = _
           Left$(linea, 3) = "sub" _
        Or Left$(linea, 8) = "private " _
        Or Left$(linea, 7) = "public " _
        Or Left$(linea, 8) = "friend " _
        Or Left$(linea, 8) = "function" _
        Or InStr(linea, "property ") > 0
End Function

Private Function ExtraerNombreMiembro(linea As String) As String
    Dim p As Long
    linea = Replace(linea, "(", " ")
    linea = Replace(linea, ")", " ")
    linea = Replace(linea, "As ", " As ")
    linea = Trim(linea)

    Dim parts() As String
    parts = Split(linea, " ")

    ' El nombre suele ser la última palabra antes de "(" o "As"
    ExtraerNombreMiembro = parts(UBound(parts) - 1)
End Function

Private Function ExtraerTipoMiembro(linea As String) As String
    linea = LCase$(linea)
    If InStr(linea, "property get") Then ExtraerTipoMiembro = "Property Get": Exit Function
    If InStr(linea, "property let") Then ExtraerTipoMiembro = "Property Let": Exit Function
    If InStr(linea, "property set") Then ExtraerTipoMiembro = "Property Set": Exit Function
    If Left$(linea, 3) = "sub" Or InStr(linea, " sub ") Then ExtraerTipoMiembro = "Sub": Exit Function
    If Left$(linea, 8) = "function" Or InStr(linea, " function ") Then ExtraerTipoMiembro = "Function": Exit Function
End Function

Private Function ExtraerAmbitoMiembro(linea As String) As String
    linea = LCase$(linea)
    If Left$(linea, 6) = "public" Then ExtraerAmbitoMiembro = "Public": Exit Function
    If Left$(linea, 7) = "private" Then ExtraerAmbitoMiembro = "Private": Exit Function
    If Left$(linea, 6) = "friend" Then ExtraerAmbitoMiembro = "Friend": Exit Function
    ExtraerAmbitoMiembro = "Public"
End Function

