Attribute VB_Name = "03_modVBIDE"

Option Compare Database
Option Explicit

' ============================================================
'  MÓDULO: modVBIDE
'  Funciones avanzadas para interactuar con el IDE de VBA
'  - Enumerar proyectos, módulos, clases, formularios
'  - Exportar módulos
'  - Insertar código
'  - Detectar tipos de componentes
'  - Asegurar referencia VBIDE
' ============================================================


' ------------------------------------------------------------
' 1. ASEGURAR REFERENCIA VBIDE
' ------------------------------------------------------------

Public Function AsegurarReferenciaVBIDE() As Boolean
    On Error GoTo ErrHandler

    Dim ref As Reference
    For Each ref In Application.References
        If ref.Name = "VBIDE" Then
            AsegurarReferenciaVBIDE = True
            Exit Function
        End If
    Next ref

    ' Si no existe, intentar agregarla
    Application.References.AddFromGuid "{0002E157-0000-0000-C000-000000000046}", 5, 3
    AsegurarReferenciaVBIDE = True
    Exit Function

ErrHandler:
    AsegurarReferenciaVBIDE = False
End Function


' ------------------------------------------------------------
' 2. ENUMERACIÓN DE PROYECTOS Y COMPONENTES
' ------------------------------------------------------------

' Devuelve el proyecto VBA actual
Public Function ProyectoActual() As VBProject
    Set ProyectoActual = Application.VBE.ActiveVBProject
End Function

' Devuelve una colección con todos los componentes del proyecto
Public Function ComponentesProyecto() As Collection
    Dim col As New Collection
    Dim comp As VBComponent

    For Each comp In ProyectoActual.VBComponents
        col.Add comp
    Next comp

    Set ComponentesProyecto = col
End Function

' Devuelve True si existe un módulo con ese nombre
Public Function ExisteComponente(nombre As String) As Boolean
    On Error Resume Next
    ExisteComponente = Not ProyectoActual.VBComponents(nombre) Is Nothing
End Function


' ------------------------------------------------------------
' 3. DETECCIÓN DE TIPOS DE COMPONENTES
' ------------------------------------------------------------

Public Function EsModuloEstandar(comp As VBComponent) As Boolean
    EsModuloEstandar = (comp.Type = vbext_ct_StdModule)
End Function

Public Function EsClase(comp As VBComponent) As Boolean
    EsClase = (comp.Type = vbext_ct_ClassModule)
End Function

Public Function EsFormulario(comp As VBComponent) As Boolean
    EsFormulario = (comp.Type = vbext_ct_MSForm)
End Function

Public Function EsModuloDocumento(comp As VBComponent) As Boolean
    EsModuloDocumento = (comp.Type = vbext_ct_Document)
End Function


' ------------------------------------------------------------
' 4. EXPORTAR MÓDULOS
' ------------------------------------------------------------

Public Function ExportarModulo(comp As VBComponent, ruta As String) As Boolean
    On Error GoTo ErrHandler

    comp.Export ruta
    ExportarModulo = True
    Exit Function

ErrHandler:
    ExportarModulo = False
End Function

Public Sub ExportarTodosLosModulos(rutaCarpeta As String)
    Dim comp As VBComponent
    Dim ruta As String

    For Each comp In ProyectoActual.VBComponents
        ruta = rutaCarpeta & "\" & comp.Name & ".bas"
        On Error Resume Next
        comp.Export ruta
    Next comp
End Sub


' ------------------------------------------------------------
' 5. INSERTAR CÓDIGO EN MÓDULOS
' ------------------------------------------------------------

Public Function InsertarCodigo(comp As VBComponent, codigo As String) As Boolean
    On Error GoTo ErrHandler

    comp.CodeModule.AddFromString codigo
    InsertarCodigo = True
    Exit Function

ErrHandler:
    InsertarCodigo = False
End Function

Public Function InsertarLinea(comp As VBComponent, linea As String) As Boolean
    On Error GoTo ErrHandler

    With comp.CodeModule
        .InsertLines .CountOfLines + 1, linea
    End With

    InsertarLinea = True
    Exit Function

ErrHandler:
    InsertarLinea = False
End Function


' ------------------------------------------------------------
' 6. OBTENER INFORMACIÓN DEL CÓDIGO
' ------------------------------------------------------------

Public Function NumeroLineas(comp As VBComponent) As Long
    NumeroLineas = comp.CodeModule.CountOfLines
End Function

Public Function TieneProcedimiento(comp As VBComponent, nombre As String) As Boolean
    On Error Resume Next
    TieneProcedimiento = (comp.CodeModule.ProcStartLine(nombre, vbext_pk_Proc) > 0)
End Function

Public Function ObtenerProcedimiento(comp As VBComponent, nombre As String) As String
    Dim inicio As Long, numLineas As Long

    inicio = comp.CodeModule.ProcStartLine(nombre, vbext_pk_Proc)
    numLineas = comp.CodeModule.ProcCountLines(nombre, vbext_pk_Proc)

    ObtenerProcedimiento = comp.CodeModule.Lines(inicio, numLineas)
End Function

