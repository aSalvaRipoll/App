Attribute VB_Name = "modAnalisisModulos"

Option Compare Database
Option Explicit

'=====================================================
' Módulo: modAnalisisModulos
' Análisis de módulos estándar y formularios
'=====================================================

Public Function ObtenerModulosVBA() As Collection
    Dim col As New Collection
    Dim vbComp As VBIDE.VBComponent

    For Each vbComp In Application.VBE.ActiveVBProject.VBComponents
        If vbComp.Type = vbext_ct_StdModule Or vbComp.Type = vbext_ct_MSForm Then
            col.Add vbComp
        End If
    Next vbComp

    Set ObtenerModulosVBA = col
End Function

Public Function ObtenerLineasModulo(vbComp As VBIDE.VBComponent) As Variant
    Dim code As String
    code = vbComp.CodeModule.Lines(1, vbComp.CodeModule.CountOfLines)
    ObtenerLineasModulo = Split(code, vbCrLf)
End Function

Public Function AnalizarModulo(vbComp As VBIDE.VBComponent) As clsModulo
    Dim m As New clsModulo
    Dim code As VBIDE.CodeModule
    Dim total As Long
    Dim i As Long
    Dim linea As String
    Dim mi As clsMiembro

    m.nombre = vbComp.Name
    m.Tipo = TipoVBComponent(vbComp)

    Set code = vbComp.CodeModule
    total = code.CountOfLines
    m.NumLineas = total

    If total > 0 Then
        m.Lineas = Split(code.Lines(1, total), vbCrLf)

        For i = 1 To total
            linea = code.Lines(i, 1)
            Select Case ClasificarLinea(linea)
                Case "Codigo":        m.NumLineasCodigo = m.NumLineasCodigo + 1
                Case "Comentario":    m.NumLineasComentario = m.NumLineasComentario + 1
                Case "Vacia":         m.NumLineasVacias = m.NumLineasVacias + 1
            End Select
        Next i
    End If

    Set m.Miembros = DetectarMiembros(code)

    m.NumMiembros = m.Miembros.Count
    For Each mi In m.Miembros
        If mi.EsProcedimiento Then m.NumProcedimientos = m.NumProcedimientos + 1
        If mi.EsFuncion Then m.NumFunciones = m.NumFunciones + 1
        If mi.EsPropiedad Then m.NumPropiedades = m.NumPropiedades + 1
        If mi.EsEvento Then m.NumEventos = m.NumEventos + 1
    Next mi

    Set AnalizarModulo = m
End Function


Public Function AnalizarModulo(vbComp As VBIDE.VBComponent) As clsModulo
    Dim m As New clsModulo
    Dim code As VBIDE.CodeModule
    Dim total As Long
    Dim i As Long
    Dim linea As String
    Dim mi As clsMiembro

    m.nombre = vbComp.Name
    m.Tipo = TipoVBComponent(vbComp)

    Set code = vbComp.CodeModule
    total = code.CountOfLines
    m.NumLineas = total

    If total > 0 Then
        m.Lineas = Split(code.Lines(1, total), vbCrLf)

        For i = 1 To total
            linea = code.Lines(i, 1)
            Select Case ClasificarLinea(linea)
                Case "Codigo":        m.NumLineasCodigo = m.NumLineasCodigo + 1
                Case "Comentario":    m.NumLineasComentario = m.NumLineasComentario + 1
                Case "Vacia":         m.NumLineasVacias = m.NumLineasVacias + 1
            End Select
        Next i
    End If

    Set m.Miembros = DetectarMiembros(code)

    m.NumMiembros = m.Miembros.Count
    For Each mi In m.Miembros
        If mi.EsProcedimiento Then m.NumProcedimientos = m.NumProcedimientos + 1
        If mi.EsFuncion Then m.NumFunciones = m.NumFunciones + 1
        If mi.EsPropiedad Then m.NumPropiedades = m.NumPropiedades + 1
        If mi.EsEvento Then m.NumEventos = m.NumEventos + 1
    Next mi

    Set AnalizarModulo = m
End Function

