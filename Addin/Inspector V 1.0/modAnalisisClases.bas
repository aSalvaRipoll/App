Attribute VB_Name = "modAnalisisClases"

Option Compare Database
Option Explicit

'=====================================================
' Módulo: modAnalisisClases
' Análisis de clases VBA
'=====================================================

Public Function ObtenerClasesVBA() As Collection
    Dim col As New Collection
    Dim vbComp As VBIDE.VBComponent

    For Each vbComp In Application.VBE.ActiveVBProject.VBComponents
        If vbComp.Type = vbext_ct_ClassModule Then
            col.Add vbComp
        End If
    Next vbComp

    Set ObtenerClasesVBA = col
End Function

Public Function ObtenerMiembrosClase(vbComp As VBIDE.VBComponent) As Collection
    Dim col As New Collection
    Dim code As VBIDE.CodeModule
    Dim i As Long
    Dim linea As String

    Set code = vbComp.CodeModule

    For i = 1 To code.CountOfLines
        linea = Trim(code.Lines(i, 1))

        ' Detectar métodos y propiedades
        If linea Like "Public *" Or linea Like "Private *" Then
            col.Add linea
        End If
    Next i

    Set ObtenerMiembrosClase = col
End Function

