Attribute VB_Name = "modRibbonInspector"

Option Compare Database
Option Explicit

' ============================================================
' Módulo: modRibbonInspector
' Callbacks de la cinta del Inspector VBA
' ============================================================

' --- Acciones principales ---
Public Sub Ribbon_Analizar(control As IRibbonControl)
    Form_frmInspector.Analizar
End Sub

Public Sub Ribbon_Reparar(control As IRibbonControl)
    Form_frmInspector.Reparar
End Sub

Public Sub Ribbon_AbrirCarpetaLogs(control As IRibbonControl)
    Inspector_AbrirCarpetaLogs
End Sub

Public Sub Ribbon_Cerrar(control As IRibbonControl)
    If CurrentProject.AllForms("frmInspector").IsLoaded Then
        DoCmd.Close acForm, "frmInspector"
    End If
End Sub

' --- Exportación ---
Public Sub Ribbon_Examinar(control As IRibbonControl)
    Form_subExportarInspector.cmdExaminar_Click
End Sub

Public Sub Ribbon_Exportar(control As IRibbonControl)
    Dim formato As String
    Dim estilo As String

    formato = Ribbon_Formato_Selected(Nothing)
    estilo = Ribbon_Estilo_Selected(Nothing)

    Form_subExportarInspector.Exportar formato, estilo
End Sub

' --- Visibilidad dinámica ---
Public Function Ribbon_MostrarEstilo(control As IRibbonControl) As Boolean
    Ribbon_MostrarEstilo = (Form_subExportarInspector.FormatoActual = ExpTodoHTML)
End Function

' ============================================================
' DropDown: Formatos
' ============================================================
Private gFormatos As Variant

Public Function Ribbon_Formato_Count(control As IRibbonControl) As Integer
    gFormatos = Array("TXT", "XLSX", "HTML")
    Ribbon_Formato_Count = UBound(gFormatos) + 1
End Function

Public Function Ribbon_Formato_ID(control As IRibbonControl, index As Integer) As String
    Ribbon_Formato_ID = gFormatos(index)
End Function

Public Function Ribbon_Formato_Label(control As IRibbonControl, index As Integer) As String
    Ribbon_Formato_Label = gFormatos(index)
End Function

Public Function Ribbon_Formato_Selected(control As IRibbonControl) As String
    Ribbon_Formato_Selected = Form_subExportarInspector.FormatoActual
End Function

' ============================================================
' DropDown: Estilos HTML
' ============================================================
Private gEstilos As Variant

Public Function Ribbon_Estilo_Count(control As IRibbonControl) As Integer
    gEstilos = Array("Claro", "Oscuro", "Minimalista")
    Ribbon_Estilo_Count = UBound(gEstilos) + 1
End Function

Public Function Ribbon_Estilo_ID(control As IRibbonControl, index As Integer) As String
    Ribbon_Estilo_ID = gEstilos(index)
End Function

Public Function Ribbon_Estilo_Label(control As IRibbonControl, index As Integer) As String
    Ribbon_Estilo_Label = gEstilos(index)
End Function

Public Function Ribbon_Estilo_Selected(control As IRibbonControl) As String
    Ribbon_Estilo_Selected = Form_subExportarInspector.EstiloActual
End Function

' ============================================================
' Cinta global: lanzador
' ============================================================
Public Sub Ribbon_AbrirInspector(control As IRibbonControl)
    DoCmd.OpenForm "frmInspector"
End Sub

Public Function Ribbon_InspectorVBA_Visible(control As IRibbonControl) As Boolean
    Ribbon_InspectorVBA_Visible = Not CurrentProject.AllForms("frmInspector").IsLoaded
End Function

