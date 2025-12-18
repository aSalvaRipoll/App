Attribute VB_Name = "modRibbonInspector"

Option Compare Database
Option Explicit

Public Sub Ribbon_Analizar(control As IRibbonControl)
    Form_frmInspector.Analizar
End Sub

Public Sub Ribbon_Reparar(control As IRibbonControl)
    Form_frmInspector.Reparar
End Sub

Public Sub Inspector_AbrirCarpetaLogs(control As IRibbonControl)
    Dim carpeta As String
    carpeta = CurrentProject.Path & "\Logs"

    If Dir(carpeta, vbDirectory) = "" Then
        MkDir carpeta
    End If

    Shell "explorer.exe """ & carpeta & """", vbNormalFocus
End Sub

Public Sub Ribbon_Cerrar(control As IRibbonControl)
    If CurrentProject.AllForms("frmInspector").IsLoaded Then
        DoCmd.Close acForm, "frmInspector"
    End If
    'DoCmd.Close acForm, "frmInspector"
End Sub

Public Sub Ribbon_Examinar(control As IRibbonControl)
    Form_subExportarInspector.cmdExaminar_Click
End Sub

Public Sub Ribbon_Exportar(control As IRibbonControl)
    Form_subExportarInspector.cmdExportar_Click
End Sub

Public Function Ribbon_MostrarEstilo(control As IRibbonControl) As Boolean
    Ribbon_MostrarEstilo = (Form_subExportarInspector.FormatoActual = ExpTodoHTML)
End Function


'Control de la cinta estandard
Public Function Ribbon_InspectorVBA_Visible(control As IRibbonControl) As Boolean
    Ribbon_InspectorVBA_Visible = Not CurrentProject.AllForms("frmInspector").IsLoaded
End Function

