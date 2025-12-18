Attribute VB_Name = "50_modNavegacion"

Option Compare Database
Option Explicit

'===============================================================
' Módulo: modNavegación
' Funciones para navegar a módulos, procedimientos y líneas
'===============================================================

Public Sub NavegarAModulo(nombreModulo As String, _
                          Optional nombreProcedimiento As String = "", _
                          Optional linea As Long = 1)

    Dim vbComp As VBIDE.VBComponent
    Dim vbMod As VBIDE.CodeModule

    On Error GoTo ErrHandler

    Set vbComp = Application.VBE.ActiveVBProject.VBComponents(nombreModulo)
    Set vbMod = vbComp.CodeModule

    Application.VBE.MainWindow.Visible = True
    DoEvents

    vbMod.CodePane.Show

    If nombreProcedimiento <> "" Then
        vbMod.CodePane.SetSelection _
            vbMod.ProcStartLine(nombreProcedimiento, vbext_pk_Proc), 1, _
            vbMod.ProcStartLine(nombreProcedimiento, vbext_pk_Proc), 1
    Else
        vbMod.CodePane.SetSelection linea, 1, linea, 1
    End If

    Exit Sub

ErrHandler:
    Inspector_Log "Error en NavegarAModulo: " & Err.Description
End Sub


'Public Sub NavegarAModulo(nombreModulo As String, _
'                          Optional nombreProcedimiento As String = "", _
'                          Optional linea As Long = 1)
'
'    Dim vbComp As VBIDE.VBComponent
'    Dim vbMod As VBIDE.CodeModule
'
'    On Error GoTo ErrHandler
'
'    Set vbComp = Application.VBE.ActiveVBProject.VBComponents(nombreModulo)
'    Set vbMod = vbComp.CodeModule
'
'    Application.VBE.MainWindow.Visible = True
'    DoEvents
'
'    vbMod.CodePane.Show
'
'    If nombreProcedimiento <> "" Then
'        vbMod.CodePane.SetSelection _
'            vbMod.ProcStartLine(nombreProcedimiento, vbext_pk_Proc), 1, _
'            vbMod.ProcStartLine(nombreProcedimiento, vbext_pk_Proc), 1
'    Else
'        vbMod.CodePane.SetSelection linea, 1, linea, 1
'    End If
'
'    Exit Sub
'
'ErrHandler:
'    Inspector_Log "Error en NavegarAModulo: " & Err.Description
'End Sub
