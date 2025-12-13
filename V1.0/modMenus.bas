Attribute VB_Name = "modMenus"

'=====================================================
' Módulo: modMenus
' Gestión unificada de la interfaz del Inspector VBA:
'   - Callbacks del Ribbon de Access
'   - Menú del IDE (CommandBars del editor VBA)
'=====================================================

Option Compare Database
Option Explicit

'-----------------------------------------------------
' Constantes comunes
'-----------------------------------------------------
Private Const TAG_INSPECTOR As String = "InspectorVBA"   ' Identificador común
Private Const FACEID_EJECUTAR As Long = 279              ' Icono: Ejecutar
Private Const FACEID_REPARAR As Long = 602               ' Icono: Reparar

'=====================================================
' SECCIÓN 1: Callbacks del Ribbon
'=====================================================

'-----------------------------------------------------
' Callback: Ejecutar Inspector desde el Ribbon
'-----------------------------------------------------
Public Sub OnEjecutarInspector(control As IRibbonControl)
    Debug.Print "Botón pulsado en Ribbon:", control.ID
    EjecutarInspectorProyecto
End Sub

'-----------------------------------------------------
' Callback: Reparar Proyecto desde el Ribbon
'-----------------------------------------------------
Public Sub OnRepararProyecto(control As IRibbonControl)
    RepararProblemasProyecto
End Sub


'=====================================================
' SECCIÓN 2: Menú del IDE (CommandBars)
'=====================================================

'-----------------------------------------------------
' Crear menú del Inspector en el IDE
'-----------------------------------------------------
Public Sub CrearMenuInspectorVBE()
    Dim cb As CommandBar
    Dim menu As CommandBarPopup
    Dim ctrl As CommandBarControl

    ' Eliminar menú previo si existe
    On Error Resume Next
    EliminarMenuInspectorVBE
    On Error GoTo 0

    ' Barra de menús del IDE (en español)
    Set cb = Application.VBE.CommandBars("Barra de menús")

    ' Crear menú principal
    Set menu = cb.Controls.Add(Type:=msoControlPopup, Temporary:=True)
    With menu
        .Caption = "Inspector VBA"
        .Tag = TAG_INSPECTOR
    End With

    ' Botón: Ejecutar Inspector
    Set ctrl = menu.Controls.Add(Type:=msoControlButton)
    With ctrl
        .Caption = "Ejecutar Inspector"
        .OnAction = "EjecutarInspectorProyecto"
        .FaceId = FACEID_EJECUTAR
        .Tag = TAG_INSPECTOR
    End With

    ' Botón: Reparar Proyecto
    Set ctrl = menu.Controls.Add(Type:=msoControlButton)
    With ctrl
        .Caption = "Reparar Proyecto"
        .OnAction = "RepararProblemasProyecto"
        .FaceId = FACEID_REPARAR
        .Tag = TAG_INSPECTOR
    End With
End Sub

'-----------------------------------------------------
' Eliminar menú del Inspector en el IDE
'-----------------------------------------------------
Public Sub EliminarMenuInspectorVBE()
    Dim cb As CommandBar
    Dim ctrl As CommandBarControl
    Dim i As Long

    On Error Resume Next

    Set cb = Application.VBE.CommandBars("Barra de menús")

    ' Recorrer controles y eliminar solo los del Inspector
    For i = cb.Controls.Count To 1 Step -1
        Set ctrl = cb.Controls(i)
        If ctrl.Tag = TAG_INSPECTOR Then
            ctrl.Delete
        End If
    Next i
End Sub

