Attribute VB_Name = "ModMenus_old"

'=====================================================
' Este módulo gestiona toda la interfaz del Inspector:
' - Ribbon de Access
' - Menú del IDE (CommandBars)
'=====================================================

Option Compare Database
Option Explicit

'' Constantes de iconos
'Private Const FACEID_EJECUTAR As Long = 279
'Private Const FACEID_REPARAR As Long = 602
'
'' Etiqueta común para identificar todos los controles del Inspector
'Private Const TAG_INSPECTOR As String = "InspectorVBA"

' Iconos del menú
Private Const FACEID_EJECUTAR As Long = 279
Private Const FACEID_REPARAR As Long = 602

' Etiqueta común para identificar todos los controles del Inspector
Private Const TAG_INSPECTOR As String = "InspectorVBA"


'=====================================================
' Gestión de menús del Ribbon
'=====================================================

' Callback del botón "Ejecutar Inspector"
Public Sub OnEjecutarInspector(control As IRibbonControl)
    Debug.Print "Botón pulsado desde la cinta:", control.ID
    Call EjecutarInspectorProyecto
End Sub

' Callback del botón "Reparar Proyecto"
Public Sub OnRepararProyecto(control As IRibbonControl)
    Call RepararProblemasProyecto
End Sub

'=====================================================
' Gestión de menús del IDE (CommandBars)
'=====================================================

'=====================================================
' Gestión de menús del IDE (CommandBars)
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



''-----------------------------------------------------
'' Crear menú del Inspector en el IDE
''-----------------------------------------------------
'Public Sub CrearMenuInspectorVBE()
'    Dim cb As CommandBar
'    Dim menu As CommandBarPopup
'    Dim ctrl As CommandBarControl
'
'    ' Eliminar menú previo si existe
'    On Error Resume Next
'    EliminarMenuInspectorVBE
'    On Error GoTo 0
'
'    ' Barra de menús del IDE (en español)
'    Set cb = Application.VBE.CommandBars("Barra de menús")
'
'    ' Crear menú principal
'    Set menu = cb.Controls.Add(Type:=msoControlPopup, Temporary:=True)
'    With menu
'        .Caption = "Inspector VBA"
'        .Tag = TAG_INSPECTOR
'    End With
'
'    ' Botón: Ejecutar Inspector
'    Set ctrl = menu.Controls.Add(Type:=msoControlButton)
'    With ctrl
'        .Caption = "Ejecutar Inspector"
'        .OnAction = "EjecutarInspectorProyecto"
'        .FaceId = FACEID_EJECUTAR
'        .Tag = TAG_INSPECTOR
'    End With
'
'    ' Botón: Reparar Proyecto
'    Set ctrl = menu.Controls.Add(Type:=msoControlButton)
'    With ctrl
'        .Caption = "Reparar Proyecto"
'        .OnAction = "RepararProblemasProyecto"
'        .FaceId = FACEID_REPARAR
'        .Tag = TAG_INSPECTOR
'    End With
'End Sub
'
''-----------------------------------------------------
'' Eliminar menú del Inspector en el IDE
''-----------------------------------------------------
'Public Sub EliminarMenuInspectorVBE()
'    Dim cb As CommandBar
'    Dim ctrl As CommandBarControl
'    Dim i As Long
'
'    On Error Resume Next
'
'    Set cb = Application.VBE.CommandBars("Barra de menús")
'
'    ' Recorrer controles y eliminar solo los del Inspector
'    For i = cb.Controls.Count To 1 Step -1
'        Set ctrl = cb.Controls(i)
'        If ctrl.Tag = TAG_INSPECTOR Then
'            ctrl.Delete
'        End If
'    Next i
'End Sub

''-----------------------------------------------------
'' Crear menú del Inspector en el IDE
''-----------------------------------------------------
'Public Sub CrearMenuInspectorVBE()
'    Dim cb As CommandBar
'    Dim menu As CommandBarPopup
'    Dim ctrl As CommandBarControl
'
'    ' Eliminar menú previo si existe
'    On Error Resume Next
'    EliminarMenuInspectorVBE
'    On Error GoTo 0
'
'    ' Barra de menús del IDE (en español)
'    Set cb = Application.VBE.CommandBars("Barra de menús")
'
'    ' Crear menú principal
'    Set menu = cb.Controls.Add(Type:=msoControlPopup, Temporary:=True)
'    With menu
'        .Caption = "Inspector VBA"
'        .Tag = TAG_INSPECTOR
'    End With
'
'    ' Botón: Ejecutar Inspector
'    Set ctrl = menu.Controls.Add(Type:=msoControlButton)
'    With ctrl
'        .Caption = "Ejecutar Inspector"
'        .OnAction = "EjecutarInspectorProyecto"
'        .FaceId = FACEID_EJECUTAR
'        .Tag = TAG_INSPECTOR
'    End With
'
'    ' Botón: Reparar Proyecto
'    Set ctrl = menu.Controls.Add(Type:=msoControlButton)
'    With ctrl
'        .Caption = "Reparar Proyecto"
'        .OnAction = "RepararProblemasProyecto"
'        .FaceId = FACEID_REPARAR
'        .Tag = TAG_INSPECTOR
'    End With
'End Sub
'
''-----------------------------------------------------
'' Eliminar menú del Inspector en el IDE
''-----------------------------------------------------
'Public Sub EliminarMenuInspectorVBE()
'    Dim cb As CommandBar
'    Dim ctrl As CommandBarControl
'    Dim i As Long
'
'    On Error Resume Next
'
'    Set cb = Application.VBE.CommandBars("Barra de menús")
'
'    ' Recorrer controles y eliminar solo los del Inspector
'    For i = cb.Controls.Count To 1 Step -1
'        Set ctrl = cb.Controls(i)
'        If ctrl.Tag = TAG_INSPECTOR Then
'            ctrl.Delete
'        End If
'    Next i
'End Sub

'Public Sub CrearMenuInspectorVBE()
'    Dim cb As CommandBar
'    Dim menu As CommandBarPopup
'    Dim ctrl As CommandBarControl
'
'    On Error Resume Next
'    Call EliminarMenuInspectorVBE
'    On Error GoTo 0
'
'    ' Usar nombre en español
'    Set cb = Application.VBE.CommandBars("Barra de menús")
'
'    ' Crear menú principal
'    Set menu = cb.Controls.Add(Type:=msoControlPopup, Temporary:=True)
'    menu.Caption = "Inspector VBA"
'
'    ' Botón: Ejecutar Inspector
'    Set ctrl = menu.Controls.Add(Type:=msoControlButton)
'    ctrl.Caption = "Ejecutar Inspector"
'    ctrl.OnAction = "EjecutarInspectorProyecto"
'    ctrl.FaceId = 279
'
'    ' Botón: Reparar Proyecto
'    Set ctrl = menu.Controls.Add(Type:=msoControlButton)
'    ctrl.Caption = "Reparar Proyecto"
'    ctrl.OnAction = "RepararProblemasProyecto"
'    ctrl.FaceId = 602
'End Sub



