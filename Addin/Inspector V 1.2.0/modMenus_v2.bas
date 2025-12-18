Attribute VB_Name = "modMenus_v2"

Option Compare Database
Option Explicit

'===============================================================
' Módulo: modMenus
' Gestión del menú del Inspector en el IDE de VBA
'===============================================================

Private Const TAG_INSPECTOR As String = "InspectorVBA"
Private Const FACEID_EJECUTAR As Long = 279
Private Const FACEID_REPARAR As Long = 602

'===============================================================
' SECCIÓN 1: Callbacks del Ribbon
'===============================================================

Public Sub OnEjecutarInspector(control As IRibbonControl)
    EjecutarInspector
End Sub

Public Sub OnRepararProyecto(control As IRibbonControl)
    RepararProyecto
End Sub

'===============================================================
' SECCIÓN 2: Menú del IDE
'===============================================================

'---------------------------------------------------------------
' Obtener la CommandBar del menú del IDE (solo español e inglés)
'---------------------------------------------------------------
Private Function ObtenerMenuBarVBE() As CommandBar
    Dim nombresPosibles As Variant
    Dim nombre As Variant
    Dim cb As CommandBar

    ' Únicamente los idiomas posibles en tu entorno
    nombresPosibles = Array("Barra de menús", "Menu Bar")

    For Each nombre In nombresPosibles
        On Error Resume Next
        Set cb = Application.VBE.CommandBars(CStr(nombre))
        On Error GoTo 0

        If Not cb Is Nothing Then
            Set ObtenerMenuBarVBE = cb
            Exit Function
        End If
    Next nombre

    Set ObtenerMenuBarVBE = Nothing
End Function

'---------------------------------------------------------------
' Crear menú del Inspector en el IDE
'---------------------------------------------------------------
Public Sub CrearMenuInspectorVBE()
    Dim cb As CommandBar
    Dim menu As CommandBarPopup
    Dim ctrl As CommandBarControl

    If Not AsegurarReferenciaVBIDE() Then
        Debug.Print "No se puede crear el menú sin VBIDE."
        Exit Sub
    End If

    Set cb = ObtenerMenuBarVBE()
    If cb Is Nothing Then
        Debug.Print "No se encontró la barra de menús del IDE."
        Exit Sub
    End If

    ' Limpiar restos previos
    On Error Resume Next
    EliminarMenuInspectorVBE
    On Error GoTo 0

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
        .OnAction = "EjecutarInspector"
        .FaceId = FACEID_EJECUTAR
        .Tag = TAG_INSPECTOR
    End With

    ' Botón: Reparar Proyecto
    Set ctrl = menu.Controls.Add(Type:=msoControlButton)
    With ctrl
        .Caption = "Reparar Proyecto"
        .OnAction = "RepararProyecto"
        .FaceId = FACEID_REPARAR
        .Tag = TAG_INSPECTOR
    End With

    Debug.Print "Menú del Inspector creado correctamente."
End Sub

'---------------------------------------------------------------
' Eliminar menú del Inspector en el IDE
'---------------------------------------------------------------
Public Sub EliminarMenuInspectorVBE()
    Dim cb As CommandBar
    Dim ctrl As CommandBarControl
    Dim i As Long

    Set cb = ObtenerMenuBarVBE()
    If cb Is Nothing Then Exit Sub

    On Error Resume Next
    For i = cb.Controls.Count To 1 Step -1
        Set ctrl = cb.Controls(i)
        If ctrl.Tag = TAG_INSPECTOR Then ctrl.Delete
    Next i
    On Error GoTo 0
End Sub

