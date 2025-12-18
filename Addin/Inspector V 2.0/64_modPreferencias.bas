Attribute VB_Name = "64_modPreferencias"

Option Compare Database
Option Explicit

' ============================================================
'  MÓDULO DE PREFERENCIAS DEL INSPECTOR
'  Guarda y recupera configuraciones del usuario.
' ============================================================


' ------------------------------------------------------------
' Guarda preferencias generales del Inspector
' ------------------------------------------------------------
Public Sub GuardarPreferenciasInspector(frm As Form)
    On Error Resume Next

    ' Guardar si la ventana estaba maximizada
    SaveSetting "InspectorVBA", "Preferencias", "VentanaMaximizada", _
                IIf(IsWindowMaximized(frm.hwnd), "1", "0")

    ' Guardar última ruta usada
    If Not IsNull(frm!txtUltimaRuta) Then
        SaveSetting "InspectorVBA", "Preferencias", "UltimaRuta", frm!txtUltimaRuta
    End If
End Sub


' ------------------------------------------------------------
' Guarda preferencias del panel de exportación
' ------------------------------------------------------------
Public Sub GuardarPreferenciasExportacion(frm As Form)
    On Error Resume Next

    ' Formato seleccionado
    If Not IsNull(frm!cboFormato) Then
        SaveSetting "InspectorVBA", "Exportacion", "Formato", frm!cboFormato
    End If

    ' Estilo HTML seleccionado
    If Not IsNull(frm!cboEstilo) Then
        SaveSetting "InspectorVBA", "Exportacion", "Estilo", frm!cboEstilo
    End If

    ' Ruta de destino
    If Not IsNull(frm!txtRutaDestino) Then
        SaveSetting "InspectorVBA", "Exportacion", "Ruta", frm!txtRutaDestino
    End If
End Sub


' ============================================================
'  UTILIDADES
' ============================================================

' Comprueba si una ventana está maximizada
Private Function IsWindowMaximized(hwnd As Long) As Boolean
    On Error Resume Next

    Dim placement As WINDOWPLACEMENT
    placement.Length = Len(placement)

    If GetWindowPlacement(hwnd, placement) Then
        IsWindowMaximized = (placement.showCmd = SW_SHOWMAXIMIZED)
    End If
End Function


