Attribute VB_Name = "modInicioFin_old"

Option Compare Database
Option Explicit

'=====================================================
' Módulo: modInicioFin
' Ciclo de vida del Inspector VBA como complemento
'=====================================================

'-----------------------------------------------------
' Determina si el Inspector está cargado como complemento
'-----------------------------------------------------
Public Function EsComplemento() As Boolean
    On Error Resume Next

    Dim addin As addin
    For Each addin In Application.VBE.Addins
        If addin.Description = "Inspector VBA" Then
            EsComplemento = addin.Connected
            Exit Function
        End If
    Next addin
End Function

'-----------------------------------------------------
' Inicio del Inspector (solo como complemento)
'-----------------------------------------------------
Public Sub InicioInspector()
    ' Solo actuamos si estamos cargados como ACCDA
    If Not EsComplemento() Then Exit Sub

    ' Asegurar referencia VBIDE activa
    If Not ReferenciaExtensibilidadActiva() Then
        ActivarReferenciaExtensibilidad
    End If

    ' Crear menú del IDE
    CrearMenuInspectorVBE
End Sub

'-----------------------------------------------------
' Fin del Inspector (solo como complemento)
'-----------------------------------------------------
Public Sub FinInspector()
    On Error Resume Next

    If EsComplemento() Then
        EliminarMenuInspectorVBE
    End If
End Sub

