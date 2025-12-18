Attribute VB_Name = "modInicioFin"

Option Compare Database
Option Explicit

'===============================================================
' Módulo: modInicioFin
' Ciclo de vida del Inspector cuando actúa como complemento ACCDA
'===============================================================

Public Function EsComplemento() As Boolean
    EsComplemento = InspectorInstalado()
End Function

'---------------------------------------------------------------
' Inicio del Inspector (llamado desde AutoExec del ACCDA)
'---------------------------------------------------------------
Public Sub InicioInspector()
    ' Si no está registrado como complemento, no hacemos nada
    If Not EsComplemento() Then Exit Sub

    ' Asegurar que la referencia VBIDE está activa (solo afecta al complemento)
    If Not ReferenciaExtensibilidadActiva() Then
        ActivarReferenciaExtensibilidad
    End If

    ' Crear menú del Inspector en el IDE
    On Error Resume Next
    CrearMenuInspectorVBE
End Sub

'---------------------------------------------------------------
' Fin del Inspector (por ejemplo, al descargar el complemento)
'---------------------------------------------------------------
Public Sub FinInspector()
    ' Solo limpiar si realmente somos complemento
    If Not EsComplemento() Then Exit Sub

    On Error Resume Next
    EliminarMenuInspectorVBE
End Sub


'---------------------------------------------------------------



'---------------------------------------------------------------
' ¿El Inspector está funcionando como complemento?
' Usa la lógica basada en registro (modEntornoInspector)
'---------------------------------------------------------------

'Public Function InspectorInstalado() As Boolean
'    Dim ai As clsAddin
'    Dim nombre As String
'
'    nombre = "InspectorVBA.accda"   ' Ajusta si usas otro nombre
'
'    ' Asegurar que la colección está cargada
'    If colAddins Is Nothing Then
'        ExportaRegistro
'        ListaComplementosAccess
'    End If
'
'    ' Buscar el AddIn en la colección
'    For Each ai In colAddins
'        If LCase$(ai.addin_Name) = LCase$(nombre) Then
'            InspectorInstalado = True
'            Exit Function
'        End If
'    Next ai
'End Function

