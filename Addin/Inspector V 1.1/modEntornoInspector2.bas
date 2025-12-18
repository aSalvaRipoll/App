Attribute VB_Name = "modEntornoInspector2"

Option Compare Database
Option Explicit

'===============================================================
' Módulo: modEntornoInspector
' Detección moderna del Add-In ACCDA del Inspector
'===============================================================

Private Const NOMBRE_ADDIN As String = "InspectorVBA.accda"

'---------------------------------------------------------------
' ¿Está instalado el Inspector como Add-In?
'---------------------------------------------------------------
Public Function InspectorInstalado() As Boolean
    Dim ai As Access.addin

    For Each ai In Application.Addins
        If LCase$(ai.Name) = LCase$(NOMBRE_ADDIN) Then
            InspectorInstalado = True
            Exit Function
        End If
    Next ai
End Function

''---------------------------------------------------------------
'' Devuelve la ruta completa del Add-In del Inspector
''---------------------------------------------------------------
'Public Function RutaAddinInspector() As String
'    Dim ai As Access.addin
'
'    For Each ai In Application.Addins
'        If LCase$(ai.Name) = LCase$(NOMBRE_ADDIN) Then
'            RutaAddinInspector = ai.FullName
'            Exit Function
'        End If
'    Next ai
'End Function

''---------------------------------------------------------------
'' Devuelve el objeto Add-In del Inspector
''---------------------------------------------------------------
'Public Function ObtenerAddinInspector() As Access.addin
'    Dim ai As Access.addin
'
'    For Each ai In Application.Addins
'        If LCase$(ai.Name) = LCase$(NOMBRE_ADDIN) Then
'            Set ObtenerAddinInspector = ai
'            Exit Function
'        End If
'    Next ai
'End Function

