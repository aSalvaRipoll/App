Attribute VB_Name = "modInicioFin_old"
Option Compare Database
Option Explicit

Public colAddins As Collection

Public Sub AutoInicioInspector()
    Set colAddins = Nothing
    If EsComplemento() Then
        CrearMenuInspectorVBE
    End If
End Sub

Public Function EsComplemento() As Boolean
    Dim ai As clsAddin
    Dim rutaActual As String

    rutaActual = LCase(CurrentProject.FullName)

    ' Asegurarnos de que colAddins está cargada
    If colAddins Is Nothing Then
        ListaComplementosAccess   ' Esto rellena colAddins
    End If

    ' Comparar solo el nombre base sin extensión
    rutaActual = Replace(rutaActual, ".accdb", "")
    rutaActual = Replace(rutaActual, ".accda", "")

    For Each ai In colAddins
        Dim rutaAddin As String
        rutaAddin = LCase(ai.library)
        rutaAddin = Replace(rutaAddin, ".accdb", "")
        rutaAddin = Replace(rutaAddin, ".accda", "")

        If rutaAddin = rutaActual Then
            EsComplemento = True
            Exit Function
        End If
    Next ai
End Function


Public Sub EliminarMenuInspectorVBE()
    Dim cb As CommandBar
    Dim ctrl As CommandBarControl

    On Error Resume Next
    Set cb = Application.VBE.CommandBars("Barra de menús")

    For Each ctrl In cb.Controls
        If ctrl.Caption = "Inspector VBA" Then
            ctrl.Delete
            Exit For
        End If
    Next ctrl
End Sub

'Public Function EsComplemento() As Boolean
'    Dim ai As clsAddin
'
'    ' Asegurarnos de que colAddins está cargada
'    If colAddins Is Nothing Then
'        ListaComplementosAccess   ' Esto rellena colAddins
'    End If
'
'    For Each ai In colAddins
'        If LCase(ai.library) = LCase(CurrentProject.FullName) Then
'            EsComplemento = True
'            Exit Function
'        End If
'    Next ai
'    Set colAddins = Nothing
'End Function


'Public Function EsComplemento() As Boolean
'    EsComplemento = (CurrentProject.FullName <> Application.CurrentDb.Name)
'End Function




'Public Function EsComplemento() As Boolean
'    Dim ai As addin
'    For Each ai In Access.Addins
'        If LCase(ai.FullName) = LCase(CurrentProject.FullName) Then
'            EsComplemento = True
'            Exit Function
'        End If
'    Next ai
'End Function

'Public Function EsComplemento() As Boolean
'    ' Devuelve True si este archivo está cargado como complemento
'    Dim addin As addin
'    For Each addin In Application.Addins
'        If LCase(addin.FullName) = LCase(CurrentProject.FullName) Then
'            EsComplemento = True
'            Exit Function
'        End If
'    Next addin
'End Function
