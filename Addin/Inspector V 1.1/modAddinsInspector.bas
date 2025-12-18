Attribute VB_Name = "modAddinsInspector"

Option Compare Database
Option Explicit

'===============================================================
' Módulo: modAddinsInspector
' Detección robusta de complementos de Access SIN APIs
'===============================================================

Private Const EXT_VALIDAS As String = "accda;accde;accdb;mda;mde;mdb"

'---------------------------------------------------------------
' Devuelve una colección con todos los complementos instalados
'---------------------------------------------------------------
Public Function ListaComplementosAccess() As Collection
    Dim col As New Collection
    Dim carpeta As String
    Dim f As String
    Dim rutaCompleta As String
    Dim ext As String

    carpeta = CarpetaAddinsAccess()

    If Len(Dir(carpeta, vbDirectory)) = 0 Then
        Set ListaComplementosAccess = col
        Exit Function
    End If

    f = Dir(carpeta & "\*.*")

    Do While Len(f) > 0
        ext = LCase$(Mid$(f, InStrRev(f, ".") + 1))

        If EsExtensionValida(ext) Then
            rutaCompleta = carpeta & "\" & f

            If FicheroExiste(rutaCompleta) Then
                col.Add CrearInfoAddin(rutaCompleta)
            End If
        End If

        f = Dir()
    Loop

    Set ListaComplementosAccess = col
End Function

'---------------------------------------------------------------
' Carpeta donde Access guarda los Add-Ins
'---------------------------------------------------------------
Public Function CarpetaAddinsAccess() As String
    CarpetaAddinsAccess = Environ$("APPDATA") & "\Microsoft\AddIns"
End Function

'---------------------------------------------------------------
' ¿Es una extensión válida de Add-In?
'---------------------------------------------------------------
Private Function EsExtensionValida(ext As String) As Boolean
    EsExtensionValida = (InStr(1, EXT_VALIDAS, ext) > 0)
End Function

'---------------------------------------------------------------
' ¿El fichero existe realmente?
'---------------------------------------------------------------
Private Function FicheroExiste(ruta As String) As Boolean
    FicheroExiste = (Len(Dir(ruta)) > 0)
End Function

'---------------------------------------------------------------
' Crea un objeto con la información del Add-In
'---------------------------------------------------------------
Private Function CrearInfoAddin(ruta As String) As Collection
    Dim info As New Collection
    Dim nombre As String

    nombre = Mid$(ruta, InStrRev(ruta, "\") + 1)

    info.Add nombre, "Nombre"
    info.Add ruta, "Ruta"
    info.Add EstaCargado(nombre), "Cargado"

    Set CrearInfoAddin = info
End Function

'---------------------------------------------------------------
' ¿El Add-In está cargado actualmente?
'---------------------------------------------------------------
Private Function EstaCargado(nombreAddin As String) As Boolean
    ' Si el proyecto actual coincide con el nombre del add-in, está cargado
    EstaCargado = (LCase$(CurrentProject.Name) = LCase$(nombreAddin))
End Function

