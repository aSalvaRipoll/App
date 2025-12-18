VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAddin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'=====================================================
' Clase: clsAddin
' Representa un Add-In ACCDA/MDA instalado en Access
'=====================================================

'-------------------------
' Datos básicos
'-------------------------
Public addin_Name As String      ' Nombre del Add-In (ej: InspectorVBA.accda)
Public library As String         ' Ruta completa del archivo

'-------------------------
' Información derivada
'-------------------------
Public BaseName As String        ' Nombre sin extensión
Public Extension As String       ' Extensión (accda/mda)
Public Folder As String          ' Carpeta donde está el Add-In
Public Exists As Boolean         ' ¿Existe físicamente el archivo?
Public Loaded As Boolean         ' ¿Está cargado en Access?
Public Version As String         ' Versión del archivo
Public Description As String     ' Descripción del archivo

'=====================================================
' Inicialización manual (llamada desde AnalizaRegistro)
'=====================================================
Public Sub Initialize()
    On Error Resume Next

    ' Nombre sin extensión
    BaseName = Left$(addin_Name, InStrRev(addin_Name, ".") - 1)

    ' Extensión
    Extension = Mid$(addin_Name, InStrRev(addin_Name, ".") + 1)

    ' Carpeta
    Folder = Left$(library, InStrRev(library, "\") - 1)

    ' ¿Existe el archivo?
    Exists = (Len(Dir$(library)) > 0)

    ' Versión y descripción (si existe)
    If Exists Then
        Version = CreateObject("Scripting.FileSystemObject").GetFileVersion(library)
        Description = CreateObject("Scripting.FileSystemObject").GetFile(library).Type
    End If

    ' ¿Está cargado en Access?
    Loaded = DetectLoaded()
End Sub

'=====================================================
' Detección de carga mediante VBIDE
'=====================================================
Private Function DetectLoaded() As Boolean
    Dim vbProj As Object
    Dim nombreSinExt As String

    nombreSinExt = LCase$(BaseName)

    On Error Resume Next
    For Each vbProj In Application.VBE.VBProjects
        If LCase$(vbProj.Name) = nombreSinExt Then
            DetectLoaded = True
            Exit Function
        End If
    Next vbProj
End Function

'=====================================================
' Acciones opcionales
'=====================================================
Public Sub Load()
    ' Preparado para futuro AddInManager
    ' No hace nada por ahora
End Sub

Public Sub Unload()
    ' Preparado para futuro AddInManager
    ' No hace nada por ahora
End Sub

