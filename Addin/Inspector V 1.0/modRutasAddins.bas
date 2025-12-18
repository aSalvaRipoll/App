Attribute VB_Name = "modRutasAddins"
Option Compare Database

Option Compare Database
Option Explicit

'=====================================================
' Módulo: modRutasAddins
' Resolución dinámica de rutas de Office y Add-Ins ACCDA
'=====================================================

'-----------------------------------------------------
' Devuelve la ruta completa del AddIn del Inspector
'-----------------------------------------------------
Public Function RutaAddinInspector() As String
    Dim ai As clsAddin

    Set ai = ObtenerAddinInspector()
    If ai Is Nothing Then Exit Function

    RutaAddinInspector = ObtenerRutaCompletaAddin(ai.library)
End Function

'-----------------------------------------------------
' Busca un archivo ACCDA/ACCDE/ACCDB en rutas estándar
'-----------------------------------------------------
Public Function ObtenerRutaCompletaAddin(nombreFichero As String) As String
    Dim fso As Object
    Dim ruta As String

    Set fso = CreateObject("Scripting.FileSystemObject")

    ' 1. Carpeta de AddIns del usuario
    ruta = Environ$("APPDATA") & "\Microsoft\AddIns\"
    If fso.FileExists(ruta & nombreFichero) Then
        ObtenerRutaCompletaAddin = ruta & nombreFichero
        Exit Function
    End If

    ' 2. Carpeta de AddIns del sistema (detectada dinámicamente)
    ruta = RutaAddinsOffice()
    If ruta <> "" Then
        If fso.FileExists(ruta & nombreFichero) Then
            ObtenerRutaCompletaAddin = ruta & nombreFichero
            Exit Function
        End If
    End If

    ' 3. Carpeta del proyecto actual
    ruta = CurrentProject.Path & "\"
    If fso.FileExists(ruta & nombreFichero) Then
        ObtenerRutaCompletaAddin = ruta & nombreFichero
        Exit Function
    End If
End Function

'-----------------------------------------------------
' Devuelve la carpeta ADDINS de Office detectada dinámicamente
'-----------------------------------------------------
Public Function RutaAddinsOffice() As String
    Dim base As String

    base = DetectarRutaOffice()
    If base = "" Then Exit Function

    RutaAddinsOffice = base & "ADDINS\"
End Function

'-----------------------------------------------------
' Detecta la ruta de instalación de Office leyendo el registro
'-----------------------------------------------------
Public Function DetectarRutaOffice() As String
    Dim versiones As Collection
    Dim ver As Variant
    Dim ruta As String

    Set versiones = DetectarVersionesOffice()

    For Each ver In versiones
        ruta = LeerClaveRegistro("HKEY_LOCAL_MACHINE\Software\Microsoft\Office\" & ver & "\Common\InstallRoot", "Path")
        If ruta <> "" Then
            DetectarRutaOffice = ruta
            Exit Function
        End If

        ' Intento adicional para instalaciones 32/64 cruzadas
        ruta = LeerClaveRegistro("HKEY_LOCAL_MACHINE\Software\WOW6432Node\Microsoft\Office\" & ver & "\Common\InstallRoot", "Path")
        If ruta <> "" Then
            DetectarRutaOffice = ruta
            Exit Function
        End If
    Next ver
End Function

'-----------------------------------------------------
' Lee las versiones de Office instaladas desde el registro
'-----------------------------------------------------
Public Function DetectarVersionesOffice() As Collection
    Dim wsh As Object
    Dim versiones As New Collection
    Dim rutaBase As String
    Dim tempFile As String
    Dim fso As Object, ts As Object
    Dim linea As String
    Dim partes As Variant

    rutaBase = "HKEY_LOCAL_MACHINE\Software\Microsoft\Office\"
    tempFile = Environ$("TEMP") & "\office_keys.txt"

    ' Exportar subclaves de Office
    Shell "cmd /c REG QUERY """ & rutaBase & """ > """ & tempFile & """", vbHide, True

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(tempFile) Then Exit Function

    Set ts = fso.OpenTextFile(tempFile, 1)

    While Not ts.AtEndOfStream
        linea = Trim(ts.ReadLine)

        ' Coincide con ...\Office\16.0
        If linea Like "*\Office\#.#" Then
            partes = Split(linea, "\")
            versiones.Add partes(UBound(partes))
        End If
    Wend

    ts.Close
    fso.DeleteFile tempFile

    Set DetectarVersionesOffice = versiones
End Function

'-----------------------------------------------------
' Lee un valor del registro
'-----------------------------------------------------
Private Function LeerClaveRegistro(rutaClave As String, valor As String) As String
    Dim wsh As Object
    On Error Resume Next

    Set wsh = CreateObject("WScript.Shell")
    LeerClaveRegistro = wsh.RegRead(rutaClave & "\" & valor)
End Function

