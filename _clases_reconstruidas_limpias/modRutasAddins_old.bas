Attribute VB_Name = "modRutasAddins_old"

Option Compare Database
Option Explicit


Public Function RutaAddinsOffice() As String
    Dim base As String

    base = DetectarRutaOffice()
    If base = "" Then Exit Function

    RutaAddinsOffice = base & "ADDINS\"
End Function


Public Function DetectarVersionesOffice() As Collection
    Dim wsh As Object
    Dim versiones As New Collection
    Dim rutaBase As String
    Dim subclaves As Variant
    Dim clave As Variant
    Dim i As Long

    rutaBase = "HKEY_LOCAL_MACHINE\Software\Microsoft\Office\"
    Set wsh = CreateObject("WScript.Shell")

    ' Exportamos la lista de subclaves de Office
    Dim tempFile As String
    tempFile = Environ$("TEMP") & "\office_keys.txt"

    Shell "cmd /c REG QUERY """ & rutaBase & """ > """ & tempFile & """", vbHide, True

    ' Leemos el archivo resultante
    Dim fso As Object, ts As Object, linea As String
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(tempFile) Then Exit Function

    Set ts = fso.OpenTextFile(tempFile, 1)

    While Not ts.AtEndOfStream
        linea = Trim(ts.ReadLine)

        ' Las versiones son subclaves numéricas: 14.0, 15.0, 16.0, 17.0...
        If linea Like "*\Office\#.#" Then
            clave = Split(linea, "\")
            versiones.Add clave(UBound(clave))
        End If
    Wend

    ts.Close
    fso.DeleteFile tempFile

    Set DetectarVersionesOffice = versiones
End Function

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
    Next ver
End Function

Private Function LeerClaveRegistro(rutaClave As String, valor As String) As String
    Dim wsh As Object
    On Error Resume Next

    Set wsh = CreateObject("WScript.Shell")
    LeerClaveRegistro = wsh.RegRead(rutaClave & "\" & valor)
End Function





'--------------------------------------------------------------------------------
Public Function ObtenerRutaCompletaAddin(nombreFichero As String) As String
    Dim rutas As Collection
    Dim ruta As Variant
    Dim fso As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set rutas = RutasEstandarAddins()

    ' Buscar en todas las rutas estándar
    For Each ruta In rutas
        If fso.FileExists(ruta & nombreFichero) Then
            ObtenerRutaCompletaAddin = ruta & nombreFichero
            Exit Function
        End If
    Next ruta
End Function

Private Function RutasEstandarAddins() As Collection
    Dim col As New Collection
    Dim base As String

    ' Carpeta de AddIns del usuario
    col.Add Environ$("APPDATA") & "\Microsoft\AddIns\"

    ' Carpeta de AddIns del sistema (Office 32/64)
    col.Add Environ$("ProgramFiles") & "\Microsoft Office\root\Office16\ADDINS\"
    col.Add Environ$("ProgramFiles(x86)") & "\Microsoft Office\root\Office16\ADDINS\"

    ' Carpeta de instalación de Access (por si acaso)
    col.Add Environ$("ProgramFiles") & "\Microsoft Office\root\Office16\"
    col.Add Environ$("ProgramFiles(x86)") & "\Microsoft Office\root\Office16\"

    ' Carpeta del proyecto actual
    col.Add CurrentProject.Path & "\"

    Set RutasEstandarAddins = col
End Function
