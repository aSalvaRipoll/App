Attribute VB_Name = "40_mod_ControlRutasExportacion"

Option Compare Database
Option Explicit

'Módulo: 40_mod_ControlRutasExportacion

'---------------------------------------------------------------
' FUNCIÓN PRINCIPAL
'   Interpreta la ruta del usuario, crea carpetas si procede,
'   normaliza extensión, pregunta por sobrescritura y devuelve
'   una ruta final válida para exportar.
'---------------------------------------------------------------
Public Function ResolverRutaExportacion( _
        ByVal rutaUsuario As String, _
        ByVal extensionRequerida As String, _
        ByRef rutaFinal As String _
    ) As Boolean

    Dim fso As Object
    Dim carpeta As String
    Dim nombre As String
    Dim ext As String
    Dim rutaCompleta As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    ResolverRutaExportacion = False
    rutaFinal = ""

    rutaUsuario = Trim$(rutaUsuario)
    If rutaUsuario = "" Then Exit Function

    '-----------------------------------------------------------
    ' 1. ¿Es archivo existente?
    '-----------------------------------------------------------
    If fso.FileExists(rutaUsuario) Then
        ExtraerCarpeta rutaUsuario, carpeta
        ExtraerNombreArchivo rutaUsuario, nombre, ext

    '-----------------------------------------------------------
    ' 2. ¿Es carpeta existente?
    '-----------------------------------------------------------
    ElseIf fso.FolderExists(rutaUsuario) Then
        carpeta = rutaUsuario
        nombre = "InformeInspector"
        ext = extensionRequerida

    '-----------------------------------------------------------
    ' 3. No existe ? preguntar si crear ruta completa
    '-----------------------------------------------------------
    Else
        ExtraerCarpeta rutaUsuario, carpeta
        ExtraerNombreArchivo rutaUsuario, nombre, ext

        If carpeta <> "" Then
            If Not CrearCarpetaSiNoExiste(carpeta) Then Exit Function
        End If

        If nombre = "" Then nombre = "InformeInspector"
        ext = extensionRequerida
    End If

    '-----------------------------------------------------------
    ' 4. Normalizar extensión
    '-----------------------------------------------------------
    If Not NormalizarExtension(nombre, ext, extensionRequerida, nombre) Then Exit Function

    '-----------------------------------------------------------
    ' 5. Construir ruta final
    '-----------------------------------------------------------
    rutaCompleta = fso.BuildPath(carpeta, nombre & "." & extensionRequerida)

    '-----------------------------------------------------------
    ' 6. Confirmar sobrescritura si el archivo existe
    '-----------------------------------------------------------
    If Not ConfirmarSobrescritura(rutaCompleta) Then Exit Function

    rutaFinal = rutaCompleta
    ResolverRutaExportacion = True
End Function

Private Sub ExtraerCarpeta(ByVal ruta As String, ByRef carpeta As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    On Error Resume Next
    carpeta = fso.GetParentFolderName(ruta)
    On Error GoTo 0
End Sub

Private Sub ExtraerNombreArchivo(ByVal ruta As String, _
                                 ByRef nombre As String, _
                                 ByRef ext As String)

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    nombre = fso.GetBaseName(ruta)
    ext = fso.GetExtensionName(ruta)
End Sub

Private Function CrearCarpetaSiNoExiste(ByVal carpeta As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    CrearCarpetaSiNoExiste = False

    If fso.FolderExists(carpeta) Then
        CrearCarpetaSiNoExiste = True
        Exit Function
    End If

    If MsgBox("La ruta indicada no existe:" & vbCrLf & carpeta & vbCrLf & _
              "¿Deseas crearla?", vbYesNo + vbQuestion, "Crear carpeta") = vbNo Then
        Exit Function
    End If

    On Error GoTo ErrCrear
    fso.CreateFolder carpeta
    CrearCarpetaSiNoExiste = True
    Exit Function

ErrCrear:
    MsgBox "No se pudo crear la carpeta:" & vbCrLf & carpeta, vbCritical
End Function

Private Function NormalizarExtension( _
        ByVal nombre As String, _
        ByVal extActual As String, _
        ByVal extRequerida As String, _
        ByRef nombreFinal As String _
    ) As Boolean

    NormalizarExtension = False

    If LCase$(extActual) = LCase$(extRequerida) Then
        nombreFinal = nombre
        NormalizarExtension = True
        Exit Function
    End If

    If MsgBox("La extensión del archivo no coincide con el formato elegido." & vbCrLf & _
              "¿Deseas cambiarla a '." & extRequerida & "'?", _
              vbYesNo + vbQuestion, "Extensión incorrecta") = vbNo Then
        Exit Function
    End If

    nombreFinal = nombre
    NormalizarExtension = True
End Function

Private Function ConfirmarSobrescritura(ByVal rutaCompleta As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ConfirmarSobrescritura = False

    If Not fso.FileExists(rutaCompleta) Then
        ConfirmarSobrescritura = True
        Exit Function
    End If

    If MsgBox("El archivo ya existe:" & vbCrLf & rutaCompleta & vbCrLf & _
              "¿Deseas sobrescribirlo?", vbYesNo + vbQuestion, "Sobrescribir archivo") = vbYes Then
        ConfirmarSobrescritura = True
    End If
End Function


'---------------------------------------------------------------
' PrepararRutaInicial
'   Interpreta la ruta del usuario SIN crear carpetas,
'   SIN preguntar nada y SIN validar sobrescritura.
'   Solo construye una ruta inicial razonable para el diálogo.
'---------------------------------------------------------------
Public Function PrepararRutaInicial( _
        ByVal rutaUsuario As String, _
        ByVal extRequerida As String, _
        ByRef rutaInicial As String _
    ) As Boolean

    Dim fso As Object
    Dim carpeta As String
    Dim nombre As String
    Dim ext As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    PrepararRutaInicial = False
    rutaInicial = ""

    rutaUsuario = Trim$(rutaUsuario)
    If rutaUsuario = "" Then Exit Function

    ' Extraer partes
    carpeta = fso.GetParentFolderName(rutaUsuario)
    nombre = fso.GetBaseName(rutaUsuario)
    ext = fso.GetExtensionName(rutaUsuario)

    ' Si no hay carpeta, usar la actual
    If carpeta = "" Then carpeta = CurDir$

    ' Si no hay nombre, usar uno por defecto
    If nombre = "" Then nombre = "InformeInspector"

    ' Si no hay extensión, usar la requerida
    If ext = "" Then ext = extRequerida

    rutaInicial = fso.BuildPath(carpeta, nombre & "." & ext)
    PrepararRutaInicial = True
End Function


