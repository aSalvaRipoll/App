Attribute VB_Name = "51_modLogs"

Option Compare Database
Option Explicit

'===============================================================
' Módulo: 51_modLogs
' Sistema de logging del Inspector VBA
'
' - Un log diario por proyecto analizado
' - Un log independiente para reparaciones
' - Nombres normalizados: MiProyecto.accdb ? MiProyecto_accdb
' - Formato:
'       YYYYMMDD-NombreProyecto_accdb.txt
'       YYYYMMDD-Rep-NombreProyecto_accdb.txt
'===============================================================


'---------------------------------------------------------------
' Obtener nombre del proyecto con extensión normalizada
'   Ej: MiProyecto.accdb ? MiProyecto_accdb
'---------------------------------------------------------------
Public Function NombreProyectoNormalizado() As String
    Dim nombre As String
    nombre = CurrentProject.Name
    nombre = Replace(nombre, ".", "_")
    NombreProyectoNormalizado = nombre
End Function


'---------------------------------------------------------------
' Ruta del log diario de operaciones
'   Ej: 20251214-MiProyecto_accdb.txt
'---------------------------------------------------------------
Public Function RutaLogDiario() As String
    Dim carpeta As String
    Dim archivo As String
    Dim fecha As String

    carpeta = CurrentProject.Path & "\Logs"
    fecha = Format$(Date, "yyyymmdd")

    archivo = fecha & "-" & NombreProyectoNormalizado() & ".txt"

    RutaLogDiario = carpeta & "\" & archivo
End Function


'---------------------------------------------------------------
' Ruta del log diario de reparaciones
'   Ej: 20251214-Rep-MiProyecto_accdb.txt
'---------------------------------------------------------------
Public Function RutaLogReparaciones() As String
    Dim carpeta As String
    Dim archivo As String
    Dim fecha As String

    carpeta = CurrentProject.Path & "\Logs"
    fecha = Format$(Date, "yyyymmdd")

    archivo = fecha & "-Rep-" & NombreProyectoNormalizado() & ".txt"

    RutaLogReparaciones = carpeta & "\" & archivo
End Function


'---------------------------------------------------------------
' Rutina genérica para añadir una línea a un archivo de texto
'   - Crea la carpeta Logs si no existe
'   - Crea el archivo si no existe
'   - Añade al final (append)
'---------------------------------------------------------------
Public Sub Inspector_AppendToFile(rutaArchivo As String, mensaje As String)

    Dim fso As Object
    Dim carpeta As String
    Dim linea As String
    Dim ts As Object

    carpeta = CurrentProject.Path & "\Logs"

    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Crear carpeta Logs si no existe
    If Not fso.FolderExists(carpeta) Then
        fso.CreateFolder carpeta
    End If

    ' Preparar línea con timestamp
    linea = Format$(Now, "yyyy-mm-dd hh:nn:ss") & " - " & mensaje

    ' Abrir archivo en modo Append (8), crear si no existe
    Set ts = fso.OpenTextFile(rutaArchivo, 8, True)
    ts.WriteLine linea
    ts.Close
End Sub


'---------------------------------------------------------------
' Log general de operaciones del Inspector
'   - Análisis, exportaciones, etc.
'---------------------------------------------------------------
Public Sub Inspector_Log(mensaje As String)

    Dim texto As String
    texto = Format$(Now, "yyyy-mm-dd hh:nn:ss") & " - " & mensaje

    Debug.Print texto
    Inspector_AppendToFile RutaLogDiario(), texto

End Sub

'---------------------------------------------------------------
' Log específico de reparaciones
'   - Cambios aplicados al proyecto
'---------------------------------------------------------------
Public Sub Inspector_LogReparacion(mensaje As String)

    Dim texto As String
    texto = "REPARACIÓN: " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " - " & mensaje

    Debug.Print texto

    'Inspector_AppendToFile RutaLogReparaciones(), mensaje
    Inspector_AppendToFile RutaLogReparaciones(), texto
End Sub

'---------------------------------------------------------------
' Abrir la carpeta de logs en el explorador
'---------------------------------------------------------------
Public Sub Inspector_AbrirCarpetaLogs()
    Dim carpeta As String
    carpeta = CurrentProject.Path & "\Logs"

    If Dir(carpeta, vbDirectory) = "" Then
        MkDir carpeta
    End If

    Shell "explorer.exe """ & carpeta & """", vbNormalFocus
End Sub

