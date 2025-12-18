VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsGestorInterpretaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


' ============================================================================
' Clase: clsGestorInterpretaciones
' Descripción: Gestor centralizado para cargar y gestionar interpretaciones
' Autor: Sistema de Numerología
' Fecha: 2024
' ============================================================================

Private pRutaBase As String
Private pCacheInterpretaciones As Object  ' Dictionary para cache

' ============================================================================
' PROPIEDADES PÚBLICAS
' ============================================================================

Public Property Get RutaBase() As String
    RutaBase = pRutaBase
End Property

Public Property Let RutaBase(ByVal valor As String)
    pRutaBase = valor
End Property

' ============================================================================
' MÉTODOS PÚBLICOS
' ============================================================================

Public Function CargarInterpretacion(ByVal tipoCalculo As String, ByVal numero As Integer) As String
    Dim nombreArchivo As String
    Dim rutaCompleta As String
    Dim clave As String
    
    On Error GoTo ErrorHandler
    
    ' Construir clave de cache
    clave = tipoCalculo & "_" & CStr(numero)
    
    ' Verificar si está en cache
    If pCacheInterpretaciones.Exists(clave) Then
        CargarInterpretacion = pCacheInterpretaciones(clave)
        Exit Function
    End If
    
    ' Construir nombre y ruta del archivo
    nombreArchivo = tipoCalculo & "_" & CStr(numero) & EXTENSION_MARKDOWN
    rutaCompleta = pRutaBase & nombreArchivo
    
    ' Cargar desde archivo
    Dim contenido As String
    contenido = LeerArchivo(rutaCompleta)
    
    ' Guardar en cache
    pCacheInterpretaciones.Add clave, contenido
    
    CargarInterpretacion = contenido
    Exit Function
    
ErrorHandler:
    err.Raise err.Number, "clsGestorInterpretaciones.CargarInterpretacion", err.Description
End Function

Public Sub LimpiarCache()
    pCacheInterpretaciones.RemoveAll
End Sub

Public Function ObtenerEstadisticasCache() As String
    Dim info As String
    
    info = "Interpretaciones en cache: " & pCacheInterpretaciones.Count & vbCrLf
    info = info & "Ruta base: " & pRutaBase
    
    ObtenerEstadisticasCache = info
End Function

' ============================================================================
' MÉTODOS PRIVADOS
' ============================================================================

Private Function LeerArchivo(ByVal rutaCompleta As String) As String
    Dim numArchivo As Integer
    Dim contenido As String
    
    On Error GoTo ErrorHandler
    
    ' Verificar que existe el archivo
    If Dir(rutaCompleta) = "" Then
        err.Raise vbObjectError + 3001, "clsGestorInterpretaciones", ERR_ARCHIVO_NO_ENCONTRADO & ": " & rutaCompleta
    End If
    
    ' Leer archivo
    numArchivo = FreeFile
    Open rutaCompleta For Input As #numArchivo
    contenido = Input$(LOF(numArchivo), numArchivo)
    Close #numArchivo
    
    LeerArchivo = contenido
    Exit Function
    
ErrorHandler:
    If numArchivo > 0 Then Close #numArchivo
    err.Raise err.Number, "clsGestorInterpretaciones.LeerArchivo", err.Description
End Function

' ============================================================================
' EVENTOS DE CLASE
' ============================================================================

Private Sub Class_Initialize()
    ' Inicializar ruta base
    pRutaBase = CurrentProject.Path & "\" & RUTA_INTERPRETACIONES
    
    ' Crear diccionario para cache
    Set pCacheInterpretaciones = CreateObject("Scripting.Dictionary")
End Sub

Private Sub Class_Terminate()
    ' Limpiar cache
    Set pCacheInterpretaciones = Nothing
End Sub

