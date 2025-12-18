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

' =============================================================================
' Clase: clsGestorInterpretaciones
' Descripción: Gestiona la lectura y procesamiento de archivos Markdown
'              con las interpretaciones numerológicas
' Versión: 2.0 - Añadido soporte para Día de Nacimiento
' =============================================================================

' NOTA: La enumeración TipoInterpretacion está declarada en modConstantesNumerologia

' Variables privadas
Private m_RutaBase As String
Private m_UltimoError As String

' =============================================================================
' EVENTOS
' =============================================================================

Private Sub Class_Initialize()
    ' Obtener la ruta base (carpeta donde está la base de datos)
    m_RutaBase = ObtenerRutaBase()
    m_UltimoError = ""
End Sub

' =============================================================================
' PROPIEDADES
' =============================================================================

Public Property Get RutaBase() As String
    RutaBase = m_RutaBase
End Property

Public Property Let RutaBase(ByVal valor As String)
    ' Permitir establecer una ruta personalizada
    If Len(valor) > 0 Then
        m_RutaBase = valor
    End If
End Property

Public Property Get UltimoError() As String
    UltimoError = m_UltimoError
End Property

' =============================================================================
' MÉTODOS PÚBLICOS
' =============================================================================

Public Function ObtenerInterpretacion(ByVal tipo As TipoInterpretacion, _
                                     ByVal numero As Integer, _
                                     Optional ByVal numeroSecundario As Integer = 0) As String
    ' Obtiene la interpretación de un número según el tipo
    ' Para sinastría, numeroSecundario es el segundo número a comparar
    
    On Error GoTo ErrorHandler
    
    Dim rutaArchivo As String
    Dim contenido As String
    
    ' Validar número según el tipo
    If Not ValidarNumero(tipo, numero) Then
        m_UltimoError = "Número inválido para " & ObtenerNombreTipo(tipo) & ": " & numero
        ObtenerInterpretacion = ""
        Exit Function
    End If
    
    ' Construir ruta del archivo
    rutaArchivo = ConstruirRutaArchivo(tipo, numero, numeroSecundario)
    
    ' Verificar si existe el archivo
    If Not ArchivoExiste(rutaArchivo) Then
        m_UltimoError = "No se encontró el archivo: " & rutaArchivo
        ObtenerInterpretacion = MensajeInterpretacionNoDisponible(tipo, numero)
        Exit Function
    End If
    
    ' Leer contenido del archivo
    contenido = LeerArchivoTexto(rutaArchivo)
    
    If Len(contenido) = 0 Then
        m_UltimoError = "El archivo está vacío: " & rutaArchivo
        ObtenerInterpretacion = MensajeInterpretacionNoDisponible(tipo, numero)
    Else
        ObtenerInterpretacion = contenido
        m_UltimoError = ""
    End If
    
    Exit Function
    
ErrorHandler:
    m_UltimoError = "Error al obtener interpretación: " & Err.Description
    ObtenerInterpretacion = ""
End Function

Public Function ObtenerInterpretacionFormateada(ByVal tipo As TipoInterpretacion, _
                                               ByVal numero As Integer, _
                                               Optional ByVal numeroSecundario As Integer = 0) As String
    ' Obtiene la interpretación y la formatea para mostrar en Access
    ' Convierte Markdown básico a texto formateado
    
    Dim contenidoMD As String
    Dim contenidoFormateado As String
    
    contenidoMD = ObtenerInterpretacion(tipo, numero, numeroSecundario)
    
    If Len(contenidoMD) = 0 Then
        ObtenerInterpretacionFormateada = ""
        Exit Function
    End If
    
    ' Convertir Markdown a formato legible
    contenidoFormateado = ConvertirMarkdownATexto(contenidoMD)
    
    ObtenerInterpretacionFormateada = contenidoFormateado
End Function

Public Function CrearEstructuraCarpetas() As Boolean
    ' Crea la estructura de carpetas para las interpretaciones
    ' Retorna True si se creó exitosamente o ya existe
    
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Dim carpetaPrincipal As String
    Dim carpetas() As String
    Dim i As Integer
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Carpeta principal
    carpetaPrincipal = m_RutaBase & "Interpretaciones"
    
    ' Lista de subcarpetas (ACTUALIZADO: añadido DiaNacimiento)
    carpetas = Split("CaminoVida,Destino,Alma,Personalidad,Madurez,Sinastria,DiaNacimiento", ",")
    
    ' Crear carpeta principal si no existe
    If Not fso.FolderExists(carpetaPrincipal) Then
        fso.CreateFolder carpetaPrincipal
        Debug.Print "Carpeta creada: " & carpetaPrincipal
    End If
    
    ' Crear subcarpetas
    For i = LBound(carpetas) To UBound(carpetas)
        Dim rutaSubcarpeta As String
        rutaSubcarpeta = carpetaPrincipal & "\" & carpetas(i)
        
        If Not fso.FolderExists(rutaSubcarpeta) Then
            fso.CreateFolder rutaSubcarpeta
            Debug.Print "Carpeta creada: " & rutaSubcarpeta
        End If
    Next i
    
    Set fso = Nothing
    CrearEstructuraCarpetas = True
    
    MsgBox "Estructura de carpetas creada exitosamente en:" & vbCrLf & vbCrLf & _
           carpetaPrincipal, vbInformation, "Carpetas Creadas"
    
    Exit Function
    
ErrorHandler:
    m_UltimoError = "Error al crear estructura: " & Err.Description
    CrearEstructuraCarpetas = False
    Set fso = Nothing
    MsgBox "Error al crear carpetas: " & Err.Description, vbCritical
End Function

Public Function CrearPlantillasBasicas() As Boolean
    ' Crea archivos de plantilla básicos para cada número
    
    On Error GoTo ErrorHandler
    
    Dim tipos() As TipoInterpretacion
    Dim numeros() As Integer
    Dim i As Integer, j As Integer
    Dim rutaArchivo As String
    Dim contenidoPlantilla As String
    
    ' Primero crear estructura de carpetas
    If Not CrearEstructuraCarpetas() Then
        CrearPlantillasBasicas = False
        Exit Function
    End If
    
    ' Tipos de interpretación (sin sinastría por ahora)
    tipos = Array(tiCaminoVida, tiDestino, tiAlma, tiPersonalidad, tiMadurez)
    
    ' Números a crear (1-9 y maestros)
    numeros = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 11, 22, 33, 44)
    
    ' Crear plantillas para cada combinación
    For i = LBound(tipos) To UBound(tipos)
        For j = LBound(numeros) To UBound(numeros)
            rutaArchivo = ConstruirRutaArchivo(tipos(i), numeros(j), 0)
            
            ' Solo crear si no existe
            If Not ArchivoExiste(rutaArchivo) Then
                contenidoPlantilla = GenerarPlantilla(tipos(i), numeros(j))
                
                If GuardarArchivoTexto(rutaArchivo, contenidoPlantilla) Then
                    Debug.Print "Plantilla creada: " & rutaArchivo
                End If
            End If
        Next j
    Next i
    
    CrearPlantillasBasicas = True
    
    MsgBox "Plantillas básicas creadas exitosamente." & vbCrLf & vbCrLf & _
           "Puedes editarlas en la carpeta:" & vbCrLf & _
           m_RutaBase & "Interpretaciones\", _
           vbInformation, "Plantillas Creadas"
    
    Exit Function
    
ErrorHandler:
    m_UltimoError = "Error al crear plantillas: " & Err.Description
    CrearPlantillasBasicas = False
    MsgBox "Error: " & Err.Description, vbCritical
End Function

' =============================================================================
' MÉTODOS PRIVADOS
' =============================================================================

Private Function ObtenerRutaBase() As String
    ' Obtiene la ruta donde está la base de datos
    Dim rutaCompleta As String
    Dim posicion As Integer
    
    rutaCompleta = CurrentDb.Name
    posicion = InStrRev(rutaCompleta, "\")
    
    If posicion > 0 Then
        ObtenerRutaBase = Left(rutaCompleta, posicion)
    Else
        ObtenerRutaBase = CurrentProject.Path & "\"
    End If
End Function

Private Function ConstruirRutaArchivo(ByVal tipo As TipoInterpretacion, _
                                     ByVal numero As Integer, _
                                     ByVal numeroSecundario As Integer) As String
    ' Construye la ruta completa del archivo de interpretación
    
    Dim carpeta As String
    Dim nombreArchivo As String
    
    ' Determinar carpeta según tipo
    Select Case tipo
        Case tiCaminoVida
            carpeta = "CaminoVida"
        Case tiDestino
            carpeta = "Destino"
        Case tiAlma
            carpeta = "Alma"
        Case tiPersonalidad
            carpeta = "Personalidad"
        Case tiMadurez
            carpeta = "Madurez"
        Case tiSinastria
            carpeta = "Sinastria"
        Case tiDiaNacimiento  ' NUEVO
            carpeta = "DiaNacimiento"
    End Select
    
    ' Construir nombre de archivo
    If tipo = tiSinastria Then
        ' Para sinastría: 01_con_05.md
        nombreArchivo = Format(numero, "00") & "_con_" & Format(numeroSecundario, "00") & ".md"
    ElseIf tipo = tiDiaNacimiento Then
        ' NUEVO: Para día de nacimiento: 01_DiaNacimiento.md
        nombreArchivo = Format(numero, "00") & "_DiaNacimiento.md"
    Else
        ' Para otros: 01_CaminoVida.md
        nombreArchivo = Format(numero, "00") & "_" & carpeta & ".md"
    End If
    
    ' Ruta completa
    ConstruirRutaArchivo = m_RutaBase & "Interpretaciones\" & carpeta & "\" & nombreArchivo
End Function

Private Function ValidarNumero(ByVal tipo As TipoInterpretacion, ByVal numero As Integer) As Boolean
    ' Valida que el número sea válido según el tipo de interpretación
    
    Select Case tipo
        Case tiDiaNacimiento
            ' Para día de nacimiento: 1-31
            ValidarNumero = (numero >= 1 And numero <= 31)
        Case Else
            ' Para otros tipos: 1-9 y números maestros
            ValidarNumero = (numero >= 1 And numero <= 9) Or _
                            numero = 11 Or numero = 22 Or numero = 33 Or numero = 44
    End Select
End Function

Private Function ArchivoExiste(ByVal ruta As String) As Boolean
    ' Verifica si existe un archivo
    On Error Resume Next
    ArchivoExiste = (Dir(ruta) <> "")
End Function

Private Function LeerArchivoTexto(ByVal ruta As String) As String
    ' Lee el contenido de un archivo de texto (UTF-8)
    
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Dim archivo As Object
    Dim contenido As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set archivo = fso.OpenTextFile(ruta, 1, False, -1) ' -1 = Unicode/UTF-8
    
    contenido = archivo.ReadAll
    archivo.Close
    
    Set archivo = Nothing
    Set fso = Nothing
    
    LeerArchivoTexto = contenido
    Exit Function
    
ErrorHandler:
    LeerArchivoTexto = ""
    If Not archivo Is Nothing Then archivo.Close
    Set archivo = Nothing
    Set fso = Nothing
End Function

Private Function GuardarArchivoTexto(ByVal ruta As String, ByVal contenido As String) As Boolean
    ' Guarda contenido en un archivo de texto (UTF-8)
    
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Dim archivo As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set archivo = fso.CreateTextFile(ruta, True, True) ' True = Unicode
    
    archivo.Write contenido
    archivo.Close
    
    Set archivo = Nothing
    Set fso = Nothing
    
    GuardarArchivoTexto = True
    Exit Function
    
ErrorHandler:
    GuardarArchivoTexto = False
    If Not archivo Is Nothing Then archivo.Close
    Set archivo = Nothing
    Set fso = Nothing
End Function

Private Function GenerarPlantilla(ByVal tipo As TipoInterpretacion, _
                                 ByVal numero As Integer) As String
    ' Genera una plantilla básica para un archivo de interpretación
    
    Dim plantilla As String
    Dim nombreTipo As String
    Dim nombreNumero As String
    
    ' Nombre del tipo
    Select Case tipo
        Case tiCaminoVida: nombreTipo = "Camino de Vida"
        Case tiDestino: nombreTipo = "Número del Destino"
        Case tiAlma: nombreTipo = "Número del Alma"
        Case tiPersonalidad: nombreTipo = "Número de Personalidad"
        Case tiMadurez: nombreTipo = "Número de Madurez"
        Case tiDiaNacimiento: nombreTipo = "Día de Nacimiento"  ' NUEVO
    End Select
    
    ' Nombre del número
    If numero = 11 Then
        nombreNumero = "11 - Maestro Iluminador"
    ElseIf numero = 22 Then
        nombreNumero = "22 - Maestro Constructor"
    ElseIf numero = 33 Then
        nombreNumero = "33 - Maestro Sanador"
    ElseIf numero = 44 Then
        nombreNumero = "44 - Maestro Visionario"
    Else
        nombreNumero = CStr(numero)
    End If
    
    ' Construir plantilla
    plantilla = "# " & nombreTipo & " - Número " & nombreNumero & vbCrLf
    plantilla = plantilla & vbCrLf
    plantilla = plantilla & "## Significado General" & vbCrLf
    plantilla = plantilla & vbCrLf
    plantilla = plantilla & "[Escribe aquí el significado general del número " & numero & " para " & nombreTipo & "]" & vbCrLf
    plantilla = plantilla & vbCrLf
    plantilla = plantilla & "## Características Principales" & vbCrLf
    plantilla = plantilla & vbCrLf
    plantilla = plantilla & "- Característica 1" & vbCrLf
    plantilla = plantilla & "- Característica 2" & vbCrLf
    plantilla = plantilla & "- Característica 3" & vbCrLf
    plantilla = plantilla & vbCrLf
    plantilla = plantilla & "## Fortalezas" & vbCrLf
    plantilla = plantilla & vbCrLf
    plantilla = plantilla & "[Describe las fortalezas]" & vbCrLf
    plantilla = plantilla & vbCrLf
    plantilla = plantilla & "## Desafíos" & vbCrLf
    plantilla = plantilla & vbCrLf
    plantilla = plantilla & "[Describe los desafíos]" & vbCrLf
    plantilla = plantilla & vbCrLf
    plantilla = plantilla & "## Consejo" & vbCrLf
    plantilla = plantilla & vbCrLf
    plantilla = plantilla & "[Consejo o reflexión final]" & vbCrLf
    
    GenerarPlantilla = plantilla
End Function

Private Function ConvertirMarkdownATexto(ByVal markdown As String) As String
    ' Convierte Markdown básico a texto formateado para mostrar en Access
    ' Conversión simple: elimina símbolos de formato pero mantiene estructura
    
    Dim resultado As String
    Dim lineas() As String
    Dim i As Integer
    Dim linea As String
    
    resultado = markdown
    
    ' Reemplazos básicos de Markdown
    resultado = Replace(resultado, "# ", vbCrLf & "═══ ")
    resultado = Replace(resultado, "## ", vbCrLf & "─── ")
    resultado = Replace(resultado, "### ", vbCrLf & "• ")
    resultado = Replace(resultado, "**", "")
    resultado = Replace(resultado, "__", "")
    resultado = Replace(resultado, "*", "")
    resultado = Replace(resultado, "_", "")
    
    ConvertirMarkdownATexto = Trim(resultado)
End Function

Private Function MensajeInterpretacionNoDisponible(ByVal tipo As TipoInterpretacion, _
                                                   ByVal numero As Integer) As String
    ' Mensaje cuando no hay interpretación disponible
    
    Dim mensaje As String
    
    mensaje = "═══ INTERPRETACIÓN NO DISPONIBLE ═══" & vbCrLf & vbCrLf
    mensaje = mensaje & "La interpretación para este número aún no ha sido creada." & vbCrLf & vbCrLf
    mensaje = mensaje & "Número: " & numero & vbCrLf
    mensaje = mensaje & "Tipo: " & ObtenerNombreTipo(tipo) & vbCrLf & vbCrLf
    mensaje = mensaje & "Puedes crear las plantillas ejecutando:" & vbCrLf
    mensaje = mensaje & "  CrearPlantillasBasicas()" & vbCrLf
    
    MensajeInterpretacionNoDisponible = mensaje
End Function

Private Function ObtenerNombreTipo(ByVal tipo As TipoInterpretacion) As String
    Select Case tipo
        Case tiCaminoVida: ObtenerNombreTipo = "Camino de Vida"
        Case tiDestino: ObtenerNombreTipo = "Destino"
        Case tiAlma: ObtenerNombreTipo = "Alma"
        Case tiPersonalidad: ObtenerNombreTipo = "Personalidad"
        Case tiMadurez: ObtenerNombreTipo = "Madurez"
        Case tiSinastria: ObtenerNombreTipo = "Sinastría"
        Case tiDiaNacimiento: ObtenerNombreTipo = "Día de Nacimiento"  ' NUEVO
        Case Else: ObtenerNombreTipo = "Desconocido"
    End Select
End Function
