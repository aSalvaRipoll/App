VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCalculoSinastria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ============================================================================
' Clase: clsCalculoSinastria
' Descripción: Calcula la compatibilidad numerológica entre dos personas
' Autor: Sistema de Numerología
' Fecha: 2024
' ============================================================================

' Enumeración de tipos de sinastría
Public Enum TipoSinastria
    General = 1
    Romantica = 2
    Laboral = 3
End Enum

' Enumeración de aspectos a comparar
Public Enum AspectoSinastria
    CaminoVida = 1
    Destino = 2
    Alma = 3
    Personalidad = 4
    Madurez = 5
End Enum

' Propiedades privadas - Persona 1
Private pNombre1 As String
Private pFecha1 As Date
Private pCaminoVida1 As Integer
Private pDestino1 As Integer
Private pAlma1 As Integer
Private pPersonalidad1 As Integer
Private pMadurez1 As Integer

' Propiedades privadas - Persona 2
Private pNombre2 As String
Private pFecha2 As Date
Private pCaminoVida2 As Integer
Private pDestino2 As Integer
Private pAlma2 As Integer
Private pPersonalidad2 As Integer
Private pMadurez2 As Integer

' Propiedades privadas - Configuración
Private pTipoSinastria As TipoSinastria
Private pRutaBase As String
Private pCalculosRealizados As Boolean

' ============================================================================
' PROPIEDADES PÚBLICAS - PERSONA 1
' ============================================================================

Public Property Get Nombre1() As String
    Nombre1 = pNombre1
End Property

Public Property Let Nombre1(ByVal valor As String)
    pNombre1 = NormalizarTexto(valor)
    pCalculosRealizados = False
End Property

Public Property Get Fecha1() As Date
    Fecha1 = pFecha1
End Property

Public Property Let Fecha1(ByVal valor As Date)
    If Not ValidarFecha(valor) Then
        Err.Raise vbObjectError + 4001, "clsCalculoSinastria", ERR_FECHA_INVALIDA
    End If
    pFecha1 = valor
    pCalculosRealizados = False
End Property

Public Property Get CaminoVida1() As Integer
    CaminoVida1 = pCaminoVida1
End Property

Public Property Get Destino1() As Integer
    Destino1 = pDestino1
End Property

Public Property Get Alma1() As Integer
    Alma1 = pAlma1
End Property

Public Property Get Personalidad1() As Integer
    Personalidad1 = pPersonalidad1
End Property

Public Property Get Madurez1() As Integer
    Madurez1 = pMadurez1
End Property

' ============================================================================
' PROPIEDADES PÚBLICAS - PERSONA 2
' ============================================================================

Public Property Get Nombre2() As String
    Nombre2 = pNombre2
End Property

Public Property Let Nombre2(ByVal valor As String)
    pNombre2 = NormalizarTexto(valor)
    pCalculosRealizados = False
End Property

Public Property Get Fecha2() As Date
    Fecha2 = pFecha2
End Property

Public Property Let Fecha2(ByVal valor As Date)
    If Not ValidarFecha(valor) Then
        Err.Raise vbObjectError + 4002, "clsCalculoSinastria", ERR_FECHA_INVALIDA
    End If
    pFecha2 = valor
    pCalculosRealizados = False
End Property

Public Property Get CaminoVida2() As Integer
    CaminoVida2 = pCaminoVida2
End Property

Public Property Get Destino2() As Integer
    Destino2 = pDestino2
End Property

Public Property Get Alma2() As Integer
    Alma2 = pAlma2
End Property

Public Property Get Personalidad2() As Integer
    Personalidad2 = pPersonalidad2
End Property

Public Property Get Madurez2() As Integer
    Madurez2 = pMadurez2
End Property

' ============================================================================
' PROPIEDADES PÚBLICAS - CONFIGURACIÓN
' ============================================================================

Public Property Get TipoSinastriaActual() As TipoSinastria
    TipoSinastriaActual = pTipoSinastria
End Property

Public Property Let TipoSinastriaActual(ByVal valor As TipoSinastria)
    pTipoSinastria = valor
End Property

Public Property Get RutaBase() As String
    RutaBase = pRutaBase
End Property

Public Property Let RutaBase(ByVal valor As String)
    pRutaBase = valor
End Property

' ============================================================================
' MÉTODO PRINCIPAL - CALCULAR TODOS LOS NÚMEROS
' ============================================================================

Public Sub CalcularTodosLosNumeros()
    On Error GoTo ErrorHandler
    
    ' Validar datos
    If Len(pNombre1) = 0 Or Len(pNombre2) = 0 Then
        Err.Raise vbObjectError + 4003, "clsCalculoSinastria", "Faltan nombres para calcular"
    End If
    
    If pFecha1 = 0 Or pFecha2 = 0 Then
        Err.Raise vbObjectError + 4004, "clsCalculoSinastria", "Faltan fechas para calcular"
    End If
    
    ' Calcular números de Persona 1
    pCaminoVida1 = CalcularCaminoVida(pFecha1)
    pDestino1 = CalcularDestino(pNombre1)
    pAlma1 = CalcularAlma(pNombre1)
    pPersonalidad1 = CalcularPersonalidad(pNombre1)
    pMadurez1 = CalcularMadurez(pNombre1, pFecha1)
    
    ' Calcular números de Persona 2
    pCaminoVida2 = CalcularCaminoVida(pFecha2)
    pDestino2 = CalcularDestino(pNombre2)
    pAlma2 = CalcularAlma(pNombre2)
    pPersonalidad2 = CalcularPersonalidad(pNombre2)
    pMadurez2 = CalcularMadurez(pNombre2, pFecha2)
    
    pCalculosRealizados = True
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, "clsCalculoSinastria.CalcularTodosLosNumeros", Err.Description
End Sub

' ============================================================================
' MÉTODOS PARA OBTENER RUTAS DE ARCHIVOS
' ============================================================================

Public Function ObtenerRutaInterpretacion(ByVal aspecto As AspectoSinastria, _
                                          ByVal numero1 As Integer, _
                                          ByVal numero2 As Integer) As String
    Dim carpetaTipo As String
    Dim carpetaAspecto As String
    Dim nombreArchivo As String
    Dim rutaCompleta As String
    
    On Error GoTo ErrorHandler
    
    ' Determinar carpeta de tipo
    Select Case pTipoSinastria
        Case TipoSinastria.General
            carpetaTipo = "General"
        Case TipoSinastria.Romantica
            carpetaTipo = "Romantica"
        Case TipoSinastria.Laboral
            carpetaTipo = "Laboral"
    End Select
    
    ' Determinar carpeta de aspecto
    Select Case aspecto
        Case AspectoSinastria.CaminoVida
            carpetaAspecto = "CaminoVida"
        Case AspectoSinastria.Destino
            carpetaAspecto = "Destino"
        Case AspectoSinastria.Alma
            carpetaAspecto = "Alma"
        Case AspectoSinastria.Personalidad
            carpetaAspecto = "Personalidad"
        Case AspectoSinastria.Madurez
            carpetaAspecto = "Madurez"
    End Select
    
    ' Construir nombre de archivo (ejemplo: 01_05.md o 11_22.md)
    nombreArchivo = Format(numero1, "00") & "_" & Format(numero2, "00") & ".md"
    
    ' Construir ruta completa
    ' Ejemplo: Sinastrias\Romantica\CaminoVida\01_05.md
    rutaCompleta = pRutaBase & carpetaTipo & "\" & carpetaAspecto & "\" & nombreArchivo
    
    ObtenerRutaInterpretacion = rutaCompleta
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, "clsCalculoSinastria.ObtenerRutaInterpretacion", Err.Description
End Function

Public Function ObtenerTodasLasRutas() As String
    Dim resultado As String
    
    On Error GoTo ErrorHandler
    
    If Not pCalculosRealizados Then
        CalcularTodosLosNumeros
    End If
    
    resultado = "=== RUTAS DE INTERPRETACIONES ===" & vbCrLf
    resultado = resultado & "Tipo: " & ObtenerNombreTipo() & vbCrLf
    resultado = resultado & "Persona 1: " & pNombre1 & vbCrLf
    resultado = resultado & "Persona 2: " & pNombre2 & vbCrLf & vbCrLf
    
    ' Camino de Vida
    resultado = resultado & "CAMINO DE VIDA (" & pCaminoVida1 & " - " & pCaminoVida2 & "):" & vbCrLf
    resultado = resultado & ObtenerRutaInterpretacion(AspectoSinastria.CaminoVida, pCaminoVida1, pCaminoVida2) & vbCrLf & vbCrLf
    
    ' Destino
    resultado = resultado & "DESTINO (" & pDestino1 & " - " & pDestino2 & "):" & vbCrLf
    resultado = resultado & ObtenerRutaInterpretacion(AspectoSinastria.Destino, pDestino1, pDestino2) & vbCrLf & vbCrLf
    
    ' Alma
    resultado = resultado & "ALMA (" & pAlma1 & " - " & pAlma2 & "):" & vbCrLf
    resultado = resultado & ObtenerRutaInterpretacion(AspectoSinastria.Alma, pAlma1, pAlma2) & vbCrLf & vbCrLf
    
    ' Personalidad
    resultado = resultado & "PERSONALIDAD (" & pPersonalidad1 & " - " & pPersonalidad2 & "):" & vbCrLf
    resultado = resultado & ObtenerRutaInterpretacion(AspectoSinastria.Personalidad, pPersonalidad1, pPersonalidad2) & vbCrLf & vbCrLf
    
    ' Madurez
    resultado = resultado & "MADUREZ (" & pMadurez1 & " - " & pMadurez2 & "):" & vbCrLf
    resultado = resultado & ObtenerRutaInterpretacion(AspectoSinastria.Madurez, pMadurez1, pMadurez2) & vbCrLf
    
    ObtenerTodasLasRutas = resultado
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, "clsCalculoSinastria.ObtenerTodasLasRutas", Err.Description
End Function

' ============================================================================
' MÉTODOS PARA CARGAR INTERPRETACIONES
' ============================================================================

Public Function CargarInterpretacion(ByVal aspecto As AspectoSinastria) As String
    Dim rutaArchivo As String
    Dim numero1 As Integer
    Dim numero2 As Integer
    
    On Error GoTo ErrorHandler
    
    If Not pCalculosRealizados Then
        CalcularTodosLosNumeros
    End If
    
    ' Obtener números según el aspecto
    Select Case aspecto
        Case AspectoSinastria.CaminoVida
            numero1 = pCaminoVida1
            numero2 = pCaminoVida2
        Case AspectoSinastria.Destino
            numero1 = pDestino1
            numero2 = pDestino2
        Case AspectoSinastria.Alma
            numero1 = pAlma1
            numero2 = pAlma2
        Case AspectoSinastria.Personalidad
            numero1 = pPersonalidad1
            numero2 = pPersonalidad2
        Case AspectoSinastria.Madurez
            numero1 = pMadurez1
            numero2 = pMadurez2
    End Select
    
    ' Obtener ruta del archivo
    rutaArchivo = ObtenerRutaInterpretacion(aspecto, numero1, numero2)
    
    ' Cargar contenido
    CargarInterpretacion = LeerArchivoInterpretacion(rutaArchivo)
    Exit Function
    
ErrorHandler:
    CargarInterpretacion = "Error al cargar interpretación: " & Err.Description
End Function

Public Function CargarTodasLasInterpretaciones() As String
    Dim resultado As String
    
    On Error GoTo ErrorHandler
    
    If Not pCalculosRealizados Then
        CalcularTodosLosNumeros
    End If
    
    resultado = "=== SINASTRÍA COMPLETA ===" & vbCrLf
    resultado = resultado & "Tipo: " & ObtenerNombreTipo() & vbCrLf
    resultado = resultado & "Persona 1: " & pNombre1 & " (" & Format(pFecha1, "dd/mm/yyyy") & ")" & vbCrLf
    resultado = resultado & "Persona 2: " & pNombre2 & " (" & Format(pFecha2, "dd/mm/yyyy") & ")" & vbCrLf
    resultado = resultado & String(70, "=") & vbCrLf & vbCrLf
    
    ' Camino de Vida
    resultado = resultado & ">>> CAMINO DE VIDA (" & pCaminoVida1 & " - " & pCaminoVida2 & ") <<<" & vbCrLf
    resultado = resultado & CargarInterpretacion(AspectoSinastria.CaminoVida) & vbCrLf & vbCrLf
    resultado = resultado & String(70, "-") & vbCrLf & vbCrLf
    
    ' Destino
    resultado = resultado & ">>> DESTINO (" & pDestino1 & " - " & pDestino2 & ") <<<" & vbCrLf
    resultado = resultado & CargarInterpretacion(AspectoSinastria.Destino) & vbCrLf & vbCrLf
    resultado = resultado & String(70, "-") & vbCrLf & vbCrLf
    
    ' Alma
    resultado = resultado & ">>> ALMA (" & pAlma1 & " - " & pAlma2 & ") <<<" & vbCrLf
    resultado = resultado & CargarInterpretacion(AspectoSinastria.Alma) & vbCrLf & vbCrLf
    resultado = resultado & String(70, "-") & vbCrLf & vbCrLf
    
    ' Personalidad
    resultado = resultado & ">>> PERSONALIDAD (" & pPersonalidad1 & " - " & pPersonalidad2 & ") <<<" & vbCrLf
    resultado = resultado & CargarInterpretacion(AspectoSinastria.Personalidad) & vbCrLf & vbCrLf
    resultado = resultado & String(70, "-") & vbCrLf & vbCrLf
    
    ' Madurez
    resultado = resultado & ">>> MADUREZ (" & pMadurez1 & " - " & pMadurez2 & ") <<<" & vbCrLf
    resultado = resultado & CargarInterpretacion(AspectoSinastria.Madurez) & vbCrLf
    
    CargarTodasLasInterpretaciones = resultado
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, "clsCalculoSinastria.CargarTodasLasInterpretaciones", Err.Description
End Function

' ============================================================================
' MÉTODO PARA OBTENER RESUMEN DE NÚMEROS
' ============================================================================

Public Function ObtenerResumenNumeros() As String
    Dim resumen As String
    
    If Not pCalculosRealizados Then
        CalcularTodosLosNumeros
    End If
    
    resumen = "=== NÚMEROS CALCULADOS ===" & vbCrLf & vbCrLf
    
    resumen = resumen & "PERSONA 1: " & pNombre1 & vbCrLf
    resumen = resumen & "  Camino de Vida: " & FormatearNumeroResultado(pCaminoVida1, True) & vbCrLf
    resumen = resumen & "  Destino: " & FormatearNumeroResultado(pDestino1, True) & vbCrLf
    resumen = resumen & "  Alma: " & FormatearNumeroResultado(pAlma1, True) & vbCrLf
    resumen = resumen & "  Personalidad: " & FormatearNumeroResultado(pPersonalidad1, True) & vbCrLf
    resumen = resumen & "  Madurez: " & FormatearNumeroResultado(pMadurez1, True) & vbCrLf & vbCrLf
    
    resumen = resumen & "PERSONA 2: " & pNombre2 & vbCrLf
    resumen = resumen & "  Camino de Vida: " & FormatearNumeroResultado(pCaminoVida2, True) & vbCrLf
    resumen = resumen & "  Destino: " & FormatearNumeroResultado(pDestino2, True) & vbCrLf
    resumen = resumen & "  Alma: " & FormatearNumeroResultado(pAlma2, True) & vbCrLf
    resumen = resumen & "  Personalidad: " & FormatearNumeroResultado(pPersonalidad2, True) & vbCrLf
    resumen = resumen & "  Madurez: " & FormatearNumeroResultado(pMadurez2, True) & vbCrLf
    
    ObtenerResumenNumeros = resumen
End Function

' ============================================================================
' MÉTODOS PRIVADOS - CÁLCULOS
' ============================================================================

Private Function CalcularCaminoVida(ByVal fecha As Date) As Integer
    Dim obj As clsCalculoCaminoVida
    Set obj = New clsCalculoCaminoVida
    obj.FechaNacimiento = fecha
    CalcularCaminoVida = obj.Calcular()
    Set obj = Nothing
End Function

Private Function CalcularDestino(ByVal nombre As String) As Integer
    Dim obj As clsCalculoDestino
    Set obj = New clsCalculoDestino
    obj.NombreCompleto = nombre
    CalcularDestino = obj.Calcular()
    Set obj = Nothing
End Function

Private Function CalcularAlma(ByVal nombre As String) As Integer
    Dim obj As clsCalculoAlma
    Set obj = New clsCalculoAlma
    obj.NombreCompleto = nombre
    CalcularAlma = obj.Calcular()
    Set obj = Nothing
End Function

Private Function CalcularPersonalidad(ByVal nombre As String) As Integer
    Dim obj As clsCalculoPersonalidad
    Set obj = New clsCalculoPersonalidad
    obj.NombreCompleto = nombre
    CalcularPersonalidad = obj.Calcular()
    Set obj = Nothing
End Function

Private Function CalcularMadurez(ByVal nombre As String, ByVal fecha As Date) As Integer
    Dim obj As clsCalculoMadurez
    Set obj = New clsCalculoMadurez
    obj.NombreCompleto = nombre
    obj.FechaNacimiento = fecha
    CalcularMadurez = obj.Calcular()
    Set obj = Nothing
End Function

' ============================================================================
' MÉTODOS PRIVADOS - UTILIDADES
' ============================================================================

Private Function LeerArchivoInterpretacion(ByVal rutaCompleta As String) As String
    Dim numArchivo As Integer
    Dim contenido As String
    
    On Error GoTo ErrorHandler
    
    ' Verificar que existe el archivo
    If Dir(rutaCompleta) = "" Then
        LeerArchivoInterpretacion = "[Archivo no encontrado: " & rutaCompleta & "]"
        Exit Function
    End If
    
    ' Leer archivo
    numArchivo = FreeFile
    Open rutaCompleta For Input As #numArchivo
    contenido = Input$(LOF(numArchivo), numArchivo)
    Close #numArchivo
    
    LeerArchivoInterpretacion = contenido
    Exit Function
    
ErrorHandler:
    If numArchivo > 0 Then Close #numArchivo
    LeerArchivoInterpretacion = "[Error al leer archivo: " & Err.Description & "]"
End Function

Private Function ObtenerNombreTipo() As String
    Select Case pTipoSinastria
        Case TipoSinastria.General
            ObtenerNombreTipo = "General"
        Case TipoSinastria.Romantica
            ObtenerNombreTipo = "Romántica"
        Case TipoSinastria.Laboral
            ObtenerNombreTipo = "Laboral"
        Case Else
            ObtenerNombreTipo = "Desconocido"
    End Select
End Function

' ============================================================================
' EVENTOS DE CLASE
' ============================================================================

Private Sub Class_Initialize()
    ' Inicializar valores
    pNombre1 = ""
    pFecha1 = 0
    pNombre2 = ""
    pFecha2 = 0
    pTipoSinastria = TipoSinastria.General
    pRutaBase = CurrentProject.Path & "\Sinastrias\"
    pCalculosRealizados = False
    
    ' Inicializar números
    pCaminoVida1 = 0
    pDestino1 = 0
    pAlma1 = 0
    pPersonalidad1 = 0
    pMadurez1 = 0
    pCaminoVida2 = 0
    pDestino2 = 0
    pAlma2 = 0
    pPersonalidad2 = 0
    pMadurez2 = 0
End Sub

Private Sub Class_Terminate()
    ' Limpiar
End Sub