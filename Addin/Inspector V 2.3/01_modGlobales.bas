Attribute VB_Name = "01_modGlobales"

Option Compare Database
Option Explicit

'===============================================================
' Módulo: 01_modGlobales
' Contenedor centralizado de:
'   - Enumeraciones
'   - Constantes
'   - Variables globales
'   - Estado del Inspector
'===============================================================


'===============================================================
' Enumeraciones
'===============================================================

' Estilos HTML usados en exportación
Public Enum EstiloHtml
    TemaClaro = 0
    TemaOscuro
    TemaSepia
    TemaContraste
    TemaMinimalista
End Enum
'---------------------------------------------------------------

' Formatos de exportación agrupados por tipo
Public Enum FormatoExportacion
    ' TXT (0–9)
    ExpResultadosTXT = 0
    ExpSimbolosTXT = 1
    ExpTodoTXT = 2

    ' Excel (10–19)
    ExpResultadosExcel = 10
    ExpSimbolosExcel = 11
    ExpTodoExcel = 12

    ' HTML (20–29)
    ExpTodoHTML = 20

    ' Futuros formatos (30–39)
    'ExpTodoAccess = 30
    'ExpTodoPDF = 31
    'ExpTodoWord = 32
    'ExpTodoMarkdown = 33
End Enum
'---------------------------------------------------------------

' Encontrado en:
'   - clsMiembro
' Representa un procedimiento, función, propiedad o evento

Public Enum TipoMiembroInspector
    tmSub = 0
    tmFunction = 1
    tmPropertyGet = 2
    tmPropertyLet = 3
    tmPropertySet = 4
    tmEvent = 5
    tmDeclare = 6
    tmUnknown = 99
End Enum

Public Enum AmbitoMiembroInspector
    amPublic = 0
    amPrivate = 1
    amFriend = 2
    amDefault = 3
End Enum
'---------------------------------------------------------------

' Encontrado en:
'   - clsModulo
' Representa un módulo del proyecto VBA

Public Enum TipoModuloInspector
    tmStdModule = 0
    tmFormModule = 1
    tmReportModule = 2
    tmUserForm = 3
    tmUnknownModule = 99
End Enum
'---------------------------------------------------------------

' Encontrado en:
'   - clsResultadoAnalisis
' Representa un resultado generado por una regla del Inspector

Public Enum SeveridadInspector
    sevInfo = 0
    sevAviso = 1
    sevError = 2
End Enum

Public Enum TipoElementoInspector
    teProyecto = 0
    teModulo = 1
    teClase = 2
    teUserForm = 3
    teFormulario = 4
    teInforme = 5
    teMiembro = 6
End Enum
'---------------------------------------------------------------
'===============================================================
' Enumeraciones globales del Inspector
'   - Estados de análisis
'   - Estados de reparación
'   - Estados de exportación
'===============================================================

'---------------------------------------------------------------
' Estado del análisis
'---------------------------------------------------------------
Public Enum EstadoAnalisis
    AnalisisNoEjecutado = 0
    AnalisisEjecutado = 1
    AnalisisConErrores = 2     ' reservado para futuro
End Enum


'---------------------------------------------------------------
' Estado de la reparación
'---------------------------------------------------------------
Public Enum EstadoReparacion
    ReparacionNoEjecutada = 0
    ReparacionEjecutada = 1
    ReparacionConErrores = 2   ' reservado para futuro
End Enum


'---------------------------------------------------------------
' Estado de la exportación
'---------------------------------------------------------------
Public Enum EstadoExportacion
    ExportacionNoEjecutada = 0
    ExportacionEjecutada = 1
    ExportacionConErrores = 2  ' reservado para futuro
End Enum


'---------------------------------------------------------------
' Palabras reservadas
'---------------------------------------------------------------
Public Enum ReservedCategory
    rcVBA = 1
    rcSQL_JET
    rcSQL_ACE
    rcACCESS
    rcDAO
    rcADO
    rcVB6
    rcOPERADORES
    rcFUNCIONES_VBA
    rcCONTROL_FLUJO_SISTEMA
End Enum

Public Type ReservedWordInfo
    nombre As String
    Tipo As String
    Categoria As ReservedCategory
End Type

'Enum ReservedCategory
'    VBA = 1
'    SQL_JET
'    SQL_ACE
'    Access
'    DAO
'    ADO
'    VB6
'    OPERADORES
'    FUNCIONES_VBA
'    CONTROL_FLUJO_SISTEMA
'End Enum
'
'Type ReservedWordInfo
'    Nombre As String
'    Tipo As String
'    Categoria As ReservedCategory
'End Type


'===============================================================
' Objetos globales
'===============================================================

' Encontrado en:
'   - modEntornoInspector
Public colAddins As Collection   ' Colección de clsAddin
'---------------------------------------------------------------

' Encontrado en:
'   - modInspectorMain
' Objetos globales
Public gCatalogoInspector As clsCatalogoInspector
Public gResultadosInspector As clsResultados
'---------------------------------------------------------------

' Encontrado en:
'   - modSimbolosInspector
Public gCatalogoSimbolos As clsCatalogoSimbolos
'---------------------------------------------------------------

' Encontrado en:
'   - clsAddin
' NO SE DEBEN MOVER DE LA CLASE, SUSTITUYEN A LAS PROPIEDADES
''-------------------------
'' Datos básicos
''-------------------------
'Public addin_Name As String      ' Nombre del Add-In (ej: InspectorVBA.accda)
'Public library As String         ' Ruta completa del archivo
'
''-------------------------
'' Información derivada
''-------------------------
'Public BaseName As String        ' Nombre sin extensión
'Public Extension As String       ' Extensión (accda/mda)
'Public Folder As String          ' Carpeta donde está el Add-In
'Public Exists As Boolean         ' ¿Existe físicamente el archivo?
'Public Loaded As Boolean         ' ¿Está cargado en Access?
'Public Version As String         ' Versión del archivo
'Public Description As String     ' Descripción del archivo
'---------------------------------------------------------------

' Encontrado en:
'   - clsSimbolo
' NO SE DEBEN MOVER DE LA CLASE, SUSTITUYEN A LAS PROPIEDADES
'' Representa un símbolo declarado en el proyecto VBA
'
'Public nombre As String          ' Nombre del símbolo
'Public TipoTexto As String       ' Tipo declarado (Long, String, TipoCliente...)
'Public categoria As String       ' Variable, Constante, Enum, UDT, Propiedad...
'Public Ambito As String          ' Local, ModuloPrivate, ModuloPublic, Global
'Public modulo As String          ' Módulo donde se declara
'Public miembro As String         ' Miembro donde se declara (si es local)
'Public LineaDeclaracion As Long  ' Línea de declaración
'Public Usado As Boolean          ' ¿Se ha encontrado referencia?
'
'' Información contextual opcional
'Public EsAPI As Boolean          ' Para futuras detecciones de Declare
'Public Comentario As String      ' Notas adicionales
'---------------------------------------------------------------

' Estado del Inspector
Public gVBIDEDisponible As Boolean

Public gUltimaRuta As String
Public gUltimoFormato As FormatoExportacion
Public gUltimoEstiloHtml As EstiloHtml
'---------------------------------------------------------------

'Palabras reservadas Access
Public gPalabrasReservadas As Object ' Dictionary


'===============================================================
' Constantes del Inspector
'===============================================================
Public Const NOMBRE_PRODUCTO As String = "InspectorVBA"
Public Const EXT_TXT As String = ".txt"
Public Const EXT_HTML As String = ".html"
Public Const EXT_XLSX As String = ".xlsx"

