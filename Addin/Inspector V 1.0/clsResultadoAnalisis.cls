VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsResultadoAnalisis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

Public Enum TipoElementoInspector
    tiClase = 1
    tiModulo = 2
    tiReferencia = 3
    tiProyecto = 4
End Enum

Public Enum SeveridadInspector
    sevInfo = 1
    sevAviso = 2
    sevError = 3
End Enum

Private pTipoElemento As TipoElementoInspector
Private pNombreElemento As String
Private pSeveridad As SeveridadInspector
Private pDescripcion As String
Private pEsReparable As Boolean
Private pCodigoReparacion As String

Public Property Get TipoElemento() As TipoElementoInspector
    TipoElemento = pTipoElemento
End Property

Public Property Let TipoElemento(ByVal valor As TipoElementoInspector)
    pTipoElemento = valor
End Property

Public Property Get NombreElemento() As String
    NombreElemento = pNombreElemento
End Property

Public Property Let NombreElemento(ByVal valor As String)
    pNombreElemento = valor
End Property

Public Property Get Severidad() As SeveridadInspector
    Severidad = pSeveridad
End Property

Public Property Let Severidad(ByVal valor As SeveridadInspector)
    pSeveridad = valor
End Property

Public Property Get descripcion() As String
    descripcion = pDescripcion
End Property

Public Property Let descripcion(ByVal valor As String)
    pDescripcion = valor
End Property

Public Property Get esReparable() As Boolean
    esReparable = pEsReparable
End Property

Public Property Let esReparable(ByVal valor As Boolean)
    pEsReparable = valor
End Property

Public Property Get codigoReparacion() As String
    codigoReparacion = pCodigoReparacion
End Property

Public Property Let codigoReparacion(ByVal valor As String)
    pCodigoReparacion = valor
End Property

Public Function Formatear() As String
    Dim s As String
    Dim sSev As String
    Dim sTipo As String
    
    Select Case pSeveridad
        Case sevInfo: sSev = "INFO"
        Case sevAviso: sSev = "AVISO"
        Case sevError: sSev = "ERROR"
    End Select
    
    Select Case pTipoElemento
        Case tiClase: sTipo = "CLASE"
        Case tiModulo: sTipo = "MÓDULO"
        Case tiReferencia: sTipo = "REFERENCIA"
        Case tiProyecto: sTipo = "PROYECTO"
    End Select
    
    s = "[" & sSev & "][" & sTipo & "] " & pNombreElemento & " - " & pDescripcion
    
    If pEsReparable Then
        s = s & " [Reparable: " & pCodigoReparacion & "]"
    End If
    
    Formatear = s
End Function



'Option Compare Database
'Option Explicit
'
'' Representa un hallazgo del inspector
'Public Enum TipoElementoInspector
'    tiClase = 1
'    tiModulo = 2
'    tiReferencia = 3
'    tiProyecto = 4
'End Enum
'
'Public Enum SeveridadInspector
'    sevInfo = 1
'    sevAviso = 2
'    sevError = 3
'End Enum
'
'Private pTipoElemento As TipoElementoInspector
'Private pNombreElemento As String
'Private pSeveridad As SeveridadInspector
'Private pDescripcion As String
'
'Private pEsReparable As Boolean
'Private pCodigoReparacion As String   ' identificador de la acción de reparación
'
'Public Property Get esReparable() As Boolean
'    esReparable = pEsReparable
'End Property
'
'Public Property Let esReparable(ByVal valor As Boolean)
'    pEsReparable = valor
'End Property
'
'Public Property Get codigoReparacion() As String
'    codigoReparacion = pCodigoReparacion
'End Property
'
'Public Property Let codigoReparacion(ByVal valor As String)
'    pCodigoReparacion = valor
'End Property
'
'
'Public Property Get TipoElemento() As TipoElementoInspector
'    TipoElemento = pTipoElemento
'End Property
'
'Public Property Let TipoElemento(ByVal valor As TipoElementoInspector)
'    pTipoElemento = valor
'End Property
'
'Public Property Get NombreElemento() As String
'    ReturnNombreElemento = pNombreElemento
'End Property
'
'Public Property Let NombreElemento(ByVal valor As String)
'    pNombreElemento = valor
'End Property
'
'Public Property Get Severidad() As SeveridadInspector
'    Severidad = pSeveridad
'End Property
'
'Public Property Let Severidad(ByVal valor As SeveridadInspector)
'    pSeveridad = valor
'End Property
'
'Public Property Get descripcion() As String
'    descripcion = pDescripcion
'End Property
'
'Public Property Let descripcion(ByVal valor As String)
'    pDescripcion = valor
'End Property
'
'Public Function Formatear() As String
'    Dim s As String
'    Dim sSev As String
'    Dim sTipo As String
'
'    Select Case pSeveridad
'        Case sevInfo: sSev = "INFO"
'        Case sevAviso: sSev = "AVISO"
'        Case sevError: sSev = "ERROR"
'    End Select
'
'    Select Case pTipoElemento
'        Case tiClase: sTipo = "CLASE"
'        Case tiModulo: sTipo = "MÓDULO"
'        Case tiReferencia: sTipo = "REFERENCIA"
'        Case tiProyecto: sTipo = "PROYECTO"
'    End Select
'
'    s = "[" & sSev & "][" & sTipo & "] " & pNombreElemento & " - " & pDescripcion
'    If pEsReparable Then
'        s = s & vbCrLf & " [Reparable: " & pCodigoReparacion & "]"
'    End If
'
'
'    Formatear = s
'End Function
'
