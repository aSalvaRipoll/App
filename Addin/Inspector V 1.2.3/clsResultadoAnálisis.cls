VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsResultadoAnálisis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

'===============================================================
' Clase: clsResultadoAnalisis
' Representa un resultado generado por una regla del Inspector
'===============================================================

Private mSeveridad As SeveridadInspector
Private mTipoElemento As TipoElementoInspector
Private mNombreElemento As String
Private mNombreMiembro As String
Private mDescripcion As String
Private mCodigoRegla As String
Private mCodigoReparacion As String
Private mEsReparable As Boolean
Private mLinea As Long
Private mDetalles As String

'---------------------------------------------------------------
' Propiedades básicas
'---------------------------------------------------------------
Public Property Get Severidad() As SeveridadInspector
    Severidad = mSeveridad
End Property

Public Property Let Severidad(ByVal Value As SeveridadInspector)
    mSeveridad = Value
End Property

Public Property Get TipoElemento() As TipoElementoInspector
    TipoElemento = mTipoElemento
End Property

Public Property Let TipoElemento(ByVal Value As TipoElementoInspector)
    mTipoElemento = Value
End Property

Public Property Get nombreElemento() As String
    nombreElemento = mNombreElemento
End Property

Public Property Let nombreElemento(ByVal Value As String)
    mNombreElemento = Trim$(Value)
End Property

Public Property Get nombreMiembro() As String
    nombreMiembro = mNombreMiembro
End Property

Public Property Let nombreMiembro(ByVal Value As String)
    mNombreMiembro = Trim$(Value)
End Property

Public Property Get descripcion() As String
    descripcion = mDescripcion
End Property

Public Property Let descripcion(ByVal Value As String)
    mDescripcion = Trim$(Value)
End Property

Public Property Get codigoRegla() As String
    codigoRegla = mCodigoRegla
End Property

Public Property Let codigoRegla(ByVal Value As String)
    mCodigoRegla = Trim$(Value)
End Property

Public Property Get codigoReparacion() As String
    codigoReparacion = mCodigoReparacion
End Property

Public Property Let codigoReparacion(ByVal Value As String)
    mCodigoReparacion = Trim$(Value)
End Property

Public Property Get esReparable() As Boolean
    esReparable = mEsReparable
End Property

Public Property Let esReparable(ByVal Value As Boolean)
    mEsReparable = Value
End Property

Public Property Get linea() As Long
    linea = mLinea
End Property

Public Property Let linea(ByVal Value As Long)
    If Value >= 0 Then mLinea = Value
End Property

Public Property Get Detalles() As String
    Detalles = mDetalles
End Property

Public Property Let Detalles(ByVal Value As String)
    mDetalles = Trim$(Value)
End Property

'---------------------------------------------------------------
' Propiedades calculadas
'---------------------------------------------------------------
Public Property Get clave() As String
    ' Clave única para evitar duplicados
    clave = LCase$(mCodigoRegla & "|" & mTipoElemento & "|" & mNombreElemento & "|" & mNombreMiembro & "|" & mLinea)
End Property

Public Property Get EsError() As Boolean
    EsError = (mSeveridad = sevError)
End Property

Public Property Get EsAviso() As Boolean
    EsAviso = (mSeveridad = sevAviso)
End Property

Public Property Get EsInfo() As Boolean
    EsInfo = (mSeveridad = sevInfo)
End Property

Public Property Get EsDeModulo() As Boolean
    EsDeModulo = (mTipoElemento = teModulo)
End Property

Public Property Get EsDeClase() As Boolean
    EsDeClase = (mTipoElemento = teClase)
End Property

Public Property Get EsDeMiembro() As Boolean
    EsDeMiembro = (mTipoElemento = teMiembro)
End Property

'---------------------------------------------------------------
' Formateo
'---------------------------------------------------------------
Public Function Formatear() As String
    Dim s As String

    s = "[" & mCodigoRegla & "] "

    Select Case mSeveridad
        Case sevInfo:  s = s & "(INFO) "
        Case sevAviso: s = s & "(AVISO) "
        Case sevError: s = s & "(ERROR) "
    End Select

    s = s & TipoElementoToText(mTipoElemento) & " " & mNombreElemento

    If Len(mNombreMiembro) > 0 Then
        s = s & "." & mNombreMiembro
    End If

    If mLinea > 0 Then
        s = s & " (Línea " & CStr(mLinea) & ")"
    End If

    s = s & " - " & mDescripcion

    Formatear = s
End Function

Public Function FormatearCorto() As String
    Dim s As String
    s = "[" & mCodigoRegla & "] " & mNombreElemento
    If Len(mNombreMiembro) > 0 Then s = s & "." & mNombreMiembro
    FormatearCorto = s
End Function

'---------------------------------------------------------------
' Utilidades
'---------------------------------------------------------------
Private Function TipoElementoToText(t As TipoElementoInspector) As String
    Select Case t
        Case teProyecto:   TipoElementoToText = "Proyecto"
        Case teModulo:     TipoElementoToText = "Módulo"
        Case teClase:      TipoElementoToText = "Clase"
        Case teUserForm:   TipoElementoToText = "UserForm"
        Case teFormulario: TipoElementoToText = "Formulario"
        Case teInforme:    TipoElementoToText = "Informe"
        Case teMiembro:    TipoElementoToText = "Miembro"
        Case Else:         TipoElementoToText = "Elemento"
    End Select
End Function

Public Function ToDebugString() As String
    ToDebugString = "Resultado: " & Formatear & " | Detalles: " & mDetalles
End Function

Public Function ToTextFile() As String
    Dim s As String

    s = "[" & mCodigoRegla & "] "

    Select Case mSeveridad
        Case sevInfo:  s = s & "INFO"
        Case sevAviso: s = s & "AVISO"
        Case sevError: s = s & "ERROR"
    End Select

    s = s & " | " & TipoElementoToText(mTipoElemento)
    s = s & " | " & mNombreElemento

    If Len(mNombreMiembro) > 0 Then
        s = s & "." & mNombreMiembro
    End If

    If mLinea > 0 Then
        s = s & " | Línea " & CStr(mLinea)
    End If

    s = s & " | " & mDescripcion

    If Len(mDetalles) > 0 Then
        s = s & " | Detalles: " & mDetalles
    End If

    ToTextFile = s
End Function

