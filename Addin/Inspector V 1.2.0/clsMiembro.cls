VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMiembro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

'=====================================================
' Clase: clsMiembro
' Representa un procedimiento, función, propiedad o evento
'=====================================================

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

Private mNombre As String
Private mTipo As TipoMiembroInspector
Private mAmbito As AmbitoMiembroInspector
Private mLineaInicio As Long
Private mLineaFin As Long
Private mNumLineas As Long
Private mNumLineasCodigo As Long
Private mNumLineasComentario As Long
Private mNumLineasVacias As Long
Private mNumLineasAtributo As Long

Private mEsPropiedad As Boolean
Private mEsFuncion As Boolean
Private mEsProcedimiento As Boolean
Private mEsEvento As Boolean
Private mEsAPI As Boolean

Private mFirmaOriginal As String

'-----------------------------------------------------
' Propiedades básicas
'-----------------------------------------------------
Public Property Get nombre() As String
    nombre = mNombre
End Property

Public Property Let nombre(ByVal Value As String)
    mNombre = Trim$(Value)
End Property

Public Property Get tipo() As TipoMiembroInspector
    tipo = mTipo
End Property

Public Property Let tipo(ByVal Value As TipoMiembroInspector)
    mTipo = Value
End Property

Public Property Get Ambito() As AmbitoMiembroInspector
    Ambito = mAmbito
End Property

Public Property Let Ambito(ByVal Value As AmbitoMiembroInspector)
    mAmbito = Value
End Property

Public Property Get LineaInicio() As Long
    LineaInicio = mLineaInicio
End Property

Public Property Let LineaInicio(ByVal Value As Long)
    If Value >= 0 Then mLineaInicio = Value
End Property

Public Property Get LineaFin() As Long
    LineaFin = mLineaFin
End Property

Public Property Let LineaFin(ByVal Value As Long)
    If Value >= 0 Then mLineaFin = Value
End Property

Public Property Get FirmaOriginal() As String
    FirmaOriginal = mFirmaOriginal
End Property

Public Property Let FirmaOriginal(ByVal Value As String)
    mFirmaOriginal = Trim$(Value)
End Property

'-----------------------------------------------------
' Métricas
'-----------------------------------------------------
Public Property Get NumLineas() As Long
    NumLineas = mNumLineas
End Property

Public Property Get NumLineasCodigo() As Long
    NumLineasCodigo = mNumLineasCodigo
End Property

Public Property Get NumLineasComentario() As Long
    NumLineasComentario = mNumLineasComentario
End Property

Public Property Get NumLineasVacias() As Long
    NumLineasVacias = mNumLineasVacias
End Property

Public Property Get NumLineasAtributo() As Long
    NumLineasAtributo = mNumLineasAtributo
End Property

'-----------------------------------------------------
' Flags semánticos
'-----------------------------------------------------
Public Property Get EsPropiedad() As Boolean
    EsPropiedad = mEsPropiedad
End Property

Public Property Get EsFuncion() As Boolean
    EsFuncion = mEsFuncion
End Property

Public Property Get EsProcedimiento() As Boolean
    EsProcedimiento = mEsProcedimiento
End Property

Public Property Get EsEvento() As Boolean
    EsEvento = mEsEvento
End Property

Public Property Get EsAPI() As Boolean
    EsAPI = mEsAPI
End Property

'-----------------------------------------------------
' Clasificación desde la firma
'-----------------------------------------------------
Public Sub EstablecerTipoDesdeTexto(ByVal firma As String)
    Dim t As String
    t = LCase$(Trim$(firma))
    mFirmaOriginal = firma

    mEsPropiedad = False
    mEsFuncion = False
    mEsProcedimiento = False
    mEsEvento = False
    mEsAPI = False

    If Left$(t, 12) = "property get" Then
        mTipo = tmPropertyGet: mEsPropiedad = True

    ElseIf Left$(t, 12) = "property let" Then
        mTipo = tmPropertyLet: mEsPropiedad = True

    ElseIf Left$(t, 12) = "property set" Then
        mTipo = tmPropertySet: mEsPropiedad = True

    ElseIf Left$(t, 3) = "sub" Or InStr(t, " sub ") > 0 Then
        mTipo = tmSub: mEsProcedimiento = True

    ElseIf Left$(t, 8) = "function" Or InStr(t, " function ") > 0 Then
        mTipo = tmFunction: mEsFuncion = True

    ElseIf Left$(t, 5) = "event" Or InStr(t, " event ") > 0 Then
        mTipo = tmEvent: mEsEvento = True

    ElseIf Left$(t, 7) = "declare" Or InStr(t, " declare ") > 0 Then
        mTipo = tmDeclare: mEsAPI = True

    Else
        mTipo = tmUnknown
    End If
End Sub

Public Sub EstablecerAmbitoDesdeTexto(ByVal firma As String)
    Dim t As String
    t = LCase$(Trim$(firma))

    If Left$(t, 6) = "public" Then
        mAmbito = amPublic
    ElseIf Left$(t, 7) = "private" Then
        mAmbito = amPrivate
    ElseIf Left$(t, 6) = "friend" Then
        mAmbito = amFriend
    Else
        mAmbito = amDefault
    End If
End Sub

'-----------------------------------------------------
' Métricas de líneas
'-----------------------------------------------------
Friend Sub IncrementarMetricas(ByVal tipoLinea As String)
    Select Case tipoLinea
        Case "Codigo":     mNumLineasCodigo = mNumLineasCodigo + 1
        Case "Comentario": mNumLineasComentario = mNumLineasComentario + 1
        Case "Vacia":      mNumLineasVacias = mNumLineasVacias + 1
        Case "Atributo":   mNumLineasAtributo = mNumLineasAtributo + 1
    End Select

    mNumLineas = mNumLineasCodigo + mNumLineasComentario + mNumLineasVacias + mNumLineasAtributo
End Sub

'-----------------------------------------------------
' Utilidades
'-----------------------------------------------------
Public Function EsPropertyCompleta() As Boolean
    EsPropertyCompleta = (mTipo = tmPropertyGet Or mTipo = tmPropertyLet Or mTipo = tmPropertySet)
End Function

Public Sub Limpiar()
    mNombre = ""
    mTipo = tmUnknown
    mAmbito = amDefault
    mLineaInicio = 0
    mLineaFin = 0
    mNumLineas = 0
    mNumLineasCodigo = 0
    mNumLineasComentario = 0
    mNumLineasVacias = 0
    mNumLineasAtributo = 0
    mEsPropiedad = False
    mEsFuncion = False
    mEsProcedimiento = False
    mEsEvento = False
    mEsAPI = False
    mFirmaOriginal = ""
End Sub

