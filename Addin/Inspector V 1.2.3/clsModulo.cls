VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsModulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

'=====================================================
' Clase: clsModulo
' Representa un módulo del proyecto VBA
'=====================================================

Private mNombre As String
Private mTipo As TipoModuloInspector
Private mLineas() As String
Private mNumLineas As Long
Private mNumLineasCodigo As Long
Private mNumLineasComentario As Long
Private mNumLineasVacias As Long
Private mNumLineasAtributo As Long
Private mMiembros As Collection

Private Sub Class_Initialize()
    Set mMiembros = New Collection
End Sub

'-----------------------------------------------------
' Nombre del módulo
'-----------------------------------------------------
Public Property Get nombre() As String
    nombre = mNombre
End Property

Public Property Let nombre(ByVal Value As String)
    mNombre = Trim$(Value)
End Property

'-----------------------------------------------------
' Tipo del módulo
'-----------------------------------------------------
Public Property Get tipo() As TipoModuloInspector
    tipo = mTipo
End Property

Public Property Let tipo(ByVal Value As TipoModuloInspector)
    mTipo = Value
End Property

'-----------------------------------------------------
' Líneas del módulo
'-----------------------------------------------------
Public Property Get lineas() As Variant
    lineas = mLineas
End Property

Public Property Let lineas(ByVal Value As Variant)
    Dim i As Long
    Dim tmp() As String

    If IsArray(Value) Then
        ReDim tmp(LBound(Value) To UBound(Value))
        For i = LBound(Value) To UBound(Value)
            tmp(i) = CStr(Value(i))
        Next i
        mLineas = tmp
        mNumLineas = UBound(mLineas) - LBound(mLineas) + 1
    End If
End Property

'-----------------------------------------------------
' Métricas
'-----------------------------------------------------
Public Property Get numLineas() As Long
    numLineas = mNumLineas
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
' Miembros
'-----------------------------------------------------
Public Property Get Miembros() As Collection
    Set Miembros = mMiembros
End Property

Public Property Set Miembros(ByVal Value As Collection)
    Set mMiembros = Value
End Property

'-----------------------------------------------------
' Calcular métricas del módulo
'-----------------------------------------------------
Public Sub CalcularMetricasModulo()
    Dim i As Long
    Dim tipoLinea As String

    mNumLineasCodigo = 0
    mNumLineasComentario = 0
    mNumLineasVacias = 0
    mNumLineasAtributo = 0

    If mNumLineas = 0 Then Exit Sub

    For i = LBound(mLineas) To UBound(mLineas)
        tipoLinea = ClasificarLinea(mLineas(i))
        Select Case tipoLinea
            Case "Codigo":     mNumLineasCodigo = mNumLineasCodigo + 1
            Case "Comentario": mNumLineasComentario = mNumLineasComentario + 1
            Case "Vacia":      mNumLineasVacias = mNumLineasVacias + 1
            Case "Atributo":   mNumLineasAtributo = mNumLineasAtributo + 1
        End Select
    Next i
End Sub

'-----------------------------------------------------
' Clasificación de líneas
'-----------------------------------------------------
Private Function ClasificarLinea(ByVal linea As String) As String
    Dim t As String
    t = Trim$(linea)

    If t = "" Then
        ClasificarLinea = "Vacia"

    ElseIf Left$(t, 1) = "'" Or LCase$(Left$(t, 3)) = "rem" Then
        ClasificarLinea = "Comentario"

    ElseIf LCase$(Left$(t, 9)) = "attribute" Then
        ClasificarLinea = "Atributo"

    Else
        ClasificarLinea = "Codigo"
    End If
End Function

'-----------------------------------------------------
' Utilidades
'-----------------------------------------------------
Public Function ObtenerLinea(ByVal indice As Long) As String
    If indice >= LBound(mLineas) And indice <= UBound(mLineas) Then
        ObtenerLinea = mLineas(indice)
    Else
        ObtenerLinea = ""
    End If
End Function

Public Function TieneOptionExplicit() As Boolean
    Dim i As Long
    Dim t As String

    For i = LBound(mLineas) To Application.Min(UBound(mLineas), 10)
        t = Trim$(LCase$(mLineas(i)))
        If t = "option explicit" Then
            TieneOptionExplicit = True
            Exit Function
        End If
        If t <> "" And Left$(t, 1) <> "'" And LCase$(Left$(t, 9)) <> "attribute" Then
            Exit For
        End If
    Next i
End Function

Public Function EsUserForm() As Boolean
    EsUserForm = (mTipo = tmUserForm)
End Function

Public Function EsFormulario() As Boolean
    EsFormulario = (mTipo = tmFormModule)
End Function

Public Function EsInforme() As Boolean
    EsInforme = (mTipo = tmReportModule)
End Function

Public Sub Limpiar()
    Erase mLineas
    mNumLineas = 0
    mNumLineasCodigo = 0
    mNumLineasComentario = 0
    mNumLineasVacias = 0
    mNumLineasAtributo = 0
    Set mMiembros = New Collection
End Sub

