VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCatalogoInspector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

'=====================================================
' Clase: clsCatalogoInspector
' Catálogo completo del proyecto analizado
'=====================================================

Private mModulos As Collection
Private mClases As Collection
Private mUserForms As Collection
Private mFormularios As Collection
Private mInformes As Collection
Private mOtros As Collection
Private mCatalogoSimbolos As clsCatalogoSimbolos


Private Sub Class_Initialize()
    Set mModulos = New Collection
    Set mClases = New Collection
    Set mUserForms = New Collection
    Set mFormularios = New Collection
    Set mInformes = New Collection
    Set mOtros = New Collection
    Set mCatalogoSimbolos = New clsCatalogoSimbolos
End Sub

'-----------------------------------------------------
' Propiedades de acceso
'-----------------------------------------------------
Public Property Get CatalogoSimbolos() As clsCatalogoSimbolos
    Set CatalogoSimbolos = mCatalogoSimbolos
End Property

Public Property Set CatalogoSimbolos(val As clsCatalogoSimbolos)
    Set mCatalogoSimbolos = val
End Property


Public Property Get Modulos() As Collection
    Set Modulos = mModulos
End Property

Public Property Get Clases() As Collection
    Set Clases = mClases
End Property

Public Property Get UserForms() As Collection
    Set UserForms = mUserForms
End Property

Public Property Get Formularios() As Collection
    Set Formularios = mFormularios
End Property

Public Property Get Informes() As Collection
    Set Informes = mInformes
End Property

Public Property Get Otros() As Collection
    Set Otros = mOtros
End Property

'-----------------------------------------------------
' Métodos de agregado (con validación y protección)
'-----------------------------------------------------
Private Function NormalizarNombre(ByVal nombre As String) As String
    NormalizarNombre = LCase$(Trim$(nombre))
End Function

Private Function ExisteEn(col As Collection, ByVal nombre As String) As Boolean
    On Error Resume Next
    Dim tmp As Object
    Set tmp = col(NormalizarNombre(nombre))
    ExisteEn = (Err.Number = 0)
    Err.Clear
End Function

Public Sub AgregarModulo(ByVal m As clsModulo)
    If m Is Nothing Then Exit Sub
    If ExisteEn(mModulos, m.nombre) Then Exit Sub
    mModulos.Add m, NormalizarNombre(m.nombre)
End Sub

Public Sub AgregarClase(ByVal c As clsClase)
    If c Is Nothing Then Exit Sub
    If ExisteEn(mClases, c.nombre) Then Exit Sub
    mClases.Add c, NormalizarNombre(c.nombre)
End Sub

Public Sub AgregarUserForm(ByVal m As clsModulo)
    If m Is Nothing Then Exit Sub
    If ExisteEn(mUserForms, m.nombre) Then Exit Sub
    mUserForms.Add m, NormalizarNombre(m.nombre)
End Sub

Public Sub AgregarFormulario(ByVal m As clsModulo)
    If m Is Nothing Then Exit Sub
    If ExisteEn(mFormularios, m.nombre) Then Exit Sub
    mFormularios.Add m, NormalizarNombre(m.nombre)
End Sub

Public Sub AgregarInforme(ByVal m As clsModulo)
    If m Is Nothing Then Exit Sub
    If ExisteEn(mInformes, m.nombre) Then Exit Sub
    mInformes.Add m, NormalizarNombre(m.nombre)
End Sub

Public Sub AgregarOtro(ByVal m As clsModulo)
    If m Is Nothing Then Exit Sub
    If ExisteEn(mOtros, m.nombre) Then Exit Sub
    mOtros.Add m, NormalizarNombre(m.nombre)
End Sub

'-----------------------------------------------------
' Métodos de utilidad
'-----------------------------------------------------
Public Function ObtenerModulo(ByVal nombre As String) As clsModulo
    On Error Resume Next
    Set ObtenerModulo = mModulos(NormalizarNombre(nombre))
End Function

Public Function ObtenerClase(ByVal nombre As String) As clsClase
    On Error Resume Next
    Set ObtenerClase = mClases(NormalizarNombre(nombre))
End Function

Public Sub Limpiar()
    Set mModulos = New Collection
    Set mClases = New Collection
    Set mUserForms = New Collection
    Set mFormularios = New Collection
    Set mInformes = New Collection
    Set mOtros = New Collection
End Sub

