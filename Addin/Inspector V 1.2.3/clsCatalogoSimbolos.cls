VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCatalogoSimbolos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

'=====================================================
' Clase: clsCatalogoSimbolos_V2
' Catálogo global de símbolos declarados en el proyecto
'=====================================================

Private m_Simbolos As Collection   ' Colección de clsSimbolo

Private Sub Class_Initialize()
    Set m_Simbolos = New Collection
End Sub

'-----------------------------------------------------
' Añadir símbolo al catálogo
'-----------------------------------------------------
Public Sub Agregar(sim As clsSimbolo)
    Dim clave As String

    ' Validación básica
    If sim Is Nothing Then Exit Sub
    If Len(sim.nombre) = 0 Then Exit Sub
    If Len(sim.modulo) = 0 Then Exit Sub

    ' Clave compuesta
    clave = LCase(sim.modulo & "." & sim.nombre & "." & sim.Ambito & "." & sim.categoria)

    ' Evitar duplicados silenciosos
    If ExisteClave(clave) Then
        Debug.Print "Aviso: símbolo duplicado ignorado -> "; clave
        Exit Sub
    End If

    m_Simbolos.Add sim, clave
End Sub

'-----------------------------------------------------
' Comprobar si existe una clave
'-----------------------------------------------------
Public Function ExisteClave(clave As String) As Boolean
    On Error Resume Next
    Dim tmp As Object
    Set tmp = m_Simbolos(clave)
    ExisteClave = (Err.Number = 0)
    Err.Clear
End Function

'-----------------------------------------------------
' Acceso a la colección completa
'-----------------------------------------------------
Public Property Get Simbolos() As Collection
    Set Simbolos = m_Simbolos
End Property

'-----------------------------------------------------
' Contar símbolos
'-----------------------------------------------------
Public Function Contar() As Long
    Contar = m_Simbolos.Count
End Function

'-----------------------------------------------------
' Buscar por nombre exacto
'-----------------------------------------------------
Public Function BuscarPorNombre(ByVal nombre As String) As Collection
    Dim colRes As New Collection
    Dim s As clsSimbolo

    For Each s In m_Simbolos
        If LCase$(s.nombre) = LCase$(nombre) Then
            colRes.Add s
        End If
    Next s

    Set BuscarPorNombre = colRes
End Function

'-----------------------------------------------------
' Buscar por módulo
'-----------------------------------------------------
Public Function BuscarPorModulo(ByVal modulo As String) As Collection
    Dim colRes As New Collection
    Dim s As clsSimbolo

    For Each s In m_Simbolos
        If LCase$(s.modulo) = LCase$(modulo) Then
            colRes.Add s
        End If
    Next s

    Set BuscarPorModulo = colRes
End Function

'-----------------------------------------------------
' Buscar por categoría (Variable, Constante, Función, etc.)
'-----------------------------------------------------
Public Function BuscarPorCategoria(ByVal categoria As String) As Collection
    Dim colRes As New Collection
    Dim s As clsSimbolo

    For Each s In m_Simbolos
        If LCase$(s.categoria) = LCase$(categoria) Then
            colRes.Add s
        End If
    Next s

    Set BuscarPorCategoria = colRes
End Function

'-----------------------------------------------------
' Símbolos no usados
'-----------------------------------------------------
Public Function SimbolosNoUsados() As Collection
    Dim colRes As New Collection
    Dim s As clsSimbolo

    For Each s In m_Simbolos
        If s.Usado = False Then
            colRes.Add s
        End If
    Next s

    Set SimbolosNoUsados = colRes
End Function

'-----------------------------------------------------
' Limpiar catálogo
'-----------------------------------------------------
Public Sub Limpiar()
    Set m_Simbolos = New Collection
End Sub

