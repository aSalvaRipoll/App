VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsResultados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

'===============================================================
' Clase: clsResultados
' Contiene todos los resultados del análisis del Inspector
'===============================================================

Private mResultados As Collection

'---------------------------------------------------------------
' Inicialización
'---------------------------------------------------------------
Private Sub Class_Initialize()
    Set mResultados = New Collection
End Sub

'---------------------------------------------------------------
' Agregar un resultado
'---------------------------------------------------------------
Public Sub AgregarResultado(res As clsResultadoAnalisis)
    On Error Resume Next
    mResultados.Add res, res.clave
End Sub

'---------------------------------------------------------------
' Obtener un resultado por clave
'---------------------------------------------------------------
Public Function ObtenerResultado(clave As String) As clsResultadoAnalisis
    On Error Resume Next
    Set ObtenerResultado = mResultados(clave)
End Function

'---------------------------------------------------------------
' Navegar a un resultado concreto
'---------------------------------------------------------------
Public Sub NavegarAResultado(clave As String)

    Dim res As clsResultadoAnalisis
    Set res = Me.ObtenerResultado(clave)

    If res Is Nothing Then
        Inspector_Log "NavegarAResultado: resultado no encontrado (" & clave & ")"
        Exit Sub
    End If

    ' Llamada al módulo de navegación
    Call NavegarAModulo(res.nombreElemento, res.nombreMiembro, res.linea)

End Sub

'---------------------------------------------------------------
' Enumerar resultados (si lo necesitas en el futuro)
'---------------------------------------------------------------
Public Function Todos() As Collection
    Set Todos = mResultados
End Function

