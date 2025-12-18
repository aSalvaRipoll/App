VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInformeSimbolos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

'===============================================================
' Clase: clsInformeSimbolos
' Genera informes basados en gCatalogoSimbolos
'===============================================================

Private Function CatalogoListo() As Boolean
    CatalogoListo = Not (gCatalogoSimbolos Is Nothing)
End Function

Private Function Normalizar(ByVal txt As String) As String
    Normalizar = LCase$(Trim$(txt))
End Function

'---------------------------------------------------------------
' Símbolos NO usados
'---------------------------------------------------------------
Public Function SimbolosNoUsados() As Collection
    Dim col As New Collection
    Dim sim As clsSimbolo

    If Not CatalogoListo() Then Exit Function

    For Each sim In gCatalogoSimbolos.Simbolos
        If Not sim.Usado Then col.Add sim
    Next sim

    Set SimbolosNoUsados = col
End Function

'---------------------------------------------------------------
' Símbolos NO usados por categoría
'---------------------------------------------------------------
Public Function SimbolosNoUsadosPorCategoria(ByVal categoria As String) As Collection
    Dim col As New Collection
    Dim sim As clsSimbolo
    Dim cat As String

    If Not CatalogoListo() Then Exit Function
    cat = Normalizar(categoria)

    For Each sim In gCatalogoSimbolos.Simbolos
        If Not sim.Usado Then
            If Normalizar(sim.categoria) = cat Then
                col.Add sim
            End If
        End If
    Next sim

    Set SimbolosNoUsadosPorCategoria = col
End Function

'---------------------------------------------------------------
' Símbolos NO usados en un módulo concreto
'---------------------------------------------------------------
Public Function SimbolosNoUsadosEnModulo(ByVal nombreModulo As String) As Collection
    Dim col As New Collection
    Dim sim As clsSimbolo
    Dim modNorm As String

    If Not CatalogoListo() Then Exit Function
    modNorm = Normalizar(nombreModulo)

    For Each sim In gCatalogoSimbolos.Simbolos
        If Not sim.Usado Then
            If Normalizar(sim.modulo) = modNorm Then
                col.Add sim
            End If
        End If
    Next sim

    Set SimbolosNoUsadosEnModulo = col
End Function

'---------------------------------------------------------------
' Símbolos usados (informe inverso)
'---------------------------------------------------------------
Public Function SimbolosUsados() As Collection
    Dim col As New Collection
    Dim sim As clsSimbolo

    If Not CatalogoListo() Then Exit Function

    For Each sim In gCatalogoSimbolos.Simbolos
        If sim.Usado Then col.Add sim
    Next sim

    Set SimbolosUsados = col
End Function

'---------------------------------------------------------------
' Símbolos por categoría (usados o no)
'---------------------------------------------------------------
Public Function SimbolosPorCategoria(ByVal categoria As String) As Collection
    Dim col As New Collection
    Dim sim As clsSimbolo
    Dim cat As String

    If Not CatalogoListo() Then Exit Function
    cat = Normalizar(categoria)

    For Each sim In gCatalogoSimbolos.Simbolos
        If Normalizar(sim.categoria) = cat Then col.Add sim
    Next sim

    Set SimbolosPorCategoria = col
End Function

'---------------------------------------------------------------
' Símbolos API no usados
'---------------------------------------------------------------
Public Function APINoUsadas() As Collection
    Dim col As New Collection
    Dim sim As clsSimbolo

    If Not CatalogoListo() Then Exit Function

    For Each sim In gCatalogoSimbolos.Simbolos
        If sim.EsAPI And Not sim.Usado Then
            col.Add sim
        End If
    Next sim

    Set APINoUsadas = col
End Function

'---------------------------------------------------------------
' Estadísticas generales
'---------------------------------------------------------------
Public Function Estadisticas() As Dictionary
    Dim d As Object
    Dim sim As clsSimbolo

    If Not CatalogoListo() Then Exit Function

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare

    d("Total") = gCatalogoSimbolos.Contar
    d("Usados") = 0
    d("NoUsados") = 0

    For Each sim In gCatalogoSimbolos.Simbolos
        If sim.Usado Then
            d("Usados") = d("Usados") + 1
        Else
            d("NoUsados") = d("NoUsados") + 1
        End If
    Next sim

    Set Estadisticas = d
End Function

