Attribute VB_Name = "15_modSimbolos"

Option Compare Database
Option Explicit

'=====================================================
' Módulo: 15_modSimbolos
' Entrada al índice global de símbolos
'=====================================================

'-----------------------------------------------------
' Inicializa el catálogo global de símbolos (si no existe)
'-----------------------------------------------------
Public Sub InicializarCatalogoSimbolos(Optional ForzarReinicio As Boolean = False)

    ' Si ya está inicializado y no se pide reinicio, no hacer nada
    If Not gCatalogoSimbolos Is Nothing Then
        If Not ForzarReinicio Then Exit Sub
    End If

    On Error GoTo ErrHandler
    Set gCatalogoSimbolos = New clsCatalogoSimbolos

    Debug.Print "Catálogo global de símbolos inicializado."
    Exit Sub

ErrHandler:
    Debug.Print "Error al inicializar el catálogo de símbolos: "; Err.Description
End Sub

'-----------------------------------------------------
' Garantiza que el catálogo está inicializado
'-----------------------------------------------------
Public Sub AsegurarCatalogoSimbolos()
    If gCatalogoSimbolos Is Nothing Then
        InicializarCatalogoSimbolos
    End If
End Sub

'===============================================================
' Cálculo de estadísticas sobre el catálogo de símbolos
'===============================================================
Public Function CalcularEstadisticas(cat As clsCatalogoSimbolos) As clsEstadisticas
    Dim est As New clsEstadisticas
    Dim s As clsSimbolo

    est.Total = cat.Simbolos.Count

    For Each s In cat.Simbolos
        If s.Usado Then
            est.Usados = est.Usados + 1
        Else
            est.NoUsados = est.NoUsados + 1
        End If
    Next s

    Set CalcularEstadisticas = est
End Function

