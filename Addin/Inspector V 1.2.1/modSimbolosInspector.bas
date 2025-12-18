Attribute VB_Name = "modSimbolosInspector"

Option Compare Database
Option Explicit

'=====================================================
' Módulo: modSimbolosInspector
' Entrada al índice global de símbolos
'=====================================================

Public gCatalogoSimbolos As clsCatalogoSimbolos

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

