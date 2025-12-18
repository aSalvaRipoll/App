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


'===============================================================
'=== NUEVO: Integración con palabras reservadas
'===============================================================

'-----------------------------------------------------
' Devuelve True si el nombre es una palabra reservada
'-----------------------------------------------------
Public Function EsPalabraReservada(ByVal nombre As String) As Boolean
    If gPalabrasReservadas Is Nothing Then Exit Function
    EsPalabraReservada = gPalabrasReservadas.Exists(LCase$(Trim$(nombre)))
End Function

'-----------------------------------------------------
' Regla R-PR-001: Nombre coincide con palabra reservada
'-----------------------------------------------------
Public Sub AnalizarNombresReservados()

    AsegurarCatalogoSimbolos
    If gCatalogoSimbolos Is Nothing Then Exit Sub

    Dim sim As clsSimbolo
    Dim res As clsResultadoAnalisis

    For Each sim In gCatalogoSimbolos.Simbolos

        If EsPalabraReservada(sim.nombre) Then

            Set res = New clsResultadoAnalisis
            res.severidad = sevAviso
            res.tipoElemento = teMiembro
            res.nombreElemento = sim.modulo
            res.nombreMiembro = sim.miembro
            res.linea = sim.LineaDeclaracion
            res.descripcion = "El nombre '" & sim.nombre & "' coincide con una palabra reservada."
            res.codigoRegla = "R-PR-001"
            res.esReparable = False

            gResultadosInspector.AgregarResultado res
        End If

    Next sim

End Sub




'Option Compare Database
'Option Explicit
'
''=====================================================
'' Módulo: 15_modSimbolos
'' Entrada al índice global de símbolos
''=====================================================
'
''-----------------------------------------------------
'' Inicializa el catálogo global de símbolos (si no existe)
''-----------------------------------------------------
'Public Sub InicializarCatalogoSimbolos(Optional ForzarReinicio As Boolean = False)
'
'    ' Si ya está inicializado y no se pide reinicio, no hacer nada
'    If Not gCatalogoSimbolos Is Nothing Then
'        If Not ForzarReinicio Then Exit Sub
'    End If
'
'    On Error GoTo ErrHandler
'    Set gCatalogoSimbolos = New clsCatalogoSimbolos
'
'    Debug.Print "Catálogo global de símbolos inicializado."
'    Exit Sub
'
'ErrHandler:
'    Debug.Print "Error al inicializar el catálogo de símbolos: "; Err.Description
'End Sub
'
''-----------------------------------------------------
'' Garantiza que el catálogo está inicializado
''-----------------------------------------------------
'Public Sub AsegurarCatalogoSimbolos()
'    If gCatalogoSimbolos Is Nothing Then
'        InicializarCatalogoSimbolos
'    End If
'End Sub
'
''===============================================================
'' Cálculo de estadísticas sobre el catálogo de símbolos
''===============================================================
'Public Function CalcularEstadisticas(cat As clsCatalogoSimbolos) As clsEstadisticas
'    Dim est As New clsEstadisticas
'    Dim s As clsSimbolo
'
'    est.Total = cat.Simbolos.Count
'
'    For Each s In cat.Simbolos
'        If s.Usado Then
'            est.Usados = est.Usados + 1
'        Else
'            est.NoUsados = est.NoUsados + 1
'        End If
'    Next s
'
'    Set CalcularEstadisticas = est
'End Function

