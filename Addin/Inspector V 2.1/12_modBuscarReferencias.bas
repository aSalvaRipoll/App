Attribute VB_Name = "12_modBuscarReferencias"

Option Compare Database
Option Explicit

'===============================================================
' Módulo: 12_modBuscarReferencias
' Detección de referencias a símbolos declarados
'===============================================================

'---------------------------------------------------------------
' Punto de entrada: recorre todo el proyecto y marca referencias
'---------------------------------------------------------------
Public Sub BuscarReferenciasEnProyecto(cat As clsCatalogoInspector)

    Dim m As clsModulo
    Dim c As clsClase
    Dim frm As clsModulo
    Dim inf As clsModulo
    Dim otro As clsModulo

    ' Módulos estándar
    For Each m In cat.Modulos
        BuscarReferenciasEnLineas m.lineas, m.nombre, cat
        BuscarReferenciasEnMiembros m.Miembros, m.nombre, cat
    Next m

    ' Clases
    For Each c In cat.Clases
        BuscarReferenciasEnLineas c.lineas, c.nombre, cat
        BuscarReferenciasEnMiembros c.Miembros, c.nombre, cat
    Next c

    ' Formularios
    For Each frm In cat.UserForms
        BuscarReferenciasEnLineas frm.lineas, frm.nombre, cat
        BuscarReferenciasEnMiembros frm.Miembros, frm.nombre, cat
    Next frm

    ' Informes
    For Each inf In cat.Informes
        BuscarReferenciasEnLineas inf.lineas, inf.nombre, cat
        BuscarReferenciasEnMiembros inf.Miembros, inf.nombre, cat
    Next inf

    ' Otros
    For Each otro In cat.Otros
        BuscarReferenciasEnLineas otro.lineas, otro.nombre, cat
        BuscarReferenciasEnMiembros otro.Miembros, otro.nombre, cat
    Next otro

End Sub

'---------------------------------------------------------------
' Busca referencias en las líneas de un módulo/clase/formulario
'---------------------------------------------------------------
Private Sub BuscarReferenciasEnLineas(lineas() As String, _
                                      ByVal nombreModulo As String, _
                                      cat As clsCatalogoInspector)

    Dim i As Long
    Dim linea As String

    For i = LBound(lineas) To UBound(lineas)
        linea = lineas(i)
        If linea <> "" Then
            MarcarReferenciasEnLinea linea, nombreModulo, "", cat
        End If
    Next i

End Sub

'---------------------------------------------------------------
' Busca referencias dentro de los miembros
'---------------------------------------------------------------
Private Sub BuscarReferenciasEnMiembros(colMiembros As Collection, _
                                        ByVal nombreModulo As String, _
                                        cat As clsCatalogoInspector)

    Dim m As clsMiembro
    Dim i As Long
    Dim linea As String

    For Each m In colMiembros
        For i = m.LineaInicio To m.LineaFin
            linea = m.ObtenerLinea(i)
            If linea <> "" Then
                MarcarReferenciasEnLinea linea, nombreModulo, m.nombre, cat
            End If
        Next i
    Next m

End Sub


'---------------------------------------------------------------
' Marca referencias a símbolos en una línea concreta
'---------------------------------------------------------------
Private Sub MarcarReferenciasEnLinea(ByVal linea As String, _
                                     ByVal nombreModulo As String, _
                                     ByVal nombreMiembro As String, _
                                     cat As clsCatalogoInspector)

    Dim sim As clsSimbolo
    Dim nombre As String
    Dim t As String

    t = LCase$(linea)

    For Each sim In cat.CatalogoSimbolos.Simbolos

        nombre = LCase$(sim.nombre)

        ' Evitar marcar la declaración como uso
        If sim.modulo = nombreModulo Then
            If sim.miembro = nombreMiembro Then
                If sim.LineaDeclaracion > 0 Then
                    If InStr(1, linea, sim.nombre, vbTextCompare) > 0 _
                    And sim.LineaDeclaracion = 0 Then
                        ' nada
                    End If
                End If
            End If
        End If

        ' Buscar referencia exacta
        If PalabraEnLinea(t, nombre) Then
            sim.Usado = True
        End If

    Next sim

End Sub


'---------------------------------------------------------------
' Determina si una palabra aparece como token independiente
'---------------------------------------------------------------
Private Function PalabraEnLinea(ByVal linea As String, _
                                ByVal palabra As String) As Boolean

    Dim tokens() As String
    Dim t As Variant

    ' Separar por delimitadores comunes
    tokens = Split(Replace(Replace(Replace(linea, "(", " "), ")", " "), ",", " "), " ")

    For Each t In tokens
        If Trim$(t) = palabra Then
            PalabraEnLinea = True
            Exit Function
        End If
    Next t

    PalabraEnLinea = False
End Function


