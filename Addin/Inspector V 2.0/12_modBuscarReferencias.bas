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

    '-----------------------------------------------------------
    ' Recorrer módulos estándar
    '-----------------------------------------------------------
    For Each m In cat.Modulos
        BuscarReferenciasEnLineas m.lineas, m.nombre
        BuscarReferenciasEnMiembros m.Miembros, m.nombre
    Next m

    '-----------------------------------------------------------
    ' Recorrer clases
    '-----------------------------------------------------------
    For Each c In cat.Clases
        BuscarReferenciasEnLineas c.lineas, c.nombre
        BuscarReferenciasEnMiembros c.Miembros, c.nombre
    Next c

    '-----------------------------------------------------------
    ' Recorrer formularios
    '-----------------------------------------------------------
    For Each frm In cat.UserForms
        BuscarReferenciasEnLineas frm.lineas, frm.nombre
        BuscarReferenciasEnMiembros frm.Miembros, frm.nombre
    Next frm

    '-----------------------------------------------------------
    ' Recorrer informes
    '-----------------------------------------------------------
    For Each inf In cat.Informes
        BuscarReferenciasEnLineas inf.lineas, inf.nombre
        BuscarReferenciasEnMiembros inf.Miembros, inf.nombre
    Next inf

    '-----------------------------------------------------------
    ' Recorrer otros módulos de documento
    '-----------------------------------------------------------
    For Each otro In cat.Otros
        BuscarReferenciasEnLineas otro.lineas, otro.nombre
        BuscarReferenciasEnMiembros otro.Miembros, otro.nombre
    Next otro

End Sub

'---------------------------------------------------------------
' Busca referencias en las líneas de un módulo/clase/formulario
'---------------------------------------------------------------
Private Sub BuscarReferenciasEnLineas(lineas() As String, _
                                      ByVal nombreModulo As String)

    Dim i As Long
    Dim linea As String

    For i = LBound(lineas) To UBound(lineas)
        linea = lineas(i)
        If linea <> "" Then
            MarcarReferenciasEnLinea linea, nombreModulo, ""
        End If
    Next i

End Sub

'---------------------------------------------------------------
' Busca referencias dentro de los miembros
'---------------------------------------------------------------
Private Sub BuscarReferenciasEnMiembros(colMiembros As Collection, _
                                        ByVal nombreModulo As String)

    Dim m As clsMiembro
    Dim i As Long
    Dim linea As String

    For Each m In colMiembros
        For i = m.LineaInicio To m.LineaFin
            linea = m.ObtenerLinea(i)
            If linea <> "" Then
                MarcarReferenciasEnLinea linea, nombreModulo, m.nombre
            End If
        Next i
    Next m

End Sub

'---------------------------------------------------------------
' Marca referencias a símbolos en una línea concreta
'---------------------------------------------------------------
Private Sub MarcarReferenciasEnLinea(ByVal linea As String, _
                                     ByVal nombreModulo As String, _
                                     ByVal nombreMiembro As String)

    Dim sim As clsSimbolo
    Dim nombre As String
    Dim t As String

    ' Normalizar
    t = LCase$(linea)

    ' Recorrer todos los símbolos declarados
    For Each sim In gCatalogoSimbolos.Simbolos

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


