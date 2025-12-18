Attribute VB_Name = "12_modBuscarReferencias"

'===============================================================
' Módulo: 12_modBuscarReferencias
' Detección de referencias a símbolos declarados
'===============================================================

Option Compare Database
Option Explicit

'===============================================================
' Módulo: 12_modBuscarReferencias
' Detección segura de referencias a símbolos declarados
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
    Dim tokens As Collection
    Dim tok As Variant
    Dim colRes As Collection
    Dim info As ReservedWordInfo

    ' 1. Limpiar comentarios y cadenas
    linea = LimpiarLinea(linea)
    If linea = "" Then Exit Sub

    ' 2. Tokenizar
    Set tokens = Tokenizar(linea)
    If tokens.Count = 0 Then Exit Sub

    ' ----------------------------------------------------------
    ' 3. Detectar palabras reservadas (NUEVO)
    ' ----------------------------------------------------------
    For Each tok In tokens

        Set colRes = M13_PalabrasReservadas.FindReservedWord(tok)

        If colRes.Count > 0 Then
            ' Registrar cada coincidencia
            For Each info In colRes
                cat.RegistrarPalabraReservada tok, info, nombreModulo, nombreMiembro
            Next info
        End If

    Next tok

    ' ----------------------------------------------------------
    ' 4. Comparar tokens con símbolos declarados (ya existente)
    ' ----------------------------------------------------------
    For Each sim In cat.CatalogoSimbolos.Simbolos

        ' Evitar marcar la declaración como uso
        If sim.modulo = nombreModulo Then
            If sim.miembro = nombreMiembro Then
                If sim.LineaDeclaracion > 0 Then
                    GoTo SiguienteSimbolo
                End If
            End If
        End If

        ' Buscar coincidencia exacta en tokens
        If TokenExiste(tokens, sim.nombre) Then
            sim.Usado = True
        End If

SiguienteSimbolo:
    Next sim

End Sub

'Private Sub MarcarReferenciasEnLinea(ByVal linea As String, _
'                                     ByVal nombreModulo As String, _
'                                     ByVal nombreMiembro As String, _
'                                     cat As clsCatalogoInspector)
'
'    Dim sim As clsSimbolo
'    Dim tokens As Collection
'    Dim tok As Variant
'
'    ' 1. Limpiar comentarios y cadenas
'    linea = LimpiarLinea(linea)
'    If linea = "" Then Exit Sub
'
'    ' 2. Tokenizar
'    Set tokens = Tokenizar(linea)
'    If tokens.Count = 0 Then Exit Sub
'
'    ' 3. Comparar tokens con símbolos
'    For Each sim In cat.CatalogoSimbolos.Simbolos
'
'        ' Evitar marcar la declaración como uso
'        If sim.modulo = nombreModulo Then
'            If sim.miembro = nombreMiembro Then
'                ' Si estamos en la línea de declaración, ignorar
'                If sim.LineaDeclaracion > 0 Then
'                    ' No marcar uso en la misma línea
'                    GoTo SiguienteSimbolo
'                End If
'            End If
'        End If
'
'        ' Buscar coincidencia exacta en tokens
'        If TokenExiste(tokens, sim.nombre) Then
'            sim.Usado = True
'        End If
'
'SiguienteSimbolo:
'    Next sim
'
'End Sub

'---------------------------------------------------------------
' Limpia comentarios y cadenas de una línea
'---------------------------------------------------------------
Private Function LimpiarLinea(ByVal linea As String) As String
    Dim i As Long
    Dim enCadena As Boolean
    Dim resultado As String
    Dim ch As String

    For i = 1 To Len(linea)
        ch = Mid$(linea, i, 1)

        ' Detectar inicio/fin de cadena
        If ch = """" Then
            enCadena = Not enCadena
            GoTo Siguiente
        End If

        ' Ignorar contenido dentro de cadenas
        If enCadena Then GoTo Siguiente

        ' Ignorar comentarios
        If ch = "'" Then Exit For

        resultado = resultado & ch

Siguiente:
    Next i

    LimpiarLinea = Trim$(resultado)
End Function

'---------------------------------------------------------------
' Tokeniza una línea en identificadores válidos
'---------------------------------------------------------------
Private Function Tokenizar(ByVal linea As String) As Collection
    Dim col As New Collection
    Dim tmp As String
    Dim parts() As String
    Dim p As Variant

    ' Reemplazar delimitadores comunes
    tmp = linea
    tmp = Replace(tmp, "(", " ")
    tmp = Replace(tmp, ")", " ")
    tmp = Replace(tmp, ",", " ")
    tmp = Replace(tmp, ".", " ")
    tmp = Replace(tmp, "=", " ")
    tmp = Replace(tmp, "+", " ")
    tmp = Replace(tmp, "-", " ")
    tmp = Replace(tmp, "*", " ")
    tmp = Replace(tmp, "/", " ")

    parts = Split(tmp, " ")

    For Each p In parts
        p = Trim$(p)
        If EsIdentificadorValido(p) Then
            col.Add p
        End If
    Next p

    Set Tokenizar = col
End Function

'---------------------------------------------------------------
' Determina si un token es un identificador válido
'---------------------------------------------------------------
Private Function EsIdentificadorValido(ByVal t As String) As Boolean
    If t = "" Then Exit Function

    ' No palabras reservadas
    Select Case LCase$(t)
        Case "if", "then", "else", "end", "sub", "function", _
             "property", "let", "set", "get", "public", "private", _
             "static", "dim", "as", "select", "case", "loop", _
             "for", "next", "while", "wend", "do", "until", _
             "enum", "type", "const", "option", "explicit"
            Exit Function
    End Select

    ' Debe empezar por letra o guion bajo
    If Not (t Like "[A-Za-z_]*") Then Exit Function

    EsIdentificadorValido = True
End Function

'---------------------------------------------------------------
' Determina si un identificador aparece en los tokens
'---------------------------------------------------------------
Private Function TokenExiste(tokens As Collection, ByVal nombre As String) As Boolean
    Dim t As Variant

    For Each t In tokens
        If StrComp(t, nombre, vbTextCompare) = 0 Then
            TokenExiste = True
            Exit Function
        End If
    Next t

    TokenExiste = False
End Function


