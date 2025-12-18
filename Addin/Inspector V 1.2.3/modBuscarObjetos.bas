Attribute VB_Name = "modBuscarObjetos"

Option Compare Database
Option Explicit

'===============================================================
' Módulo: modBuscarObjetos
' Detección de declaraciones en módulos VBA
' (variables, constantes, enums, UDTs)
'===============================================================

'---------------------------------------------------------------
' Detecta declaraciones dentro de un módulo o clase
'---------------------------------------------------------------
Public Sub DetectarDeclaraciones(code As VBIDE.CodeModule, _
                                 ByVal nombreModulo As String, _
                                 ByVal nombreMiembro As String)

    Dim i As Long
    Dim linea As String
    Dim t As String
    Dim sim As clsSimbolo

    For i = 1 To code.CountOfLines
        linea = Trim$(code.Lines(i, 1))
        If linea = "" Then GoTo Siguiente

        t = LCase$(linea)

        '=======================================================
        ' CONST
        '=======================================================
        If Left$(t, 5) = "const" Then
            Set sim = New clsSimbolo
            sim.nombre = ExtraerNombreDeclaracion(linea)
            sim.categoria = "Constante"
            sim.Ambito = DeterminarAmbito(linea)
            sim.modulo = nombreModulo
            sim.miembro = nombreMiembro
            sim.LineaDeclaracion = i
            gCatalogoSimbolos.Agregar sim
            GoTo Siguiente
        End If

        '=======================================================
        ' VARIABLES (Dim, Private, Public, Static)
        '=======================================================
        If Left$(t, 3) = "dim" _
        Or Left$(t, 7) = "private" _
        Or Left$(t, 6) = "public" _
        Or Left$(t, 6) = "static" Then

            Set sim = New clsSimbolo
            sim.nombre = ExtraerNombreDeclaracion(linea)
            sim.categoria = "Variable"
            sim.Ambito = DeterminarAmbito(linea)
            sim.modulo = nombreModulo
            sim.miembro = nombreMiembro
            sim.LineaDeclaracion = i
            gCatalogoSimbolos.Agregar sim
            GoTo Siguiente
        End If

        '=======================================================
        ' ENUM
        '=======================================================
        If Left$(t, 4) = "enum" Then
            Set sim = New clsSimbolo
            sim.nombre = ExtraerNombreDeclaracion(linea)
            sim.categoria = "Enum"
            sim.Ambito = DeterminarAmbito(linea)
            sim.modulo = nombreModulo
            sim.LineaDeclaracion = i
            gCatalogoSimbolos.Agregar sim
            GoTo Siguiente
        End If

        '=======================================================
        ' UDT (Type ... End Type)
        '=======================================================
        If Left$(t, 4) = "type" Then
            Set sim = New clsSimbolo
            sim.nombre = ExtraerNombreDeclaracion(linea)
            sim.categoria = "UDT"
            sim.Ambito = DeterminarAmbito(linea)
            sim.modulo = nombreModulo
            sim.LineaDeclaracion = i
            gCatalogoSimbolos.Agregar sim
            GoTo Siguiente
        End If

Siguiente:
    Next i
End Sub

'---------------------------------------------------------------
' Extrae el nombre de una declaración (Dim, Const, Enum, Type)
'---------------------------------------------------------------
Public Function ExtraerNombreDeclaracion(ByVal linea As String) As String
    Dim tmp As String
    Dim parts() As String

    tmp = Replace(linea, "As ", " As ")
    tmp = Replace(tmp, "(", " ")
    tmp = Replace(tmp, ")", " ")
    tmp = Trim$(tmp)

    parts = Split(tmp, " ")

    ' Ejemplos:
    ' Dim x As Long        ? parts(1)
    ' Private y As String  ? parts(1)
    ' Public z             ? parts(1)
    ' Const IVA = 21       ? parts(1)
    ' Enum TipoCliente     ? parts(1)
    ' Type DatosCliente    ? parts(1)

    If UBound(parts) >= 1 Then
        ExtraerNombreDeclaracion = parts(1)
    Else
        ExtraerNombreDeclaracion = parts(0)
    End If
End Function

'---------------------------------------------------------------
' Determina el ámbito de una declaración
'---------------------------------------------------------------
Public Function DeterminarAmbito(ByVal linea As String) As String
    Dim t As String
    t = LCase$(Trim$(linea))

    If Left$(t, 6) = "public" Then
        DeterminarAmbito = "ModuloPublic"
    ElseIf Left$(t, 7) = "private" Then
        DeterminarAmbito = "ModuloPrivate"
    ElseIf Left$(t, 6) = "static" Then
        DeterminarAmbito = "Local"
    ElseIf Left$(t, 3) = "dim" Then
        DeterminarAmbito = "Local"
    Else
        DeterminarAmbito = "Desconocido"
    End If
End Function

'---------------------------------------------------------------
' Detecta declaraciones locales dentro de un miembro
'---------------------------------------------------------------
Public Sub DetectarDeclaracionesLocales(code As VBIDE.CodeModule, _
                                        ByVal miembro As clsMiembro, _
                                        ByVal nombreModulo As String)

    Dim i As Long
    Dim linea As String
    Dim t As String
    Dim sim As clsSimbolo

    For i = miembro.LineaInicio To miembro.LineaFin
        linea = Trim$(code.Lines(i, 1))
        If linea = "" Then GoTo Siguiente

        t = LCase$(linea)

        '=======================================================
        ' Constantes locales
        '=======================================================
        If Left$(t, 5) = "const" Then
            Set sim = New clsSimbolo
            sim.nombre = ExtraerNombreDeclaracion(linea)
            sim.categoria = "ConstanteLocal"
            sim.Ambito = "Local"
            sim.modulo = nombreModulo
            sim.miembro = miembro.nombre
            sim.LineaDeclaracion = i
            gCatalogoSimbolos.Agregar sim
            GoTo Siguiente
        End If

        '=======================================================
        ' Variables locales (Dim, Static)
        '=======================================================
        If Left$(t, 3) = "dim" _
        Or Left$(t, 6) = "static" Then

            Set sim = New clsSimbolo
            sim.nombre = ExtraerNombreDeclaracion(linea)
            sim.categoria = "VariableLocal"
            sim.Ambito = "Local"
            sim.modulo = nombreModulo
            sim.miembro = miembro.nombre
            sim.LineaDeclaracion = i
            gCatalogoSimbolos.Agregar sim
            GoTo Siguiente
        End If

Siguiente:
    Next i
End Sub
