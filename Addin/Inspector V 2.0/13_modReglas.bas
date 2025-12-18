Attribute VB_Name = "13_modReglas"

Option Compare Database
Option Explicit

'===============================================================
' Módulo: 13_modReglas
' Motor moderno de reglas del Inspector VBA
'===============================================================

'---------------------------------------------------------------
' Ejecuta todas las reglas sobre un catálogo analizado
'---------------------------------------------------------------
Public Function EjecutarReglas(cat As clsCatalogoInspector) As Collection
    Dim resultados As New Collection
    Dim procesados As Object
    Set procesados = CreateObject("Scripting.Dictionary")

    AplicarReglasProyecto cat, resultados
    AplicarReglasModulos cat.Modulos, resultados, procesados, "Modulo"
    AplicarReglasClases cat.Clases, resultados, procesados
    AplicarReglasModulos cat.UserForms, resultados, procesados, "UserForm"
    AplicarReglasModulos cat.Formularios, resultados, procesados, "Formulario"
    AplicarReglasModulos cat.Informes, resultados, procesados, "Informe"

    Set EjecutarReglas = resultados
End Function

'---------------------------------------------------------------
' Reglas de proyecto (futuras)
'---------------------------------------------------------------
Private Sub AplicarReglasProyecto(cat As clsCatalogoInspector, resultados As Collection)
    ' (Reservado para reglas globales)
End Sub

'---------------------------------------------------------------
' Reglas de módulos
'---------------------------------------------------------------
Private Sub AplicarReglasModulos(col As Collection, resultados As Collection, _
                                 procesados As Object, tipoContenedor As String)
    Dim m As clsModulo
    For Each m In col
        If Not procesados.Exists(m.nombre) Then
            procesados.Add m.nombre, True
            Regla_OptionExplicit m, resultados
            Regla_ModuloVacio m, resultados
            Regla_MiembrosDuplicados m, resultados, tipoContenedor
        End If
    Next m
End Sub

'---------------------------------------------------------------
' Reglas de clases
'---------------------------------------------------------------
Private Sub AplicarReglasClases(col As Collection, resultados As Collection, procesados As Object)
    Dim c As clsClase
    For Each c In col
        If Not procesados.Exists(c.nombre) Then
            procesados.Add c.nombre, True
            Regla_OptionExplicit_Clase c, resultados
            Regla_ClaseVacia c, resultados
            Regla_MiembrosDuplicados_Clase c, resultados
        End If
    Next c
End Sub

'===============================================================
' REGLA 1: Falta Option Explicit
'===============================================================
Private Sub Regla_OptionExplicit(m As clsModulo, resultados As Collection)
    Dim i As Long, t As String
    If m.numLineas = 0 Then Exit Sub

    For i = 0 To Application.Min(10, m.numLineas - 1)
        t = Trim$(LCase$(m.lineas(i)))
        If t = "option explicit" Then Exit Sub
        If t <> "" And Left$(t, 1) <> "'" Then Exit For
    Next i

    resultados.Add CrearResultado( _
        sevError, teModulo, m.nombre, "", tmUnknown, _
        "Falta Option Explicit en el módulo.", "R001", "ADD_OPTION_EXPLICIT", True, 1)
End Sub

Private Sub Regla_OptionExplicit_Clase(c As clsClase, resultados As Collection)
    Dim i As Long, t As String
    If c.numLineas = 0 Then Exit Sub

    For i = 0 To Application.Min(10, c.numLineas - 1)
        t = Trim$(LCase$(c.lineas(i)))
        If t = "option explicit" Then Exit Sub
        If t <> "" And Left$(t, 1) <> "'" Then Exit For
    Next i

    resultados.Add CrearResultado( _
        sevError, teClase, c.nombre, "", tmUnknown, _
        "Falta Option Explicit en la clase.", "R001", "ADD_OPTION_EXPLICIT", True, 1)
End Sub

'===============================================================
' REGLA 2: Módulo o clase vacía
'===============================================================
Private Sub Regla_ModuloVacio(m As clsModulo, resultados As Collection)
    If m.NumLineasCodigo = 0 And m.NumLineasComentario = 0 And m.NumLineasAtributo = 0 Then
        resultados.Add CrearResultado( _
            sevAviso, teModulo, m.nombre, "", tmUnknown, _
            "El módulo está vacío.", "R002", "", False, 0)
    End If
End Sub

Private Sub Regla_ClaseVacia(c As clsClase, resultados As Collection)
    If c.NumLineasCodigo = 0 And c.NumLineasComentario = 0 And c.NumLineasAtributo = 0 Then
        resultados.Add CrearResultado( _
            sevAviso, teClase, c.nombre, "", tmUnknown, _
            "La clase está vacía.", "R002", "", False, 0)
    End If
End Sub

'===============================================================
' REGLA 3: Miembros duplicados
'===============================================================
Private Sub Regla_MiembrosDuplicados(m As clsModulo, resultados As Collection, tipoContenedor As String)
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim mi As clsMiembro, clave As String, desc As String

    For Each mi In m.Miembros
        clave = LCase$(mi.nombre)
        If dict.Exists(clave) Then
            desc = "Miembro duplicado en " & _
                   IIf(tipoContenedor = "UserForm", "un UserForm.", _
                   IIf(tipoContenedor = "Formulario", "un formulario de Access.", _
                   IIf(tipoContenedor = "Informe", "un informe de Access.", "el módulo.")))

            resultados.Add CrearResultado( _
                sevError, teMiembro, m.nombre, mi.nombre, mi.tipo, _
                desc, "R003", "", False, mi.LineaInicio)
        Else
            dict.Add clave, True
        End If
    Next mi
End Sub

Private Sub Regla_MiembrosDuplicados_Clase(c As clsClase, resultados As Collection)
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim mi As clsMiembro

    For Each mi In c.Miembros
        If dict.Exists(LCase$(mi.nombre)) Then
            resultados.Add CrearResultado( _
                sevError, teMiembro, c.nombre, mi.nombre, mi.tipo, _
                "Miembro duplicado en la clase.", "R003", "", False, mi.LineaInicio)
        Else
            dict.Add LCase$(mi.nombre), True
        End If
    Next mi
End Sub

'===============================================================
' Constructor centralizado de resultados
'===============================================================
Private Function CrearResultado( _
    sev As SeveridadInspector, _
    tipoElemento As TipoElementoInspector, _
    nombreElemento As String, _
    nombreMiembro As String, _
    tipoMiembro As TipoMiembroInspector, _
    descripcion As String, _
    codigoRegla As String, _
    codigoReparacion As String, _
    esReparable As Boolean, _
    linea As Long _
) As clsResultadoAnalisis

    Dim res As New clsResultadoAnalisis
    res.severidad = sev
    res.tipoElemento = tipoElemento
    res.nombreElemento = nombreElemento
    res.nombreMiembro = nombreMiembro
    res.tipoMiembro = tipoMiembro
    res.descripcion = descripcion
    res.codigoRegla = codigoRegla
    res.codigoReparacion = codigoReparacion
    res.esReparable = esReparable
    res.linea = linea

    Set CrearResultado = res
End Function


