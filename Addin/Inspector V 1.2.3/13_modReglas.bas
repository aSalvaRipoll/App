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

    ' Evitar procesar dos veces el mismo módulo
    Set procesados = CreateObject("Scripting.Dictionary")

    ' Reglas de proyecto
    AplicarReglasProyecto cat, resultados

    ' Reglas de módulos estándar
    AplicarReglasModulos cat.Modulos, resultados, procesados, "Modulo"

    ' Reglas de clases
    AplicarReglasClases cat.Clases, resultados, procesados

    ' Reglas de UserForms (VBA)
    AplicarReglasModulos cat.UserForms, resultados, procesados, "UserForm"

    ' Reglas de formularios de Access
    AplicarReglasModulos cat.Formularios, resultados, procesados, "Formulario"

    ' Reglas de informes de Access
    AplicarReglasModulos cat.Informes, resultados, procesados, "Informe"

    Set EjecutarReglas = resultados
End Function

'===============================================================
' REGLAS DE PROYECTO
'===============================================================
Private Sub AplicarReglasProyecto(cat As clsCatalogoInspector, resultados As Collection)
    ' Reglas globales del proyecto (futuras)
End Sub

'===============================================================
' REGLAS DE MÓDULOS
'===============================================================
Private Sub AplicarReglasModulos(col As Collection, resultados As Collection, _
                                 procesados As Object, tipoContenedor As String)

    Dim m As clsModulo

    For Each m In col

        ' Evitar procesar dos veces el mismo módulo
        If Not procesados.Exists(m.nombre) Then
            procesados.Add m.nombre, True

            Regla_OptionExplicit m, resultados
            Regla_ModuloVacio m, resultados
            Regla_MiembrosDuplicados m, resultados, tipoContenedor
        End If

    Next m
End Sub

'===============================================================
' REGLAS DE CLASES
'===============================================================
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
' REGLA 1: Falta Option Explicit (inteligente)
'===============================================================
Private Sub Regla_OptionExplicit(m As clsModulo, resultados As Collection)
    Dim r As clsResultadoAnalisis
    Dim i As Long
    Dim linea As String

    If m.numLineas = 0 Then Exit Sub

    ' Buscar Option Explicit en las primeras líneas significativas
    For i = 0 To Application.Min(10, m.numLineas - 1)
        linea = Trim$(LCase$(m.lineas(i)))

        If linea = "option explicit" Then Exit Sub
        If Left$(linea, 1) <> "'" And linea <> "" Then Exit For
    Next i

    ' Si llegamos aquí, falta Option Explicit
    Set r = CrearResultado( _
        sevError, _
        teModulo, _
        m.nombre, _
        "", _
        tmUnknown, _
        "Falta Option Explicit en el módulo.", _
        "R001", _
        "ADD_OPTION_EXPLICIT", _
        True, _
        1 _
    )
    resultados.Add r
End Sub

Private Sub Regla_OptionExplicit_Clase(c As clsClase, resultados As Collection)
    Dim r As clsResultadoAnalisis
    Dim i As Long
    Dim linea As String

    If c.numLineas = 0 Then Exit Sub

    For i = 0 To Application.Min(10, c.numLineas - 1)
        linea = Trim$(LCase$(c.lineas(i)))

        If linea = "option explicit" Then Exit Sub
        If Left$(linea, 1) <> "'" And linea <> "" Then Exit For
    Next i

    Set r = CrearResultado( _
        sevError, _
        teClase, _
        c.nombre, _
        "", _
        tmUnknown, _
        "Falta Option Explicit en la clase.", _
        "R001", _
        "ADD_OPTION_EXPLICIT", _
        True, _
        1 _
    )
    resultados.Add r
End Sub

'===============================================================
' REGLA 2: Módulo vacío (inteligente)
'===============================================================
Private Sub Regla_ModuloVacio(m As clsModulo, resultados As Collection)
    Dim r As clsResultadoAnalisis

    ' Considerar atributos como contenido
    If m.NumLineasCodigo = 0 And m.NumLineasComentario = 0 And m.NumLineasAtributo = 0 Then
        Set r = CrearResultado( _
            sevAviso, _
            teModulo, _
            m.nombre, _
            "", _
            tmUnknown, _
            "El módulo está vacío.", _
            "R002", _
            "", _
            False, _
            0 _
        )
        resultados.Add r
    End If
End Sub

Private Sub Regla_ClaseVacia(c As clsClase, resultados As Collection)
    Dim r As clsResultadoAnalisis

    If c.NumLineasCodigo = 0 And c.NumLineasComentario = 0 And c.NumLineasAtributo = 0 Then
        Set r = CrearResultado( _
            sevAviso, _
            teClase, _
            c.nombre, _
            "", _
            tmUnknown, _
            "La clase está vacía.", _
            "R002", _
            "", _
            False, _
            0 _
        )
        resultados.Add r
    End If
End Sub

'===============================================================
' REGLA 3: Miembros duplicados (distinción por tipo)
'===============================================================
Private Sub Regla_MiembrosDuplicados(m As clsModulo, resultados As Collection, tipoContenedor As String)
    Dim dict As Object
    Dim mi As clsMiembro
    Dim r As clsResultadoAnalisis
    Dim clave As String

    Set dict = CreateObject("Scripting.Dictionary")

    For Each mi In m.Miembros
        clave = LCase$(mi.nombre)

        If dict.Exists(clave) Then

            Dim descripcion As String
            descripcion = "Miembro duplicado en "

            Select Case tipoContenedor
                Case "UserForm": descripcion = descripcion & "un UserForm."
                Case "Formulario": descripcion = descripcion & "un formulario de Access."
                Case "Informe": descripcion = descripcion & "un informe de Access."
                Case Else: descripcion = descripcion & "el módulo."
            End Select

            Set r = CrearResultado( _
                sevError, _
                teMiembro, _
                m.nombre, _
                mi.nombre, _
                mi.tipo, _
                descripcion, _
                "R003", _
                "", _
                False, _
                mi.LineaInicio _
            )
            resultados.Add r

        Else
            dict.Add clave, True
        End If
    Next mi
End Sub

Private Sub Regla_MiembrosDuplicados_Clase(c As clsClase, resultados As Collection)
    Dim dict As Object
    Dim mi As clsMiembro
    Dim r As clsResultadoAnalisis

    Set dict = CreateObject("Scripting.Dictionary")

    For Each mi In c.Miembros
'Option Compare Database
'Option Explicit
'
''===============================================================
'' Módulo: modReglasInspector
'' Motor moderno de reglas del Inspector VBA
''===============================================================
'
''---------------------------------------------------------------
'' Ejecuta todas las reglas sobre un catálogo analizado
''---------------------------------------------------------------
'Public Function EjecutarReglas(cat As clsCatalogoInspector) As Collection
'    Dim resultados As New Collection
'    Dim procesados As Object
'
'    ' Evitar procesar dos veces el mismo módulo
'    Set procesados = CreateObject("Scripting.Dictionary")
'
'    ' Reglas de proyecto
'    AplicarReglasProyecto cat, resultados
'
'    ' Reglas de módulos estándar
'    AplicarReglasModulos cat.Modulos, resultados, procesados, "Modulo"
'
'    ' Reglas de clases
'    AplicarReglasClases cat.Clases, resultados, procesados
'
'    ' Reglas de UserForms (VBA)
'    AplicarReglasModulos cat.UserForms, resultados, procesados, "UserForm"
'
'    ' Reglas de formularios de Access
'    AplicarReglasModulos cat.Formularios, resultados, procesados, "Formulario"
'
'    ' Reglas de informes de Access
'    AplicarReglasModulos cat.Informes, resultados, procesados, "Informe"
'
'    Set EjecutarReglas = resultados
'End Function
'
''===============================================================
'' REGLAS DE PROYECTO
''===============================================================
'Private Sub AplicarReglasProyecto(cat As clsCatalogoInspector, resultados As Collection)
'    ' Reglas globales del proyecto (futuras)
'End Sub
'
''===============================================================
'' REGLAS DE MÓDULOS
''===============================================================
'Private Sub AplicarReglasModulos(col As Collection, resultados As Collection, _
'                                 procesados As Object, tipoContenedor As String)
'
'    Dim m As clsModulo
'
'    For Each m In col
'
'        ' Evitar procesar dos veces el mismo módulo
'        If Not procesados.Exists(m.nombre) Then
'            procesados.Add m.nombre, True
'
'            Regla_OptionExplicit m, resultados
'            Regla_ModuloVacio m, resultados
'            Regla_MiembrosDuplicados m, resultados, tipoContenedor
'        End If
'
'    Next m
'End Sub
'
''===============================================================
'' REGLAS DE CLASES
''===============================================================
'Private Sub AplicarReglasClases(col As Collection, resultados As Collection, procesados As Object)
'    Dim c As clsClase
'
'    For Each c In col
'
'        If Not procesados.Exists(c.nombre) Then
'            procesados.Add c.nombre, True
'
'            Regla_OptionExplicit_Clase c, resultados
'            Regla_ClaseVacia c, resultados
'            Regla_MiembrosDuplicados_Clase c, resultados
'        End If
'
'    Next c
'End Sub
'
''===============================================================
'' REGLA 1: Falta Option Explicit (inteligente)
''===============================================================
'
'Private Sub Regla_OptionExplicit(m As clsModulo, resultados As Collection)
'    Dim r As clsResultadoAnalisis
'    Dim i As Long
'    Dim linea As String
'
'    If m.numLineas = 0 Then Exit Sub
'
'    ' Buscar Option Explicit en las primeras líneas significativas
'    For i = 0 To Application.Min(10, m.numLineas - 1)
'        linea = Trim$(LCase$(m.lineas(i)))
'
'        If linea = "option explicit" Then Exit Sub
'        If Left$(linea, 1) <> "'" And linea <> "" Then Exit For
'    Next i
'
'    ' Si llegamos aquí, falta Option Explicit
'    Set r = CrearResultado( _
'        sevError, _
'        teModulo, _
'        m.nombre, _
'        "", _
'        tmUnknown, _
'        "Falta Option Explicit en el módulo.", _
'        "R001", _
'        "ADD_OPTION_EXPLICIT", _
'        True, _
'        1 _
'    )
''    Set r = CrearResultado( _
''        sevError, _
''        teModulo, _
''        m.nombre, _
''        "", _
''        "Falta Option Explicit en el módulo.", _
''        "R001", _
''        "ADD_OPTION_EXPLICIT", _
''        True, _
''        1 _
''    )
'    resultados.Add r
'End Sub
'
'Private Sub Regla_OptionExplicit_Clase(c As clsClase, resultados As Collection)
'    Dim r As clsResultadoAnalisis
'    Dim i As Long
'    Dim linea As String
'
'    If c.numLineas = 0 Then Exit Sub
'
'    For i = 0 To Application.Min(10, c.numLineas - 1)
'        linea = Trim$(LCase$(c.lineas(i)))
'
'        If linea = "option explicit" Then Exit Sub
'        If Left$(linea, 1) <> "'" And linea <> "" Then Exit For
'    Next i
'
'    Set r = CrearResultado( _
'        sevError, _
'        teClase, _
'        c.nombre, _
'        "", _
'        tmUnknown, _
'        "Falta Option Explicit en la clase.", _
'        "R001", _
'        "ADD_OPTION_EXPLICIT", _
'        True, _
'        1 _
'    )
'
''    Set r = CrearResultado( _
''        sevError, _
''        teClase, _
''        c.nombre, _
''        "", _
''        "Falta Option Explicit en la clase.", _
''        "R001", _
''        "ADD_OPTION_EXPLICIT", _
''        True, _
''        1 _
''    )
'    resultados.Add r
'End Sub
'
''===============================================================
'' REGLA 2: Módulo vacío (inteligente)
''===============================================================
'Private Sub Regla_ModuloVacio(m As clsModulo, resultados As Collection)
'    Dim r As clsResultadoAnalisis
'
'    ' Considerar atributos como contenido
'    If m.NumLineasCodigo = 0 And m.NumLineasComentario = 0 And m.NumLineasAtributo = 0 Then
'        Set r = CrearResultado( _
'            sevAviso, _
'            teModulo, _
'            m.nombre, _
'            "", _
'            tmUnknown, _
'            "El módulo está vacío.", _
'            "R002", _
'            "", _
'            False, _
'            0 _
'        )
'
''        Set r = CrearResultado( _
''            sevAviso, _
''            teModulo, _
''            m.nombre, _
''            "", _
''            "El módulo está vacío.", _
''            "R002", _
''            "", _
''            False, _
''            0 _
''        )
'        resultados.Add r
'    End If
'End Sub
'
'Private Sub Regla_ClaseVacia(c As clsClase, resultados As Collection)
'    Dim r As clsResultadoAnalisis
'
'    If c.NumLineasCodigo = 0 And c.NumLineasComentario = 0 And c.NumLineasAtributo = 0 Then
'        Set r = CrearResultado( _
'            sevAviso, _
'            teClase, _
'            c.nombre, _
'            "", _
'            tmUnknown, _
'            "La clase está vacía.", _
'            "R002", _
'            "", _
'            False, _
'            0 _
'        )
'
''        Set r = CrearResultado( _
''            sevAviso, _
''            teClase, _
''            c.nombre, _
''            "", _
''            "La clase está vacía.", _
''            "R002", _
''            "", _
''            False, _
''            0 _
''        )
'        resultados.Add r
'    End If
'End Sub
'
''===============================================================
'' REGLA 3: Miembros duplicados (distinción por tipo)
''===============================================================
'Private Sub Regla_MiembrosDuplicados(m As clsModulo, resultados As Collection, tipoContenedor As String)
'    Dim dict As Object
'    Dim mi As clsMiembro
'    Dim r As clsResultadoAnalisis
'    Dim clave As String
'
'    Set dict = CreateObject("Scripting.Dictionary")
'
'    For Each mi In m.Miembros
'        clave = LCase$(mi.nombre)
'
'        If dict.Exists(clave) Then
'
'            Dim descripcion As String
'            descripcion = "Miembro duplicado en "
'
'            Select Case tipoContenedor
'                Case "UserForm": descripcion = descripcion & "un UserForm."
'                Case "Formulario": descripcion = descripcion & "un formulario de Access."
'                Case "Informe": descripcion = descripcion & "un informe de Access."
'                Case Else: descripcion = descripcion & "el módulo."
'            End Select
'
'            Set r = CrearResultado( _
'                sevError, _
'                teMiembro, _
'                m.nombre, _
'                mi.nombre, _
'                mi.tipo, _
'                descripcion, _
'                "R003", _
'                "", _
'                False, _
'                mi.LineaInicio _
'            )
'
''            Set r = CrearResultado( _
''                sevError, _
''                teMiembro, _
''                m.nombre, _
''                mi.nombre, _
''                descripcion, _
''                "R003", _
''                "", _
''                False, _
''                mi.LineaInicio _
''            )
'            resultados.Add r
'
'        Else
'            dict.Add clave, True
'        End If
'    Next mi
'End Sub
'
'Private Sub Regla_MiembrosDuplicados_Clase(c As clsClase, resultados As Collection)
'    Dim dict As Object
'    Dim mi As clsMiembro
'    Dim r As clsResultadoAnalisis
'
'    Set dict = CreateObject("Scripting.Dictionary")
'
'    For Each mi In c.Miembros
'        If dict.Exists(LCase$(mi.nombre)) Then
'            Set r = CrearResultado( _
'                sevError, _
'                teMiembro, _
'                c.nombre, _
'                mi.nombre, _
'                mi.tipo, _
'                "Miembro duplicado en la clase.", _
'                "R003", _
'                "", _
'                False, _
'                mi.LineaInicio _
'            )
'
''            Set r = CrearResultado( _
''                sevError, _
''                teMiembro, _
''                c.nombre, _
''                mi.nombre, _
''                "Miembro duplicado en la clase.", _
''                "R003", _
''                "", _
''                False, _
''                mi.LineaInicio _
''            )
'            resultados.Add r
'        Else
'            dict.Add LCase$(mi.nombre), True
'        End If
'    Next mi
'End Sub
'
''===============================================================
'' Constructor de resultados
''===============================================================
'Private Function CrearResultado( _
'    sev As SeveridadInspector, _
'    tipo As TipoElementoInspector, _
'    nombreElemento As String, _
'    nombreMiembro As String, _
'    tipoMiembro As TipoMiembroInspector, _
'    descripcion As String, _
'    codigoRegla As String, _
'    codigoReparacion As String, _
'    esReparable As Boolean, _
'    linea As Long _
') As clsResultadoAnalisis
'
'    Dim res As New clsResultadoAnalisis
'
'    res.severidad = sev
'    res.TipoElemento = tipo
'    res.nombreElemento = nombreElemento
'    res.nombreMiembro = nombreMiembro
'    res.tipoMiembro = tipoMiembro
'    res.descripcion = descripcion
'    res.codigoRegla = codigoRegla
'    res.codigoReparacion = codigoReparacion
'    res.esReparable = esReparable
'    res.linea = linea
'
'    Set CrearResultado = res
'End Function
'
'
''Private Function CrearResultado( _
''    sev As SeveridadInspector, _
''    tipo As TipoElementoInspector, _
''    nombreElemento As String, _
''    nombreMiembro As String, _
''    descripcion As String, _
''    codigoRegla As String, _
''    codigoReparacion As String, _
''    esReparable As Boolean, _
''    linea As Long _
'') As clsResultadoAnalisis
''
''    Dim r As New clsResultadoAnalisis
''
''    r.severidad = sev
''    r.TipoElemento = tipo
''    r.nombreElemento = nombreElemento
''    r.nombreMiembro = nombreMiembro
''    r.descripcion = descripcion
''    r.codigoRegla = codigoRegla
''    r.codigoReparacion = codigoReparacion
''    r.esReparable = esReparable
''    r.linea = linea
''
''    Set CrearResultado = r
''End Function
'

'
