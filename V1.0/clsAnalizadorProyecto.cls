VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAnalizadorProyecto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

Private pResultados As Collection

Public Property Get Resultados() As Collection
    Set Resultados = pResultados
End Property

Private Sub Class_Initialize()
    Set pResultados = New Collection
End Sub

Private Sub AgregarResultado( _
        ByVal Tipo As TipoElementoInspector, _
        ByVal nombre As String, _
        ByVal sev As SeveridadInspector, _
        ByVal descripcion As String, _
        Optional ByVal esReparable As Boolean = False, _
        Optional ByVal codigoReparacion As String = "")
    
    Dim r As clsResultadoAnalisis
    Set r = New clsResultadoAnalisis
    
    r.TipoElemento = Tipo
    r.NombreElemento = nombre
    r.Severidad = sev
    r.descripcion = descripcion
    r.esReparable = esReparable
    r.codigoReparacion = codigoReparacion
    
    pResultados.Add r
End Sub

Public Sub AnalizarProyectoActual()
    Dim vbProj As Object
    Set vbProj = Application.VBE.ActiveVBProject
    
    AnalizarClases vbProj
    AnalizarModulos vbProj
    AnalizarReferencias vbProj
End Sub

' --- Clases ---
Private Sub AnalizarClases(vbProj As Object)
    Dim comp As Object
    Dim nombre As String
    
    For Each comp In vbProj.VBComponents
        If comp.Type = vbext_ct_ClassModule Then
            nombre = comp.Name
            AnalizarClaseIndividual comp, nombre
        End If
    Next comp
End Sub

Private Sub AnalizarClaseIndividual(comp As Object, nombre As String)
    ' Nombre
    If Len(nombre) = 0 Then
        AgregarResultado tiClase, "(sin nombre)", sevError, _
            "Clase sin VB_Name.", False
    End If
    
    ' Option Explicit
    If Not TieneOptionExplicit(comp) Then
        AgregarResultado tiClase, nombre, sevAviso, _
            "Clase sin Option Explicit.", True, "ADD_OPTION_EXPLICIT"
    End If
    
    ' Instanciación
    If Not ClaseEsInstanciable(nombre) Then
        AgregarResultado tiClase, nombre, sevAviso, _
            "No se ha podido instanciar (puede ser normal si la clase no tiene constructor accesible).", _
            False
    Else
        AgregarResultado tiClase, nombre, sevInfo, "Instanciación OK.", False
    End If
End Sub

' --- Módulos estándar ---
Private Sub AnalizarModulos(vbProj As Object)
    Dim comp As Object
    Dim nombre As String
    
    For Each comp In vbProj.VBComponents
        If comp.Type = vbext_ct_StdModule Then
            nombre = comp.Name
            AnalizarModuloIndividual comp, nombre
        End If
    Next comp
End Sub

Private Sub AnalizarModuloIndividual(comp As Object, nombre As String)
    If Not TieneOptionExplicit(comp) Then
        AgregarResultado tiModulo, nombre, sevAviso, _
            "Módulo sin Option Explicit.", True, "ADD_OPTION_EXPLICIT"
    End If
End Sub

' --- Referencias ---
Private Sub AnalizarReferencias(vbProj As Object)
    On Error Resume Next
    Dim ref As Reference
    
    For Each ref In vbProj.References
        If ref.IsBroken Then
            AgregarResultado tiReferencia, ref.Name, sevError, _
                "Referencia rota.", True, "FIX_REFERENCE:" & ref.Name
        End If
    Next ref
End Sub

' --- Utilidades ---
'Private Function TieneOptionExplicit(comp As Object) As Boolean
'    Dim i As Long
'    Dim linea As String
'
'    With comp.CodeModule
'        For i = 1 To .CountOfLines
'            linea = Trim(.Lines(i, 1))
'            If UCase$(linea) = "OPTION EXPLICIT" Then
'                TieneOptionExplicit = True
'                Exit Function
'            End If
'            If Len(linea) > 0 And Left$(linea, 1) <> "'" Then Exit For
'        Next i
'    End With
'End Function

Private Function TieneOptionExplicit(comp As Object) As Boolean
    Dim i As Long
    Dim linea As String
    
    With comp.CodeModule
        For i = 1 To .CountOfLines
            linea = Trim(.Lines(i, 1))
            
            ' Saltar líneas vacías
            If Len(linea) = 0 Then GoTo Siguiente
            
            ' Saltar comentarios
            If Left$(linea, 1) = "'" Then GoTo Siguiente
            
            ' Saltar atributos VBA
            If Left$(UCase$(linea), 9) = "ATTRIBUTE" Then GoTo Siguiente
            
            ' Si encontramos Option Explicit ? OK
            If UCase$(linea) = "OPTION EXPLICIT" Then
                TieneOptionExplicit = True
                Exit Function
            End If
            
            ' Si llegamos aquí, es la primera línea real y no es Option Explicit
            TieneOptionExplicit = False
            Exit Function
            
Siguiente:
        Next i
    End With
End Function


Private Function ClaseEsInstanciable(nombreClase As String) As Boolean
    Dim f As New clsFactory
    Dim obj As Object
    
    Set obj = f.CrearInstancia(nombreClase)
    ClaseEsInstanciable = Not (obj Is Nothing)
End Function



'Option Compare Database
'Option Explicit
'
'Private pResultados As Collection
'
'Public Property Get Resultados() As Collection
'    Set Resultados = pResultados
'End Property
'
'Private Sub Class_Initialize()
'    Set pResultados = New Collection
'End Sub
'
''Private Sub AgregarResultado( _
''        ByVal tipo As TipoElementoInspector, _
''        ByVal nombre As String, _
''        ByVal sev As SeveridadInspector, _
''        ByVal descripcion As String)
''
''    Dim r As clsResultadoAnalisis
''    Set r = New clsResultadoAnalisis
''
''    r.TipoElemento = tipo
''    r.NombreElemento = nombre
''    r.Severidad = sev
''    r.descripcion = descripcion
''
''    pResultados.Add r
''End Sub
'
'Private Sub AgregarResultado( _
'        ByVal tipo As TipoElementoInspector, _
'        ByVal nombre As String, _
'        ByVal sev As SeveridadInspector, _
'        ByVal Descripcion As String, _
'        Optional ByVal EsReparable As Boolean = False, _
'        Optional ByVal CodigoReparacion As String = "")
'
'    Dim r As clsResultadoAnalisis
'    Set r = New clsResultadoAnalisis
'
'    r.TipoElemento = tipo
'    r.NombreElemento = nombre
'    r.Severidad = sev
'    r.Descripcion = Descripcion
'    r.EsReparable = EsReparable
'    r.CodigoReparacion = CodigoReparacion
'
'    pResultados.Add r
'End Sub
'
'' ============================================================
'' MÉTODO PRINCIPAL DE ANÁLISIS
'' ============================================================
'Public Sub AnalizarProyectoActual()
'    Dim vbProj As Object
'    Set vbProj = Application.VBE.ActiveVBProject
'
'    AnalizarClases vbProj
'    AnalizarModulos vbProj
'    AnalizarReferencias vbProj
'End Sub
'
'' ============================================================
'' ANÁLISIS DE CLASES
'' ============================================================
'Private Sub AnalizarClases(vbProj As Object)
'    Dim comp As Object
'    Dim nombre As String
'
'    For Each comp In vbProj.VBComponents
'        If comp.Type = vbext_ct_ClassModule Then
'            nombre = comp.Name
'            AnalizarClaseIndividual comp, nombre
'        End If
'    Next comp
'End Sub
'
'Private Sub AnalizarClaseIndividual(comp As Object, nombre As String)
'
'    ' 1. Nombre válido
'    If Len(nombre) = 0 Then
'        AgregarResultado tiClase, "(sin nombre)", sevError, "Clase sin VB_Name."
'    End If
'
'    ' 2. Option Explicit
'    If Not TieneOptionExplicit(comp) Then
'        AgregarResultado tiClase, nombre, sevAviso, "Clase sin Option Explicit."
'    End If
'
'    ' 3. Instanciación básica (como hicimos antes)
'    If Not ClaseEsInstanciable(nombre) Then
'        AgregarResultado tiClase, nombre, sevError, "No se puede instanciar con New (posible clase corrupta o no probada)."
'    Else
'        AgregarResultado tiClase, nombre, sevInfo, "Instanciación OK."
'    End If
'
'End Sub
'
'' ============================================================
'' ANÁLISIS DE MÓDULOS ESTÁNDAR
'' ============================================================
'Private Sub AnalizarModulos(vbProj As Object)
'    Dim comp As Object
'    Dim nombre As String
'
'    For Each comp In vbProj.VBComponents
'        If comp.Type = vbext_ct_StdModule Then
'            nombre = comp.Name
'            AnalizarModuloIndividual comp, nombre
'        End If
'    Next comp
'End Sub
'
'Private Sub AnalizarModuloIndividual(comp As Object, nombre As String)
'
'    ' 1. Option Explicit
'    If Not TieneOptionExplicit(comp) Then
'        AgregarResultado tiModulo, nombre, sevAviso, "Módulo sin Option Explicit."
'    End If
'
'    ' Más reglas aquí si quieres…
'
'End Sub
'
'' ============================================================
'' ANÁLISIS DE REFERENCIAS
'' ============================================================
'Private Sub AnalizarReferencias(vbProj As Object)
'    On Error Resume Next
'    Dim ref As Object
'
'    For Each ref In vbProj.References
'        If ref.IsBroken Then
'            AgregarResultado tiReferencia, ref.Name, sevError, "Referencia rota."
'        End If
'    Next ref
'End Sub
'
'' ============================================================
'' UTILIDADES
'' ============================================================
'Private Function TieneOptionExplicit(comp As Object) As Boolean
'    Dim i As Long
'    Dim linea As String
'
'    With comp.CodeModule
'        For i = 1 To .CountOfLines
'            linea = Trim(.Lines(i, 1))
'            If UCase$(linea) = "OPTION EXPLICIT" Then
'                TieneOptionExplicit = True
'                Exit Function
'            End If
'            ' Si llegamos a la primera línea de código no-comentario, paramos
'            If Len(linea) > 0 And Left$(linea, 1) <> "'" Then Exit For
'        Next i
'    End With
'End Function
'
''Private Function ClaseEsInstanciable(nombreClase As String) As Boolean
''    On Error GoTo ErrHandler
''
''    ' Aquí usamos el mismo enfoque manual que antes:
''    ' Tú puedes ir ampliando este Select Case según tus proyectos.
''    Dim obj As Object
''
''    Select Case nombreClase
''        Case "clsCalculoAlma": Set obj = New clsCalculoAlma
''        Case "clsCalculoBase": Set obj = New clsCalculoBase
''        Case "clsCalculoDestino": Set obj = New clsCalculoDestino
''        Case "clsCalculoCaminoVida": Set obj = New clsCalculoCaminoVida
''        Case "clsCalculoPersonalidad": Set obj = New clsCalculoPersonalidad
''        Case "clsCalculoMadurez": Set obj = New clsCalculoMadurez
''        Case "clsCalculoSinastria": Set obj = New clsCalculoSinastria
''        Case "clsCalculadorPeriodoActual": Set obj = New clsCalculadorPeriodoActual
''        Case "clsGestorInterpretaciones": Set obj = New clsGestorInterpretaciones
''        Case "clsConversorNumerologico": Set obj = New clsConversorNumerologico
''
''        ' En otros proyectos, podrás añadir otras clases aquí.
''
''        Case Else
''            ' No sabemos cómo probar esta clase en este proyecto
''            ClaseEsInstanciable = False
''            Exit Function
''    End Select
''
''    ClaseEsInstanciable = True
''    Exit Function
''
''ErrHandler:
''    ClaseEsInstanciable = False
''End Function
'
'Private Function ClaseEsInstanciable(nombreClase As String) As Boolean
'    Dim obj As Object
'    On Error GoTo ErrHandler
'
'    Set obj = CrearInstancia(nombreClase)
'    ClaseEsInstanciable = Not obj Is Nothing
'    Exit Function
'
'ErrHandler:
'    ClaseEsInstanciable = False
'End Function
'
