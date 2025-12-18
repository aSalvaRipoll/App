VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsResultados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

'===============================================================
' Clase: clsResultados
' Contenedor de resultados del análisis del Inspector
'---------------------------------------------------------------
' Responsabilidades:
'   - Almacenar todos los resultados (clsResultadoAnalisis)
'   - Exponerlos de forma controlada (Todos)
'   - Permitir obtener un resultado por clave
'   - Navegar a un resultado en el editor
'   - Ordenar resultados por cualquier campo visible
'===============================================================

Private mResultados As Collection

'---------------------------------------------------------------
' Inicialización
'---------------------------------------------------------------
Private Sub Class_Initialize()
    Set mResultados = New Collection
End Sub

'---------------------------------------------------------------
' Agregar un resultado
'---------------------------------------------------------------
Public Sub AgregarResultado(res As clsResultadoAnalisis)
    On Error Resume Next
    mResultados.Add res, res.clave
End Sub

'---------------------------------------------------------------
' Obtener un resultado por clave
'---------------------------------------------------------------
Public Function ObtenerResultado(clave As String) As clsResultadoAnalisis
    On Error Resume Next
    Set ObtenerResultado = mResultados(clave)
End Function

'---------------------------------------------------------------
' Enumerar todos los resultados
'---------------------------------------------------------------
Public Function Todos() As Collection
    Set Todos = mResultados
End Function

'---------------------------------------------------------------
' Navegar a un resultado concreto
'---------------------------------------------------------------
Public Sub NavegarAResultado(clave As String)

    Dim res As clsResultadoAnalisis
    Set res = Me.ObtenerResultado(clave)

    If res Is Nothing Then
        Inspector_Log "NavegarAResultado: resultado no encontrado (" & clave & ")"
        Exit Sub
    End If

    ' Llamada al módulo de navegación
    Call NavegarAModulo(res.nombreElemento, res.nombreMiembro, res.linea)

End Sub

'===============================================================
' ORDENACIÓN DE RESULTADOS
'===============================================================

'---------------------------------------------------------------
' Ordenar por un campo (Severidad, Elemento, Miembro, Línea, Clave)
'---------------------------------------------------------------
Public Sub OrdenarPor(campo As String, Optional asc As Boolean = True)
    Dim i As Long, j As Long
    Dim tmp As clsResultadoAnalisis

    For i = 1 To mResultados.Count - 1
        For j = i + 1 To mResultados.Count

            If Comparar(mResultados(i), mResultados(j), campo, asc) Then
                Set tmp = mResultados(i)
                Set mResultados(i) = mResultados(j)
                Set mResultados(j) = tmp
            End If

        Next j
    Next i
End Sub

'---------------------------------------------------------------
' Comparación genérica entre dos resultados
'---------------------------------------------------------------
Private Function Comparar(a As clsResultadoAnalisis, _
                          b As clsResultadoAnalisis, _
                          campo As String, _
                          asc As Boolean) As Boolean

    Dim va, vb

    Select Case campo
        Case "Severidad":      va = a.severidad:      vb = b.severidad
        Case "Elemento":       va = a.nombreElemento: vb = b.nombreElemento
        Case "Miembro":        va = a.nombreMiembro:  vb = b.nombreMiembro
        Case "Linea":          va = a.linea:          vb = b.linea
        Case Else:             va = a.clave:          vb = b.clave
    End Select

    If asc Then
        Comparar = (va > vb)
    Else
        Comparar = (va < vb)
    End If

End Function

'Option Compare Database
'Option Explicit
'
''===============================================================
'' Clase: clsResultados
'' Contiene todos los resultados del análisis del Inspector
''===============================================================
'
'Private mResultados As Collection
'
''---------------------------------------------------------------
'' Inicialización
''---------------------------------------------------------------
'Private Sub Class_Initialize()
'    Set mResultados = New Collection
'End Sub
'
''---------------------------------------------------------------
'' Agregar un resultado
''---------------------------------------------------------------
'Public Sub AgregarResultado(res As clsResultadoAnalisis)
'    On Error Resume Next
'    mResultados.Add res, res.clave
'End Sub
'
''---------------------------------------------------------------
'' Obtener un resultado por clave
''---------------------------------------------------------------
'Public Function ObtenerResultado(clave As String) As clsResultadoAnalisis
'    On Error Resume Next
'    Set ObtenerResultado = mResultados(clave)
'End Function
'
''---------------------------------------------------------------
'' Navegar a un resultado concreto
''---------------------------------------------------------------
'Public Sub NavegarAResultado(clave As String)
'
'    Dim res As clsResultadoAnalisis
'    Set res = Me.ObtenerResultado(clave)
'
'    If res Is Nothing Then
'        Inspector_Log "NavegarAResultado: resultado no encontrado (" & clave & ")"
'        Exit Sub
'    End If
'
'    ' Llamada al módulo de navegación
'    Call NavegarAModulo(res.nombreElemento, res.nombreMiembro, res.linea)
'
'End Sub
'
''---------------------------------------------------------------
'' Enumerar resultados (si lo necesitas en el futuro)
''---------------------------------------------------------------
'Public Function Todos() As Collection
'    Set Todos = mResultados
'End Function
'
'
''===============================================================
'' ORDENACIÓN DE RESULTADOS
''===============================================================
'Public Sub OrdenarPor(campo As String, Optional asc As Boolean = True)
'    Dim i As Long, j As Long
'    Dim tmp As clsResultadoAnalisis
'
'    For i = 1 To mResultados.Count - 1
'        For j = i + 1 To mResultados.Count
'
'            If Comparar(mResultados(i), mResultados(j), campo, asc) Then
'                Set tmp = mResultados(i)
'                Set mResultados(i) = mResultados(j)
'                Set mResultados(j) = tmp
'            End If
'
'        Next j
'    Next i
'End Sub
'
'Private Function Comparar(a As clsResultadoAnalisis, b As clsResultadoAnalisis, campo As String, asc As Boolean) As Boolean
'    Dim va, vb
'
'    Select Case campo
'        Case "Severidad":      va = a.severidad:      vb = b.severidad
'        Case "Elemento":       va = a.nombreElemento: vb = b.nombreElemento
'        Case "Miembro":        va = a.nombreMiembro:  vb = b.nombreMiembro
'        Case "Linea":          va = a.linea:          vb = b.linea
'        Case Else:             va = a.clave:          vb = b.clave
'    End Select
'
'    If asc Then
'        Comparar = (va > vb)
'    Else
'        Comparar = (va < vb)
'    End If
'End Function
'
'
'
'
'''---------------------------------------------------------------
''' Enumerar resultados (si lo necesitas en el futuro)
'''---------------------------------------------------------------
''Public Function Todos() As Collection
''    Set Todos = mResultados
''End Function
''
'''===============================================================
''' Ordenación de resultados
'''===============================================================
''Public Sub OrdenarPor(campo As String, Optional asc As Boolean = True)
''    Dim i As Long, j As Long
''    Dim tmp As clsResultadoAnalisis
''
''    For i = 1 To mResultados.Count - 1
''        For j = i + 1 To mResultados.Count
''
''            If Comparar(mResultados(i), mResultados(j), campo, asc) Then
''                Set tmp = mResultados(i)
''                Set mResultados(i) = mResultados(j)
''                Set mResultados(j) = tmp
''            End If
''
''        Next j
''    Next i
''End Sub
''
''Private Function Comparar(a As clsResultadoAnalisis, b As clsResultadoAnalisis, campo As String, asc As Boolean) As Boolean
''    Dim va, vb
''
''    Select Case campo
''        Case "Severidad":      va = a.severidad:      vb = b.severidad
''        Case "Elemento":       va = a.nombreElemento: vb = b.nombreElemento
''        Case "Miembro":        va = a.nombreMiembro:  vb = b.nombreMiembro
''        Case "Linea":          va = a.linea:          vb = b.linea
''        Case Else:             va = a.clave:          vb = b.clave
''    End Select
''
''    If asc Then
''        Comparar = (va > vb)
''    Else
''        Comparar = (va < vb)
''    End If
''End Function
'
