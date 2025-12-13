VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCalculoDiaNacimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =============================================================================
' Clase: clsCalculoDiaNacimiento
' Descripción: Calcula el Número del Día de Nacimiento
' Hereda de: clsCalculoBase (mediante composición)
' 
' El Día de Nacimiento revela talentos naturales innatos, habilidades
' con las que se nace, y el "superpoder" natural de la persona.
'
' Método de cálculo:
' El número del día de nacimiento (1-31) es el resultado directo.
' Los números maestros (11, 22, 29) NO se reducen.
' Los números kármicos (13, 14, 16, 19) tienen interpretación especial.
' =============================================================================

Private m_Base As clsCalculoBase
Private m_FechaNacimiento As Date
Private m_EsKarmico As Boolean
Private m_NumeroKarmico As String

' =============================================================================
' EVENTOS
' =============================================================================

Private Sub Class_Initialize()
    Set m_Base = New clsCalculoBase
    m_FechaNacimiento = #1/1/1900#
    m_EsKarmico = False
    m_NumeroKarmico = ""
End Sub

Private Sub Class_Terminate()
    Set m_Base = Nothing
End Sub

' =============================================================================
' PROPIEDADES
' =============================================================================

Public Property Let FechaNacimiento(ByVal valor As Date)
    m_FechaNacimiento = valor
End Property

Public Property Get FechaNacimiento() As Date
    FechaNacimiento = m_FechaNacimiento
End Property

Public Property Get Resultado() As Integer
    Resultado = m_Base.Resultado
End Property

Public Property Get EsNumeroMaestro() As Boolean
    EsNumeroMaestro = m_Base.EsNumeroMaestro
End Property

Public Property Get EsNumeroKarmico() As Boolean
    EsNumeroKarmico = m_EsKarmico
End Property

Public Property Get NumeroKarmico() As String
    ' Retorna el número kármico completo (ej: "13/4", "14/5", etc.)
    NumeroKarmico = m_NumeroKarmico
End Property

Public Property Get DetalleCalculo() As String
    DetalleCalculo = m_Base.DetalleCalculo
End Property

Public Property Get UltimoError() As String
    UltimoError = m_Base.UltimoError
End Property

Public Property Get NombreCalculo() As String
    NombreCalculo = "Día de Nacimiento"
End Property

' =============================================================================
' MÉTODOS PÚBLICOS
' =============================================================================

Public Function Calcular() As Integer
    ' Calcula el Número del Día de Nacimiento
    
    On Error GoTo ErrorHandler
    
    Dim diaNacimiento As Integer
    Dim numeroReducido As Integer
    
    ' Validar fecha
    If m_FechaNacimiento = #1/1/1900# Then
        m_Base.m_UltimoError = "Debe establecer una fecha de nacimiento válida"
        Calcular = 0
        Exit Function
    End If
    
    ' Limpiar detalle anterior
    m_Base.LimpiarDetalle
    m_Base.AgregarDetalle "=========================================="
    m_Base.AgregarDetalle "CÁLCULO DEL DÍA DE NACIMIENTO"
    m_Base.AgregarDetalle "=========================================="
    m_Base.AgregarDetalle "Fecha de nacimiento: " & Format(m_FechaNacimiento, "dd/mm/yyyy")
    m_Base.AgregarDetalle ""
    
    ' Extraer día de nacimiento
    diaNacimiento = Day(m_FechaNacimiento)
    m_Base.AgregarDetalle "Día de nacimiento: " & diaNacimiento
    m_Base.AgregarDetalle ""
    
    ' El resultado directo es el día de nacimiento
    m_Base.m_Resultado = diaNacimiento
    
    ' Verificar si es número maestro (11, 22, 29)
    m_Base.m_EsNumeroMaestro = (diaNacimiento = 11 Or diaNacimiento = 22 Or diaNacimiento = 29)
    
    ' Verificar si es número kármico (13, 14, 16, 19)
    m_EsKarmico = (diaNacimiento = 13 Or diaNacimiento = 14 Or diaNacimiento = 16 Or diaNacimiento = 19)
    
    If m_Base.m_EsNumeroMaestro Then
        m_Base.AgregarDetalle "----------------------------------------"
        m_Base.AgregarDetalle "¡NÚMERO MAESTRO!"
        m_Base.AgregarDetalle "----------------------------------------"
        
        Select Case diaNacimiento
            Case 11
                m_Base.AgregarDetalle "Número 11: El Iluminador Maestro"
                m_Base.AgregarDetalle "Misión de inspiración espiritual y despertar consciencias"
            Case 22
                m_Base.AgregarDetalle "Número 22: El Arquitecto Maestro"
                m_Base.AgregarDetalle "Misión de construcción de legados globales"
            Case 29
                numeroReducido = 11  ' 2+9=11
                m_Base.AgregarDetalle "Número 29/11: El Líder Iluminado"
                m_Base.AgregarDetalle "Reduce a: 2+9=" & numeroReducido
                m_Base.AgregarDetalle "Iluminación a través del servicio compasivo"
        End Select
        
    ElseIf m_EsKarmico Then
        m_Base.AgregarDetalle "----------------------------------------"
        m_Base.AgregarDetalle "¡NÚMERO KÁRMICO!"
        m_Base.AgregarDetalle "----------------------------------------"
        
        Select Case diaNacimiento
            Case 13
                numeroReducido = 4  ' 1+3=4
                m_NumeroKarmico = "13/4"
                m_Base.AgregarDetalle "Número 13/4: El Constructor Kármico"
                m_Base.AgregarDetalle "Reduce a: 1+3=" & numeroReducido
                m_Base.AgregarDetalle "Karma relacionado con trabajo, responsabilidad y transformación"
            Case 14
                numeroReducido = 5  ' 1+4=5
                m_NumeroKarmico = "14/5"
                m_Base.AgregarDetalle "Número 14/5: El Libertador Kármico"
                m_Base.AgregarDetalle "Reduce a: 1+4=" & numeroReducido
                m_Base.AgregarDetalle "Karma relacionado con abuso de libertad y excesos"
            Case 16
                numeroReducido = 7  ' 1+6=7
                m_NumeroKarmico = "16/7"
                m_Base.AgregarDetalle "Número 16/7: El Renacido Kármico"
                m_Base.AgregarDetalle "Reduce a: 1+6=" & numeroReducido
                m_Base.AgregarDetalle "Karma relacionado con abuso de poder y orgullo espiritual"
            Case 19
                numeroReducido = 10  ' 1+9=10, luego 1+0=1
                m_NumeroKarmico = "19/1"
                m_Base.AgregarDetalle "Número 19/10/1: El Líder Kármico"
                m_Base.AgregarDetalle "Reduce a: 1+9=10, 1+0=" & (numeroReducido Mod 10)
                m_Base.AgregarDetalle "Karma relacionado con abuso de liderazgo y egoísmo"
        End Select
        
    Else
        ' Números regulares o compuestos
        If diaNacimiento > 9 Then
            numeroReducido = m_Base.ReducirADigito(diaNacimiento, False)
            m_Base.AgregarDetalle "Número compuesto que reduce a: " & numeroReducido
            m_Base.AgregarDetalle "Interpretación: " & diaNacimiento & " (vibración completa)"
        Else
            m_Base.AgregarDetalle "Número básico: " & diaNacimiento
        End If
    End If
    
    m_Base.AgregarDetalle ""
    m_Base.AgregarDetalle "=========================================="
    m_Base.AgregarDetalle "RESULTADO: " & m_Base.m_Resultado
    
    If m_Base.m_EsNumeroMaestro Then
        m_Base.AgregarDetalle "TIPO: Número Maestro"
    ElseIf m_EsKarmico Then
        m_Base.AgregarDetalle "TIPO: Número Kármico (" & m_NumeroKarmico & ")"
    End If
    
    m_Base.AgregarDetalle "=========================================="
    
    Calcular = m_Base.m_Resultado
    Exit Function
    
ErrorHandler:
    m_Base.m_UltimoError = "Error en cálculo Día de Nacimiento: " & Err.Description
    Calcular = 0
End Function

Public Function ObtenerInterpretacion() As String
    ' Obtiene la interpretación del día de nacimiento desde archivo Markdown
    ' Utiliza el gestor de interpretaciones para cargar el archivo correcto
    
    On Error GoTo ErrorHandler
    
    Dim gestor As Object  ' clsGestorInterpretaciones
    Dim archivoInterpretacion As String
    Dim interpretacion As String
    
    ' Validar que se haya calculado primero
    If m_Base.m_Resultado = 0 Then
        ObtenerInterpretacion = "Debe calcular el Día de Nacimiento primero"
        Exit Function
    End If
    
    ' Crear gestor de interpretaciones
    Set gestor = CreateObject("clsGestorInterpretaciones")
    
    ' Construir nombre de archivo según el día
    ' Formato: XX_DiaNacimiento.md donde XX es el día con ceros a la izquierda
    archivoInterpretacion = Format(m_Base.m_Resultado, "00") & "_DiaNacimiento.md"
    
    ' Cargar interpretación
    interpretacion = gestor.CargarInterpretacion("DiaNacimiento", archivoInterpretacion)
    
    ' Liberar objeto
    Set gestor = Nothing
    
    ObtenerInterpretacion = interpretacion
    Exit Function
    
ErrorHandler:
    m_Base.m_UltimoError = "Error al obtener interpretación: " & Err.Description
    ObtenerInterpretacion = "Error: " & Err.Description
    Set gestor = Nothing
End Function

Public Function ObtenerResumen() As String
    ' Retorna un resumen breve del día de nacimiento calculado
    
    Dim resumen As String
    
    If m_Base.m_Resultado = 0 Then
        ObtenerResumen = "No calculado"
        Exit Function
    End If
    
    resumen = "Día de Nacimiento: " & m_Base.m_Resultado
    
    If m_Base.m_EsNumeroMaestro Then
        Select Case m_Base.m_Resultado
            Case 11
                resumen = resumen & " - El Iluminador Maestro"
            Case 22
                resumen = resumen & " - El Arquitecto Maestro"
            Case 29
                resumen = resumen & " - El Líder Iluminado (29/11)"
        End Select
    ElseIf m_EsKarmico Then
        Select Case m_Base.m_Resultado
            Case 13
                resumen = resumen & " - El Constructor Kármico (13/4)"
            Case 14
                resumen = resumen & " - El Libertador Kármico (14/5)"
            Case 16
                resumen = resumen & " - El Renacido Kármico (16/7)"
            Case 19
                resumen = resumen & " - El Líder Kármico (19/1)"
        End Select
    Else
        ' Añadir descripción breve según el número
        resumen = resumen & " - " & ObtenerDescripcionBreve(m_Base.m_Resultado)
    End If
    
    ObtenerResumen = resumen
End Function

' =============================================================================
' MÉTODOS PRIVADOS
' =============================================================================

Private Function ObtenerDescripcionBreve(ByVal numero As Integer) As String
    ' Retorna una descripción breve del número del día de nacimiento
    
    Select Case numero
        Case 1: ObtenerDescripcionBreve = "El Líder Natural"
        Case 2: ObtenerDescripcionBreve = "El Diplomático Sensible"
        Case 3: ObtenerDescripcionBreve = "El Comunicador Creativo"
        Case 4: ObtenerDescripcionBreve = "El Constructor Sistemático"
        Case 5: ObtenerDescripcionBreve = "El Aventurero Versátil"
        Case 6: ObtenerDescripcionBreve = "El Cuidador Responsable"
        Case 7: ObtenerDescripcionBreve = "El Buscador Profundo"
        Case 8: ObtenerDescripcionBreve = "El Ejecutor Poderoso"
        Case 9: ObtenerDescripcionBreve = "El Humanitario Compasivo"
        Case 10: ObtenerDescripcionBreve = "El Líder Innovador (10/1)"
        Case 12: ObtenerDescripcionBreve = "El Comunicador Constructivo (12/3)"
        Case 15: ObtenerDescripcionBreve = "El Armonizador Versátil (15/6)"
        Case 17: ObtenerDescripcionBreve = "El Visionario Materializado (17/8)"
        Case 18: ObtenerDescripcionBreve = "El Humanitario Poderoso (18/9)"
        Case 20: ObtenerDescripcionBreve = "El Diplomático Amplificado (20/2)"
        Case 21: ObtenerDescripcionBreve = "El Creador Relacional (21/3)"
        Case 23: ObtenerDescripcionBreve = "El Comunicador Aventurero (23/5)"
        Case 24: ObtenerDescripcionBreve = "El Cuidador Práctico (24/6)"
        Case 25: ObtenerDescripcionBreve = "El Sanador Intuitivo (25/7)"
        Case 26: ObtenerDescripcionBreve = "El Constructor de Relaciones (26/8)"
        Case 27: ObtenerDescripcionBreve = "El Humanitario Empático (27/9)"
        Case 28: ObtenerDescripcionBreve = "El Líder Independiente (28/1)"
        Case 30: ObtenerDescripcionBreve = "El Creativo Amplificado (30/3)"
        Case 31: ObtenerDescripcionBreve = "El Constructor Creativo (31/4)"
        Case Else: ObtenerDescripcionBreve = "Día " & numero
    End Select
End Function
