VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDTO_Motor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ================================================================
' Nombre:    clsDTO_Motor
' Tipo:      Módulo de clase
' Propósito: Objeto contenedor de los datos durante el tratamiento
'           por el motor finético V 3.0
' Autor:     Alba Salvá
' Fecha:     11/02/2026
' ================================================================

Option Compare Database
Option Explicit

Public TextoOriginal As String
Public TextoCorregido As String
Public TextoNormalizado As String

' Sílabas generadas automáticamente (con marcadores)
Public SilabasAuto As String

' Sílabas después de marcar acentos
Public SilabasAcentuadas As String

' Sílabas después de revisión manual
Public SilabasFinal As String

' Índices globales de sílabas tónicas (1-based)
Public SilabasTonicas As String

' Índices globales de sílabas secundarias (1-based)
Public SilabasSecundarias As String

' Índices globales de ligaduras (1-based)
Public Ligaduras As String

' Resultado fonético final (string lineal de IdFonema)
Public IdsFonemas As String

' Cadena fonética final
Public FonemasFinal As String

'Variable para contener el tiempo de ejecución del proceso.
Public Tiempo As String

' ----------------------------------------------------------------
' Procedimiento: NormalizaEntrada
' Propósito:     Normalizar el texto de entrada para su uso posterior
'               en las diversas operaciones del motor fonético V 3.0
' Tipo proc.:    Sub
' Acceso proc.:  Public
' Autor:         Alba Salvá
' Fecha:         11/02/2026
' ----------------------------------------------------------------
' ================================================================
'   NORMALIZACIÓN
' ================================================================
Public Sub NormalizaEntrada()

    Dim s As String

    s = Me.TextoOriginal

    ' Eliminar caracteres no fonéticos
    s = Replace(s, vbTab, " ")
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")

    ' Normalizar apóstrofos
    s = Replace(s, "’", "'")
    s = Replace(s, "´", "'")
    s = Replace(s, "`", "'")
    
    ' Normalizar guiones tipográficos
    s = Replace(s, "–", "-")
    s = Replace(s, "—", "-")
    s = Replace(s, " –", "-")
    s = Replace(s, " —", "-")
    s = Replace(s, "– ", "-")
    s = Replace(s, "— ", "-")
   
'    s = Replace(s, " -", "-")
'    s = Replace(s, "- ", "-")
    s = Replace(s, " -", "-")
    s = Replace(s, "- ", "-")
   
'    s = Replace(s, "–", "-")
'    s = Replace(s, "—", "-")
    
    ' Sustituir comillas dobles
    s = Replace(s, "“", """")
    s = Replace(s, "”", """")

'    s = Replace(s, " -", "-")
'    s = Replace(s, "- ", "-")
    
'    ' Eliminar guión
'    s = Replace(s, "-", " ")

    ' LimpiarPuntuacion
    s = Replace(s, ",", " ")
    s = Replace(s, ".", " ")
    s = Replace(s, ";", " ")
    s = Replace(s, ":", " ")
    s = Replace(s, "¿", " ")
    s = Replace(s, "?", " ")
    s = Replace(s, "¡", " ")
    s = Replace(s, "!", " ")

    ' Sustituimos guines por ligaduras manuales
    s = Replace(s, "-", "_")
    
    ' Limpiar ligaduras múltiples
    Do While InStr(s, "__") > 0
        s = Replace(s, "__", "_")
    Loop

    ' Limpiar espacios múltiples
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop

    s = Trim$(s)

    ' Guardar texto limpio
    Me.TextoCorregido = s
    
    ' Guardar texto normalizado
    Me.TextoNormalizado = LCase$(s)

End Sub
' ============================================================
'   NORMALIZACIÓN ORTOGRÁFICA IB
' ============================================================
Function Normalizar_IB(t As String) As String
    Dim s As String

    s = t

    

    ' 2. No tocar vocales balears (á, à, è, é, í, ï, ò, ó, ú, ü, ç)

 
    ' 4. Espacios múltiples
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop

    s = Trim$(s)

    Normalizar_IB = s
End Function
' ----------------------------------------------------------------
' Procedimiento: Class_Initialize
' Propósito:     Inicializar las cadenas de texto contenedoras
' Tipo proc.:    Constructor (Initialize)
' Acceso proc.:  Private
' Autor:        Alba Salvá
' Fecha:        11/02/2026
' ----------------------------------------------------------------
Private Sub Class_Initialize()

' Textos base
TextoOriginal = vbNullString
TextoCorregido = vbNullString
TextoNormalizado = vbNullString

' Sílabas generadas automáticamente (con marcadores)
SilabasAuto = vbNullString

' Sílabas después de marcar acentos
SilabasAcentuadas = vbNullString

' Sílabas después de revisión manual
SilabasFinal = vbNullString

' Índices globales de sílabas tónicas (1-based)
SilabasTonicas = vbNullString

' Índices globales de sílabas secundarias (1-based)
SilabasSecundarias = vbNullString

' Índices globales de ligaduras (1-based)
Ligaduras = vbNullString

' Resultado fonético final (string lineal de IdFonema)
IdsFonemas = vbNullString
FonemasFinal = vbNullString

'Tiempo empleado
Tiempo = vbNullString

End Sub

Private Sub Class_Terminate()

' Textos base
TextoOriginal = vbNullString
TextoCorregido = vbNullString
TextoNormalizado = vbNullString

' Sílabas generadas automáticamente (con marcadores)
SilabasAuto = vbNullString

' Sílabas después de marcar acentos
SilabasAcentuadas = vbNullString

' Sílabas después de revisión manual
SilabasFinal = vbNullString

' Índices globales de sílabas tónicas (1-based)
SilabasTonicas = vbNullString

' Índices globales de sílabas secundarias (1-based)
SilabasSecundarias = vbNullString

' Índices globales de ligaduras (1-based)
Ligaduras = vbNullString

' Resultado fonético final (string lineal de IdFonema)
IdsFonemas = vbNullString
FonemasFinal = vbNullString

'Tiempo empleado
Tiempo = vbNullString

End Sub


'Private Function LimpiarPuntuacion(ByVal s As String) As String
'    s = Replace(s, ",", " ")
'    s = Replace(s, ".", " ")
'    s = Replace(s, ";", " ")
'    s = Replace(s, ":", " ")
'    s = Replace(s, "¿", " ")
'    s = Replace(s, "?", " ")
'    s = Replace(s, "¡", " ")
'    s = Replace(s, "!", " ")
'    LimpiarPuntuacion = s
'End Function
