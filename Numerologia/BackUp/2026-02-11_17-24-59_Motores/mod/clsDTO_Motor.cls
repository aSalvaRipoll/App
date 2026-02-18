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

' Sílabas después de revisión manual
Public SilabasFinal As String

' Índices globales de sílabas tónicas (1-based)
Public SilabaTonica As String

' Índices globales de sílabas secundarias (1-based)
Public SilabaSecundaria As String

' Resultado fonético final (string lineal de IdFonema)
Public TextoFinal As String

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

    ' Sustituir unicode problemático
    s = Replace(s, "–", "-")
    s = Replace(s, "—", "-")
    s = Replace(s, "“", """")
    s = Replace(s, "”", """")

    s = Replace(s, " -", "-")
    s = Replace(s, "- ", "-")

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

' Sílabas después de revisión manual
SilabasFinal = vbNullString

' Índices globales de sílabas tónicas (1-based)
SilabaTonica = vbNullString

' Índices globales de sílabas secundarias (1-based)
SilabaSecundaria = vbNullString

' Resultado fonético final (string lineal de IdFonema)
TextoFinal = vbNullString

End Sub
