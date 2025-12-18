VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSimbolo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

'=====================================================
' Clase: clsSimbolo
' Representa un símbolo declarado en el proyecto VBA
'=====================================================

Public nombre As String          ' Nombre del símbolo
Public TipoTexto As String       ' Tipo declarado (Long, String, TipoCliente...)
Public categoria As String       ' Variable, Constante, Enum, UDT, Propiedad...
Public Ambito As String          ' Local, ModuloPrivate, ModuloPublic, Global
Public modulo As String          ' Módulo donde se declara
Public miembro As String         ' Miembro donde se declara (si es local)
Public LineaDeclaracion As Long  ' Línea de declaración
Public Usado As Boolean          ' ¿Se ha encontrado referencia?

' Información contextual opcional
Public EsAPI As Boolean          ' Para futuras detecciones de Declare
Public Comentario As String      ' Notas adicionales

'-----------------------------------------------------
' Inicialización rápida y segura
'-----------------------------------------------------
Public Sub Init(ByVal sNombre As String, _
                ByVal sCategoria As String, _
                ByVal sAmbito As String, _
                ByVal sModulo As String, _
                ByVal sMiembro As String, _
                ByVal lLinea As Long)

    ' Validación mínima
    If Len(sNombre) = 0 Then Err.Raise 5, , "clsSimbolo.Init: nombre vacío."
    If Len(sModulo) = 0 Then Err.Raise 5, , "clsSimbolo.Init: módulo vacío."

    ' Normalización
    nombre = Trim$(sNombre)
    categoria = Trim$(sCategoria)
    Ambito = Trim$(sAmbito)
    modulo = Trim$(sModulo)
    miembro = Trim$(sMiembro)
    LineaDeclaracion = lLinea

    Usado = False
End Sub

'-----------------------------------------------------
' Clave única del símbolo (útil para catálogos)
'-----------------------------------------------------
Public Property Get clave() As String
    clave = LCase$(modulo & "." & nombre & "." & Ambito & "." & categoria)
End Property

'-----------------------------------------------------
' Marcar símbolo como usado
'-----------------------------------------------------
Public Sub MarcarUsado()
    Usado = True
End Sub

'-----------------------------------------------------
' Propiedades semánticas útiles
'-----------------------------------------------------
Public Property Get EsLocal() As Boolean
    EsLocal = (LCase$(Ambito) = "local")
End Property

Public Property Get EsPublico() As Boolean
    EsPublico = (LCase$(Ambito) = "modulopublic" Or LCase$(Ambito) = "global")
End Property

Public Property Get EsPrivado() As Boolean
    EsPrivado = (LCase$(Ambito) = "moduloprivate")
End Property

'-----------------------------------------------------
' Descripción corta (para depuración y resultados)
'-----------------------------------------------------
Public Property Get DescripcionCorta() As String
    Dim txt As String
    txt = nombre & " (" & categoria & ", " & Ambito & ") en " & modulo
    If Len(miembro) > 0 Then txt = txt & "." & miembro
    DescripcionCorta = txt
End Property


