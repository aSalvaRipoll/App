VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMotorFonetico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'' clsMotorFonetico.cls
'Option Compare Database
'Option Explicit
'
'Public TextoOriginal As String
'Public TextoNormalizado As String
'
'' Sílabas en texto, antes y después de la revisión
'Public SilabasAuto() As String
'Public SilabasFinal() As String
'
'' Índices de sílabas tónicas (1-based: 1 = primera sílaba)
'Public SilabaTonica() As Byte
'
'' Resultado fonético
'Public Fonemas As Collection      ' colección por sílabas o lineal, según decidas
'Public TextoFinal As String       ' cadena descodificada final

Option Compare Database
Option Explicit

Public TextoOriginal As String
Public TextoNormalizado As String

' Sílabas en texto, antes y después de la revisión
'Public SilabasAuto() As String
Public SilabasAuto As String
'Public SilabasFinal() As String
Public SilabasFinal As String

' Índices de sílabas tónicas (1-based)
'Public SilabaTonica() As Byte
Public SilabaTonica As String

' Resultado fonético final (string lineal de IdFonema)
Public TextoFinal As String




