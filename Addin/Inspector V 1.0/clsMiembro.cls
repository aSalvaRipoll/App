VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMiembro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'=====================================================
' Clase: clsMiembro
' Representa un miembro de un módulo o clase
'=====================================================

Option Compare Database
Option Explicit

Public nombre As String
Public Tipo As String          ' Sub, Function, Property Get/Let/Set
Public Ambito As String        ' Public, Private, Friend
Public lineaInicio As Long
Public LineaFin As Long
Public NumLineas As Long

Public NumLineasCodigo As Long
Public NumLineasComentario As Long
Public NumLineasVacias As Long

Public EsEvento As Boolean
Public EsPropiedad As Boolean
Public EsFuncion As Boolean
Public EsProcedimiento As Boolean

