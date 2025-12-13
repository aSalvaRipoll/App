VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsModulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'=====================================================
' Clase: clsModulo
' Representa un módulo VBA (estándar, objeto o UserForm)
'=====================================================

Option Compare Database
Option Explicit

Public nombre As String
Public Tipo As String              ' StdModule, Document, UserForm
Public Lineas() As String
Public Miembros As Collection

Public NumLineas As Long
Public NumLineasCodigo As Long
Public NumLineasComentario As Long
Public NumLineasVacias As Long

Public NumMiembros As Long
Public NumProcedimientos As Long
Public NumFunciones As Long
Public NumPropiedades As Long
Public NumEventos As Long

Private Sub Class_Initialize()
    Set Miembros = New Collection
End Sub

