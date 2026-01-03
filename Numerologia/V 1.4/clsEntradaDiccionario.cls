VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEntradaDiccionario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

Public texto As String
Public fonema As String
Public idioma As String
Public tipo As String   ' "Nombre" o "Apellido"
Public Fuente As String ' EU, CA, GA, ES

