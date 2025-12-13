VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

Public Function CrearInstancia(nombreClase As String) As Object
    On Error GoTo ErrHandler
    
    ' Truco: CallByName resuelve el nombre de la clase en el proyecto activo.
    Dim dummy As Object
    Set dummy = New Collection
    
    Set CrearInstancia = CallByName(dummy, nombreClase, VbGet)
    Exit Function

ErrHandler:
    Set CrearInstancia = Nothing
End Function

