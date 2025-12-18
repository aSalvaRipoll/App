Attribute VB_Name = "modFactory"

'Option Compare Database
'Option Explicit
'
'' Crea una instancia de cualquier clase del proyecto por nombre
'Public Function CrearInstancia(nombreClase As String) As Object
'    On Error GoTo ErrHandler
'
'    ' Creamos un objeto vacío
'    Dim dummy As Object
'    Set dummy = New Collection
'
'    ' Llamamos al constructor de la clase usando CallByName
'    ' Esto funciona porque VBA permite instanciar clases internas
'    ' a través de un miembro ficticio del módulo estándar.
'    Set CrearInstancia = CallByName(dummy, nombreClase, VbGet)
'    Exit Function
'
'ErrHandler:
'    Set CrearInstancia = Nothing
'End Function

