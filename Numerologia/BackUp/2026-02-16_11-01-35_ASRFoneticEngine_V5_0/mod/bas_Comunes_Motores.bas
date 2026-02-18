Attribute VB_Name = "bas_Comunes_Motores"

Option Compare Database
Option Explicit


Public Type ConfigFon
    ModoSibilantes As Byte
    ModoLateral As Byte
    ModoH As Byte
    ModoX As Byte
    ModoLigadura As Byte
    'DigrafosActivos As Collection
End Type

Public CFG As ConfigFon
Public ObjDTO As clsDTO_Motor

Public strDebug As String

Public modoDebug As Boolean
Public DebugMotor As Boolean

Public Function CreaSQL(Idioma As String) As String

    CreaSQL = _
        "SELECT ID, [" & Idioma & "]" & vbCrLf & _
        "FROM tbmDigrafosIdioma" & vbCrLf & _
        "WHERE Nz([" & Idioma & "],'') <> '';"

End Function

