VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDTOcalculos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

Public IDPersona As Long
Public IDFonetica As Long   ' ? ESTE ES EL QUE FALTABA

Public Nombre As String
Public Ape1 As String
Public Ape2 As String

Public FechaNacimiento As Date

Public SistemaFonetico As ModoFonetico
Public SistemaCalculo As ModoCalculo
Public NumCiclos As Byte
Public SistemaCiclos As ModoCiclos
Public SistemaTarot As ModoTarot

Public Sub CargarDatos(Persona As clsPersona, _
                       Fonetica As clsFonetica, _
                       SistemaCalculo As ModoCalculo, _
                       NumCiclos As Byte, _
                       SistemaCiclos As ModoCiclos, _
                       SistemaTarot As ModoTarot)

    ' --- Datos identificativos ---
    IDPersona = Persona.ID_Persona
    IDFonetica = Fonetica.IDFonetica
    
    ' --- Datos convertidos ---
    Me.Nombre = Fonetica.FonNombre
    Me.Ape1 = Fonetica.FonApe1
    Me.Ape2 = Fonetica.FonApe2

    ' --- Datos base ---
    Me.FechaNacimiento = Persona.FechaNacimiento
    Me.SistemaFonetico = Fonetica.ModoFonetico
    
    ' --- Sistemas ---
    Me.SistemaCalculo = SistemaCalculo
    Me.NumCiclos = NumCiclos
    Me.SistemaCiclos = SistemaCiclos
    Me.SistemaTarot = SistemaTarot

End Sub


