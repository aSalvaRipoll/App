VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTransito_0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

'Clase clsTransito
' DTO para un tránsito anual: contiene valores físicos, mentales,
' emocionales y espirituales, junto con las letras que los generan.

Public IDPersona As Long
Public IDResultado As Long
Public idTransito As Long
Public IDFonetica As Long

Public orden As Byte
Public anio As Integer
Public Edad As Byte

Public Fisico As String
Public Mental As String
Public Emocional As String
Public Espiritual As String

Public LetraFisico As String
Public LetraMental As String
Public LetraEmocional As String
Public LetraEspiritual As String

Public Esencia As Integer
Public AnioPersonal As Integer

Private Sub Class_Initialize()

    orden = 0
    anio = 0
    Edad = 0

    Fisico = ""
    Mental = ""
    Emocional = ""
    Espiritual = ""

    LetraFisico = ""
    LetraMental = ""
    LetraEmocional = ""
    LetraEspiritual = ""

    Esencia = 0
    AnioPersonal = 0

End Sub

