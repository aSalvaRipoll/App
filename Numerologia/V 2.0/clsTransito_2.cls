VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTransito_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

' DTO para un tránsito anual completo:
' - valores físico/mental/emocional/espiritual
' - letras/fonemas que los generan
' - esencia y año personal
' - metadatos de cálculo
' - identificadores para persistencia

Public IDPersona    As Long
Public IDResultado  As Long
Public idTransito   As Long
Public IDFonetica   As Long

Public orden    As Byte
Public anio     As Integer
Public Edad     As Byte

' Valores reducidos
Public Fisico       As String
Public Mental       As String
Public Emocional    As String
Public Espiritual   As String

' Letras/fonemas que generan cada tránsito
Public LetraFisico      As String
Public LetraMental      As String
Public LetraEmocional   As String
Public LetraEspiritual  As String

' Valores superiores
Public Esencia      As Integer
Public AnioPersonal As Integer

' Metadatos del cálculo
Public SistemaFonetico  As ModoFonetico
Public SistemaCalculo   As ModoCalculo

' Trazabilidad opcional
Public ValorFisicoBruto     As Integer
Public ValorMentalBruto     As Integer
Public ValorEmocionalBruto  As Integer
Public ValorEspiritualBruto As Integer

Public CicloFisico      As Integer
Public CicloMental      As Integer
Public CicloEmocional   As Integer
Public CicloEspiritual  As Integer

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

    SistemaFonetico = mfFonetico
    SistemaCalculo = mcClasico

    ValorFisicoBruto = 0
    ValorMentalBruto = 0
    ValorEmocionalBruto = 0
    ValorEspiritualBruto = 0

    CicloFisico = 0
    CicloMental = 0
    CicloEmocional = 0
    CicloEspiritual = 0

End Sub

