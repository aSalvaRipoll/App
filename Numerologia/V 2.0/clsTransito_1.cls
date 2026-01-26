VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTransito_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit
' clsTransito

Private mEdad          As Integer
Private mFisico        As Byte
Private mMental        As Byte
Private mEmocional     As Byte
Private mEspiritual    As Byte
Private mEsencia       As Byte
Private mAnioPersonal  As Byte

' Metadatos / trazabilidad
Private mFuenteFisico      As String
Private mFuenteMental      As String
Private mFuenteEmocional   As String
Private mFuenteEspiritual  As String

Private mSistemaFonetico    As ModoFonetico
Private mSistemaCalculo     As ModoCalculo

Public Property Get Edad() As Integer
    Edad = mEdad
End Property
Public Property Let Edad(ByVal v As Integer)
    mEdad = v
End Property

Public Property Get Fisico() As Byte
    Fisico = mFisico
End Property
Public Property Let Fisico(ByVal v As Byte)
    mFisico = v
End Property

Public Property Get Mental() As Byte
    Mental = mMental
End Property
Public Property Let Mental(ByVal v As Byte)
    mMental = v
End Property

Public Property Get Emocional() As Byte
    Emocional = mEmocional
End Property
Public Property Let Emocional(ByVal v As Byte)
    mEmocional = v
End Property

Public Property Get Espiritual() As Byte
    Espiritual = mEspiritual
End Property
Public Property Let Espiritual(ByVal v As Byte)
    mEspiritual = v
End Property

Public Property Get Esencia() As Byte
    Esencia = mEsencia
End Property
Public Property Let Esencia(ByVal v As Byte)
    mEsencia = v
End Property

Public Property Get AnioPersonal() As Byte
    AnioPersonal = mAnioPersonal
End Property
Public Property Let AnioPersonal(ByVal v As Byte)
    mAnioPersonal = v
End Property

' Trazabilidad: de dónde viene cada tránsito
Public Property Get FuenteFisico() As String
    FuenteFisico = mFuenteFisico
End Property
Public Property Let FuenteFisico(ByVal v As String)
    mFuenteFisico = v
End Property

Public Property Get FuenteMental() As String
    FuenteMental = mFuenteMental
End Property
Public Property Let FuenteMental(ByVal v As String)
    mFuenteMental = v
End Property

Public Property Get FuenteEmocional() As String
    FuenteEmocional = mFuenteEmocional
End Property
Public Property Let FuenteEmocional(ByVal v As String)
    mFuenteEmocional = v
End Property

Public Property Get FuenteEspiritual() As String
    FuenteEspiritual = mFuenteEspiritual
End Property
Public Property Let FuenteEspiritual(ByVal v As String)
    mFuenteEspiritual = v
End Property

Public Property Get SistemaFonetico() As ModoFonetico
    SistemaFonetico = mSistemaFonetico
End Property
Public Property Let SistemaFonetico(ByVal v As ModoFonetico)
    mSistemaFonetico = v
End Property

Public Property Get SistemaCalculo() As ModoCalculo
    SistemaCalculo = mSistemaCalculo
End Property
Public Property Let SistemaCalculo(ByVal v As ModoCalculo)
    mSistemaCalculo = v
End Property

