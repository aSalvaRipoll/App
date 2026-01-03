VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsResultado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

Public IDResultado As Long
Public IDPersona As Long
Public IDFonetica As Long
Public Activo As Boolean

' ============================
'  NÚMEROS PRINCIPALES
' ============================
Public NumeroCaminoVida As String ' Fecha de nacimiento
Public NumeroDestino As String ' Expresión (todas las letras)
Public NumeroMadurez As String ' CaminoVida + Destino
Public NumeroAlma As String ' Vocales
Public NumeroPersonalidad As String ' Consonantes

' ============================
'  NÚMEROS DERIVADOS
' ============================
'Public NumeroCaminoVida As String
'Public NumeroMadurez As String
Public NumeroPoder As String
Public NumeroRespuestaSubconsciente As String
Public NumeroDiaNacimiento As Byte
Public NumeroAñoPersonal As String
Public NumeroEdadPersonal As String

' ============================
'  DEUDAS KÁRMICAS
' ============================
Public DeudasKarmicas As Collection   ' {13, 14, 16, 19}

' ============================
'  PINÁCULOS Y DESAFÍOS
' ============================
'Public Pinaculos As Collection        ' {P1, P2, P3, P4}
'Public Desafios As Collection         ' {D1, D2, D3, D4}

' ============================
'  CICLOS
' ============================
'Public Ciclos As Collection           ' {C1, C2, C3}

' ============================
'  LETRAS Y SÍMBOLOS
' ============================
Public PrimeraLetra As String
Public PrimeraVocal As String
Public PrimeraConsonante As String
Public PiedraAngular As String
Public PiedraToque As String

' ============================
'  LISTAS ESPECIALES
' ============================
'Public Ausentes As Collection         ' números ausentes
'Public Dominantes As Collection       ' números dominantes

' ============================
'  PLANOS DE EXPRESIÓN
' ============================
Public PlanoFisico As String
Public PlanoEmocional As String
Public PlanoMental As String
Public PlanoIntuitivo As String

' ============================
'  TAROT
' ============================
Public SistemaTarot As String         ' Marseille, Rider, GD...
'Public ArcanosPrincipales As Collection   ' {Destino?Arcano, Alma?Arcano...}

