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

' ============================
'  DATOS IDENTIFICATIVOS
' ============================
' Identificadores
Public IDResultado As Long
Public IDPersona As Long
Public IDFonetica As Long
Public Activo As Boolean
Public VersionMotor As String

'Datos para Cálculos auxiliares
Public SistemaCalculo As ModoCalculo
Public NumCiclos As Byte
Public SistemaCiclos As ModoCiclos
Public SistemaTarot As ModoTarot

' ============================
'  NÚMEROS PRINCIPALES
' ============================
Public NumeroCaminoVida As String ' Fecha de nacimiento
Public NumeroDestino As String ' Expresión (todas las letras)
Public NumeroAlma As String ' Vocales
Public NumeroPersonalidad As String ' Consonantes
Public NumeroDiaNacimiento As Byte

' ============================
'  NÚMEROS DERIVADOS
' ============================
Public NumeroMadurez As String
Public NumeroPoder As String
Public RespuestaSubconsciente As Byte

' ============================
'  PLANOS DE EXPRESIÓN
' ============================
Public PlanoFisico As String
Public PlanoEmocional As String
Public PlanoMental As String
Public PlanoIntuitivo As String

' ============================
'  ÁRBOL DE VIDA
' ============================
Public ArbolPaterno As String
Public ArbolMaterno As String
Public ArbolDeVida As String

' ============================
'  LETRAS Y SÍMBOLOS
' ============================
Public PrimeraLetra As String
Public PrimeraVocal As String
Public PrimeraConsonante As String
Public PiedraAngular As String
Public PiedraToque As String

' ============================
'  TEMPORALES
' ============================
Public Edad As Byte             ' Edad de la persona
Public AnioPersonal As String   ' Año Personal
Public EdadPersonal As String   ' Edad Personal
Public Esencia As String        ' Número de esencia

' ============================
'  LISTAS ESPECIALES
' ============================
Public Ausentes As String        ' números ausentes
Public Dominantes As String      ' números dominantes

' ============================
'  FECHA DEL CÁLCULO
' ============================
Public FechaCalculo As Date



' ============================
'  Resto Temporales
' ============================


' ============================
'  CICLOS
' ============================
'Public Ciclos As Collection    ' Tabla y clase
' ============================
'  PINÁCULOS Y DESAFÍOS
' ============================
'Public Pinaculos As Collection   'Tabla y Clase     ' {P1, P2, P3, P4}
'Public Desafios As Collection    'Tabla y Clase     ' {D1, D2, D3, D4}

' ============================
'  TAROT
' ============================
        ' Marseille, Rider, GD...
'Public ArcanosPrincipales As Collection   ' {Destino?Arcano, Alma?Arcano...}

' ESTOS YA NO SE HACEN
' ============================
'  DEUDAS KÁRMICAS
' ============================
'Public DeudasKarmicas As Collection   ' {13, 14, 16, 19}


