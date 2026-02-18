VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTransito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

' =============================================================================
'  clsTransito
'  DTO completo para un tránsito anual.
'
'  Contiene:
'   - Identificadores para persistencia (persona, resultado, tránsito, fonética)
'   - Datos temporales (orden, año, edad)
'   - Valores de tránsitos (físico, mental, emocional, espiritual)
'   - Letras/fonemas que generan cada tránsito
'   - Valores superiores (Esencia, Año Personal)
'   - Metadatos del cálculo (modo fonético, modo cálculo)
'   - Trazabilidad interna (valor bruto, ciclo de secuencia)
'
'  Este DTO está diseñado para:
'   - informes
'   - exportación
'   - depuración
'   - persistencia en tablas
'   - interpretación avanzada
' =============================================================================

' -------------------------
' Identificadores
' -------------------------
Public IDPersona        As Long
Public IDResultado      As Long
Public idTransito       As Long
Public IDFonetica       As Long

' -------------------------
' Datos temporales
' -------------------------
Public orden            As Byte     ' posición dentro de la tabla (1–11)
Public anio             As Integer  ' año calendario
Public Edad             As Byte     ' edad correspondiente
Public EdadFin          As Byte
Public EsActual         As Boolean

' -------------------------
' Valores reducidos (1–9, 11, 22)
' -------------------------
Public Fisico           As Byte
Public Mental           As Byte
Public Emocional        As Byte
Public Espiritual       As Byte

' -------------------------
' Letras/fonemas que generan cada tránsito
' -------------------------
Public LetraFisico      As String
Public LetraMental      As String
Public LetraEmocional   As String
Public LetraEspiritual  As String

' -------------------------
' Valores superiores
' -------------------------
Public Esencia          As Byte
Public AnioPersonal     As Byte

' -------------------------
' Metadatos del cálculo
' -------------------------
Public SistemaFonetico  As ModoFonetico
Public SistemaCalculo   As ModoCalculo

' -------------------------
' Trazabilidad interna
' -------------------------
Public ValorFisicoBruto     As Byte
Public ValorMentalBruto     As Byte
Public ValorEmocionalBruto  As Byte
Public ValorEspiritualBruto As Byte

Public CicloFisico          As Byte
Public CicloMental          As Byte
Public CicloEmocional       As Byte
Public CicloEspiritual      As Byte

' =============================================================================
' Inicialización por defecto
' =============================================================================
Private Sub Class_Initialize()

    orden = 0
    anio = 0
    Edad = 0
    EdadFin = 0
    EsActual = False
    
    Fisico = 0
    Mental = 0
    Emocional = 0
    Espiritual = 0

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

