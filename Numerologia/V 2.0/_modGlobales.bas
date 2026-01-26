Attribute VB_Name = "_modGlobales"

Option Compare Database
Option Explicit

' _modGlobales
' Módulo que contiene todos los elementos globales de la aplicación

Public Enum ResGuardar
    RgCancelado = 0
    RgCreado
    RgActualizado
    RgGuardado
End Enum
    


' Tipos de cálculo
Public Enum tipoCalculo
    CaminoVida = 1
    Destino = 2
    Alma = 3
    Personalidad = 4
    Madurez = 5
    AnoPersonal = 6
    MesPersonal = 7
    DiaPersonal = 8
    Ciclo1 = 9
    Ciclo2 = 10
    Ciclo3 = 11
    Ciclo4 = 12
    Pinaculo1 = 13
    Pinaculo2 = 14
    Pinaculo3 = 15
    Pinaculo4 = 16
    Desafio1 = 17
    Desafio2 = 18
    Desafio3 = 19
    Desafio4 = 20
    NumeroExpresion = 21
    NumeroPoder = 22
    numeroFaltante = 23
    NumeroDominante = 24
    PrimeraLetra = 25
    PrimeraVocal = 26
    PrimeraConsonante = 27
    RespuestaSubconsciente = 28
    PlanoExpresion = 29
End Enum

'Public Enum idioma
'    iEspañol = 1
'    iCatala
'    iEuskera
'    iGalego
'End Enum

'-------------------------------------------------------------------------------
' Revisar tbmModos si se cambia algo de estas enumeraciones
'-------------------------------------------------------------------------------
Public Enum ModoFonetico
    mfFonetico = 1
    mfTradicional = 2
End Enum

Public Enum ModoCalculo
    mcClasico = 1
    mcModerno = 2
End Enum

Public Enum ModoCiclos
    ccFijo = 0
    ccClasico = 1
    ccModerno = 2
End Enum


Public Enum ModoTarot
    mtTradicional = 1
    mtJavane = 2
End Enum

'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------

Public Type tInterCadena
    original As String
'    Intermedio As String
    Número As Byte
'    Final As String
    esMaestro As Boolean
    esKarmico As Boolean
End Type


Public colIdiomas As Collection

Public DicFonemas As Scripting.Dictionary

Public DicNombres As Scripting.Dictionary
Public DicApellidos As Scripting.Dictionary

Public DicNom As Scripting.Dictionary
Public DicApe As Scripting.Dictionary

Public ColNombres As Collection
Public ColApellidos As Collection


' ============================================================
'   TIPO DE RESULTADO NUMEROLÓGICO
' ============================================================
Public Type tResultado
    cadena As String      ' Presentación final (ej: "128/11/2")
    Inicial As Integer    ' Valor inicial bruto
    Medio As Byte         ' Primera reducción a 2 dígitos (si existe)
    Final As Byte         ' Reducción final a 1 dígito
    Maestro As Byte       ' 11,22,33,44 o 0
    Karmico As Byte       ' 13,14,16,19 o 0
End Type

'Public Type SalidaDatos
'    Vocales As String
'    Consonantes As String
'    Completo As String
'End Type

Public Type tAcumuladores
    Vocales As Integer
    Consonantes As Integer
    Completo As Integer
End Type



'Variables públicas para traspaso de información entre formularios
Public IdiomaActual As clsIdioma ' <-- ESTA es la que recibirá el valor
Public CampoDestino As String
Public IdiomaSeleccionado As clsIdioma  '<-- Esta devuelve el valor


Public Persona As clsPersona
Public Fonetica As clsFonetica
'Public Resultado As clsResultado
'Public Inclusion As clsInclusion
'Public PinaDes As clsPinaDes
'Public Ciclos As clsCiclos
'Public Progres As clsProgresiones
'Public Transit As clsTransitos

#If 1 = 2 Then
    Dim mfFonetico, mfTradicional, mcClasico, mcModerno, ccFijo, ccClasico, ccModerno, mtTradicional, mtJavane
#End If

