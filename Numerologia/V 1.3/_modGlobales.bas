Attribute VB_Name = "_modGlobales"
Option Compare Database
Option Explicit

' _modGlobales
' Módulo que contiene todos los elementos globales de la aplicación

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

Public colIdiomas As Collection

Public DicFonemas As Scripting.Dictionary

Public DicNombres As Scripting.Dictionary
Public DicApellidos As Scripting.Dictionary

Public ColNombres As Collection
Public ColApellidos As Collection


' ============================================================
'   TIPO DE RESULTADO NUMEROLÓGICO
' ============================================================
Public Type tResultado
    Cadena As String      ' Presentación final (ej: "128/11/2")
    Inicial As Integer    ' Valor inicial bruto
    Medio As Byte         ' Primera reducción a 2 dígitos (si existe)
    Final As Byte         ' Reducción final a 1 dígito
    Maestro As Byte       ' 11,22,33,44 o 0
    Karmico As Byte       ' 13,14,16,19 o 0
End Type

Public Type SalidaDatos
    Vocales As String
    Consonantes As String
    Completo As String
End Type

'Variables públicas para traspaso de información entre formularios
Public IdiomaActual As clsIdioma ' <-- ESTA es la que recibirá el valor
Public CampoDestino As String
Public IdiomaSeleccionado As clsIdioma  '<-- Esta devuelve el valor

Public Persona As clsPersona
Public Fonetica As clsFonetica

