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
'Public Enum tipoCalculo
'    CaminoVida = 1
'    Destino = 2
'    Alma = 3
'    Personalidad = 4
'    Madurez = 5
'    AnoPersonal = 6
'    MesPersonal = 7
'    DiaPersonal = 8
'    Ciclo1 = 9
'    Ciclo2 = 10
'    Ciclo3 = 11
'    Ciclo4 = 12
'    Pinaculo1 = 13
'    Pinaculo2 = 14
'    Pinaculo3 = 15
'    Pinaculo4 = 16
'    Desafio1 = 17
'    Desafio2 = 18
'    Desafio3 = 19
'    Desafio4 = 20
'    NumeroExpresion = 21
'    NumeroPoder = 22
'    numeroFaltante = 23
'    NumeroDominante = 24
'    PrimeraLetra = 25
'    PrimeraVocal = 26
'    PrimeraConsonante = 27
'    RespuestaSubconsciente = 28
'    PlanoExpresion = 29
'End Enum

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

Public Enum eIdiomaFonetico
    idOtros = 0
    idCastellano = 1
    idCatala = 2
    idMallorquin = 3
    idValenciano = 4
    idEuskera = 5
    idGalego = 6
'    idPortugues = 7
    idPortuguesEU = 8
    idPortuguesBR = 9
    idFrances = 10
    idIngles = 11
'    idInglesUSA = 12
'    idEnUsaAfro = 13
End Enum

'Tipo equivalencia Para clasificar la equivalencia:
'
'ETIM -> equivalencia etimológica
'
'FON -> adaptación fonética
'
'PREST -> préstamo lingüístico
'
'EVOL -> evolución histórica
'
'CULT -> equivalencia cultural moderna
'
'VAR -> variante ortográfica
'
'NOEQ -> falso equivalente (si quieres marcarlo)

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

Public Type tAppVersion
    vMajor As Integer
    vMinor As Integer
    vVersion As Integer
End Type


Public colIdiomas As Collection
Public colFonemas As Collection

'Public DicFonemas As Scripting.Dictionary

'Variables públicas para traspaso de información entre formularios
Public IdiomaActual As clsIdioma ' <-- ESTA es la que recibirá el valor
Public CampoDestino As String
Public IdiomaSeleccionado As clsIdioma  '<-- Esta devuelve el valor
Public Const RutaImagen As String = "N:\Numerologia\App\Numerologia\img\PNG\"

Public Persona As clsPersona
Public Fonetica As clsFonetica
'Public Resultado As clsResultado
'Public Inclusion As clsInclusion
'Public PinaDes As clsPinaDes
'Public Ciclos As clsCiclos
'Public Progres As clsProgresiones
'Public Transit As clsTransitos

Public AppVersion As tAppVersion

#If 1 = 2 Then
    Dim mfFonetico, mfTradicional, mcClasico, mcModerno, ccFijo, ccClasico, ccModerno, mtTradicional, mtJavane
    Dim idOtros, idCastellano, idCatala, idMallorquin, idValenciano, idEuskera, idGalego, idPortugues, idPortuguesEU, idPortuguesBR, idFrances, idIngles
#End If


Sub InitApp()
    With AppVersion
        .vMajor = GetProperty("vMajor", 0)
        .vMinor = GetProperty("vMinor", 0)
        .vVersion = GetProperty("vVersion", 0)
    End With
End Sub


Public Function GetProperty(strName As String, strDefault As String) _
   As Variant
   
   Dim dbs As Object
'Created by Helen Feddema 31-Mar-2017
'Modified by Helen Feddema 31-Mar-2017
'Called from various procedures
On Error GoTo ErrorHandler
   
   'Attempt to get the value of the specified property
   Set dbs = CurrentDb
   GetProperty = dbs.Properties(strName).Value
ErrorHandlerExit:
   Exit Function
ErrorHandler:
   If Err.Number = 3270 Then
      'The property was not found; use default value
      GetProperty = strDefault
      Resume Next
   Else
      MsgBox "Error No: " & Err.Number _
         & " in GetProperty procedure; " _
         & "Description: " & Err.Description
      Resume ErrorHandlerExit
   End If
End Function


