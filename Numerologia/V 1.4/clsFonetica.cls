VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFonetica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' ============================================================
'   CLASE FONETICA — CONFIGURACIÓN + RESULTADOS
' ============================================================

Option Compare Database
Option Explicit

' --- Identificación ---
Private mIDFonetica As Long
Private mIDPersona As Long

' --- Modo de procesamiento ---
' 0 = Tradicional (sin conversión)
' 1 = Fonético (con conversión)
Private mModoFonetico As Byte

' --- Reglas ---
Private mUsarHmuda As Boolean
Private mUsarUmuda As Boolean

' --- Idiomas aplicados (OBJETOS) ---
Private mIdiomaNombre As clsIdioma
Private mIdiomaApe1 As clsIdioma
Private mIdiomaApe2 As clsIdioma

' --- Resultados fonéticos ---
Private mFonNombre As String
Private mFonApe1 As String
Private mFonApe2 As String

' --- Datos originales (para modo tradicional) ---
Private mNombreOriginal As String
Private mApe1Original As String
Private mApe2Original As String

' --- Gestión ---
Private mFechaCalculo As Date


' ============================================================
'   PROPIEDADES PÚBLICAS
' ============================================================

' --- Identificación ---
Public Property Get IDFonetica() As Long: IDFonetica = mIDFonetica: End Property
Public Property Let IDFonetica(ByVal v As Long): mIDFonetica = v: End Property

Public Property Get IDPersona() As Long: IDPersona = mIDPersona: End Property
Public Property Let IDPersona(ByVal v As Long): mIDPersona = v: End Property


' --- Modo ---
Public Property Get ModoFonetico() As Byte: ModoFonetico = mModoFonetico: End Property
Public Property Let ModoFonetico(ByVal v As Byte): mModoFonetico = v: End Property


' --- Reglas ---
Public Property Get UsarHmuda() As Boolean: UsarHmuda = mUsarHmuda: End Property
Public Property Let UsarHmuda(ByVal v As Boolean): mUsarHmuda = v: End Property

Public Property Get UsarUmuda() As Boolean: UsarUmuda = mUsarUmuda: End Property
Public Property Let UsarUmuda(ByVal v As Boolean): mUsarUmuda = v: End Property


' --- Idiomas ---
Public Property Get IdiomaNombre() As clsIdioma: Set IdiomaNombre = mIdiomaNombre: End Property
Public Property Set IdiomaNombre(ByVal v As clsIdioma): Set mIdiomaNombre = v: End Property

Public Property Get IdiomaApe1() As clsIdioma: Set IdiomaApe1 = mIdiomaApe1: End Property
Public Property Set IdiomaApe1(ByVal v As clsIdioma): Set mIdiomaApe1 = v: End Property

Public Property Get IdiomaApe2() As clsIdioma: Set IdiomaApe2 = mIdiomaApe2: End Property
Public Property Set IdiomaApe2(ByVal v As clsIdioma): Set mIdiomaApe2 = v: End Property


' --- Resultados fonéticos ---
Public Property Get FonNombre() As String: FonNombre = mFonNombre: End Property
Public Property Let FonNombre(ByVal v As String): mFonNombre = v: End Property

Public Property Get FonApe1() As String: FonApe1 = mFonApe1: End Property
Public Property Let FonApe1(ByVal v As String): mFonApe1 = v: End Property

Public Property Get FonApe2() As String: FonApe2 = mFonApe2: End Property
Public Property Let FonApe2(ByVal v As String): mFonApe2 = v: End Property


' --- Datos originales ---
Public Property Get NombreOriginal() As String: NombreOriginal = mNombreOriginal: End Property
Public Property Let NombreOriginal(ByVal v As String): mNombreOriginal = v: End Property

Public Property Get Ape1Original() As String: Ape1Original = mApe1Original: End Property
Public Property Let Ape1Original(ByVal v As String): mApe1Original = v: End Property

Public Property Get Ape2Original() As String: Ape2Original = mApe2Original: End Property
Public Property Let Ape2Original(ByVal v As String): mApe2Original = v: End Property


' --- Gestión ---
Public Property Get FechaCalculo() As Date: FechaCalculo = mFechaCalculo: End Property
Public Property Let FechaCalculo(ByVal v As Date): mFechaCalculo = v: End Property

