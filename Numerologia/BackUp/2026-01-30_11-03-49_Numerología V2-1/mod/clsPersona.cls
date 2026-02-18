VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPersona"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' ============================================================
'   CLASE PERSONA — ESTRUCTURA BASE
'   (Datos personales, fonéticos y seleccionados análisis)
' ============================================================

Option Compare Database
Option Explicit

' ============================================================
'   CLASE PERSONA — ESTRUCTURA CORRECTA
' ============================================================

' --- Identificación ---
Private mIDPersona As Long

' --- Datos personales ---
Private mNombre As String
Private mApe1 As String
Private mApe2 As String
Private mFechaNacimiento As Date
Private mGenero As Byte

' --- Idiomas (OBJETOS, no IDs) ---
Private mIdiomaNombre As clsIdioma
Private mIdiomaApe1 As clsIdioma
Private mIdiomaApe2 As clsIdioma

' --- Fonética ---
'Private mFonNombre As String
'Private mFonApe1 As String
'Private mFonApe2 As String

' --- Parámetros de análisis ---
'Private mSistema As Long
'Private mTarot As Long
'Private mUsarHmuda As Boolean
'Private mUsarUmuda As Boolean

' --- Gestión ---
Private mFechaAlta As Date
Private mFechaModificacion As Date


' ============================================================
'   PROPIEDADES PÚBLICAS
' ============================================================

' --- Identificación ---
Public Property Get ID_Persona() As Long: ID_Persona = mIDPersona: End Property
Public Property Let ID_Persona(ByVal v As Long): mIDPersona = v: End Property

' --- Datos personales ---
Public Property Get Nombre() As String: Nombre = mNombre: End Property
Public Property Let Nombre(ByVal v As String): mNombre = v: End Property

Public Property Get Ape1() As String: Ape1 = mApe1: End Property
Public Property Let Ape1(ByVal v As String): mApe1 = v: End Property

Public Property Get Ape2() As String: Ape2 = mApe2: End Property
Public Property Let Ape2(ByVal v As String): mApe2 = v: End Property

Public Property Get FechaNacimiento() As Date: FechaNacimiento = mFechaNacimiento: End Property
Public Property Let FechaNacimiento(ByVal v As Date): mFechaNacimiento = v: End Property

Public Property Get ID_Genero() As Byte: ID_Genero = mGenero: End Property
Public Property Let ID_Genero(ByVal v As Byte): mGenero = v: End Property


' --- Idiomas (OBJETOS) ---
Public Property Get IdiomaNombre() As clsIdioma
    Set IdiomaNombre = mIdiomaNombre
End Property

Public Property Set IdiomaNombre(ByVal v As clsIdioma)
    Set mIdiomaNombre = v
End Property

Public Property Get IdiomaApe1() As clsIdioma
    Set IdiomaApe1 = mIdiomaApe1
End Property

Public Property Set IdiomaApe1(ByVal v As clsIdioma)
    Set mIdiomaApe1 = v
End Property

Public Property Get IdiomaApe2() As clsIdioma
    Set IdiomaApe2 = mIdiomaApe2
End Property

Public Property Set IdiomaApe2(ByVal v As clsIdioma)
    Set mIdiomaApe2 = v
End Property


' --- Fonética ---
'Public Property Get FonNombre() As String: FonNombre = mFonNombre: End Property
'Public Property Let FonNombre(ByVal v As String): mFonNombre = v: End Property
'
'Public Property Get FonApe1() As String: FonApe1 = mFonApe1: End Property
'Public Property Let FonApe1(ByVal v As String): mFonApe1 = v: End Property
'
'Public Property Get FonApe2() As String: FonApe2 = mFonApe2: End Property
'Public Property Let FonApe2(ByVal v As String): mFonApe2 = v: End Property


' --- Parámetros ---
'Public Property Get Sistema() As Long: Sistema = mSistema: End Property
'Public Property Let Sistema(ByVal v As Long): mSistema = v: End Property
'
'Public Property Get Tarot() As Long: Tarot = mTarot: End Property
'Public Property Let Tarot(ByVal v As Long): mTarot = v: End Property
'
'Public Property Get UsarHmuda() As Boolean: UsarHmuda = mUsarHmuda: End Property
'Public Property Let UsarHmuda(ByVal v As Boolean): mUsarHmuda = v: End Property
'
'Public Property Get UsarUmuda() As Boolean: UsarUmuda = mUsarUmuda: End Property
'Public Property Let UsarUmuda(ByVal v As Boolean): mUsarUmuda = v: End Property


' --- Gestión ---
Public Property Get FechaAlta() As Date: FechaAlta = mFechaAlta: End Property
Public Property Let FechaAlta(ByVal v As Date): mFechaAlta = v: End Property

Public Property Get FechaModificacion() As Date: FechaModificacion = mFechaModificacion: End Property
Public Property Let FechaModificacion(ByVal v As Date): mFechaModificacion = v: End Property

