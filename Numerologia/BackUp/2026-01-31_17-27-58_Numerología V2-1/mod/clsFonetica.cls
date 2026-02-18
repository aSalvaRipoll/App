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
'   CLASE FONEMAS — DTO PURO (DATOS DE CONVERSIÓN)
' ============================================================

Option Compare Database
Option Explicit

' --- Identificación ---
Public IDFonetica As Long
Public IDPersona As Long

' --- Modo de procesamiento ---
' 1 = Tradicional clásico (sin conversión)
' 2 = Fonético (con conversión)
' 3 = Tradicional moderno (sin conversión) No existe
'Public ModoFonetico As Byte
Public SistemaFonetico As ModoFonetico

'' --- Reglas aplicadas ---  No se usan, embebidas en reglas.
'Public UsarHmuda As Boolean
'Public UsarUmuda As Boolean

' --- Idiomas (abreviados: ES, EN, EU, FR...) ---
Public IdiomaNombre As Byte
Public IdiomaApe1 As Byte
Public IdiomaApe2 As Byte

' --- Resultados fonéticos (texto final) ---
Public FonNombre As String
Public FonApe1 As String
Public FonApe2 As String

' --- Datos originales (para trazabilidad) --- No necesarios, están en clsPersona
Public FechaNacimiento As Date
'Public NombreOriginal As String
'Public Ape1Original As String
'Public Ape2Original As String

' --- Gestión ---
Public FechaCalculo As Date
Public FechaModificacion As Date
Public Activo As Boolean ' Para permitir borrado lógico
Public VersionMotor As String
