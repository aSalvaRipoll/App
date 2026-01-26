VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFonema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

Public NumOrden As Integer
Public idFonema As Byte
Public GrafemaOri As String
Public ASCII As String
Public Valor As Byte

' Fonética
Public EsVocal As Boolean
Public DatosFonéticos As String

