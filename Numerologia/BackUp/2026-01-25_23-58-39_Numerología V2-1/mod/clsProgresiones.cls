VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsProgresiones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

Public IDPersona As Long
Public IDResultado As Long
Public IDProg As Long
    
Public ProgActual As Byte
Public ProgAnterior As Byte
Public ProgSiguiente As Byte

