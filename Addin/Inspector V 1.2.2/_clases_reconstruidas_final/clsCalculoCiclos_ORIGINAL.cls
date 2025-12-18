VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCalculoCiclos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ============================================================================
' Proyecto:     Sistema de Numerología Tradicional y Fonético
' Clase:        clsCalculoCiclos
' Descripción:  Calcula los Ciclos de Vida
' Autor:        Alba Salvá
' Fecha:        Diciembre 2025
' ============================================================================

'-----------------------------------------------------------------------------
'Variables
'-----------------------------------------------------------------------------
'Ciclo 1
Private objBase As clsCalculoBase
Private pEdadInicio1 As Integer
Private pEdadFin1 As Integer

'Ciclo 2
'Private objBase As clsCalculoBase
Private pEdadInicio2 As Integer
Private pEdadFin2 As Integer

'Ciclo 3
'Private objBase As clsCalculoBase
Private pEdadInicio3 As Integer
Private pEdadFin3 As Integer

'Ciclo 4
'Private objBase As clsCalculoBase
Private pEdadInicio4 As Integer
Private pEdadFin4 As Integer
Private pEsAplicable4 As Boolean


