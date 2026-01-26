VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPinaDes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ====================================================================
' Clase generada automáticamente desde la tabla: tbuPinaDes
' ====================================================================

Option Compare Database
Option Explicit

' Public Class clsPinaDes
' DTO para almacenar los Pináculos y Desafíos numerológicos,
' incluyendo fechas exactas de inicio y fin.

    Public IDPinaDes As Long ' Campo: IDPinaDes
    Public IDResultado As Long ' Campo: IDResultado
    Public IDPersona As Long ' Campo: IDPersona
    Public IDFonetica As Long
    
    Public Pina1 As String ' Campo: Pina1
    Public Pina2 As String ' Campo: Pina2
    Public Pina3 As String ' Campo: Pina3
    Public Pina4 As String ' Campo: Pina4
    
    Public Desa1 As String ' Campo: Desa1
    Public Desa2 As String ' Campo: Desa2
    Public Desa3 As String ' Campo: Desa3
    Public Desa4 As String ' Campo: Desa4
    
    Public EdadIni1 As Date ' Campo: fIni1
    Public EdadIni2 As Date ' Campo: fIni2
    Public EdadIni3 As Date ' Campo: fIni3
    Public EdadIni4 As Date ' Campo: fIni4
    
    Public EdadFin1 As Date ' Campo: fFin1
    Public EdadFin2 As Date ' Campo: fFin2
    Public EdadFin3 As Date ' Campo: fFin3
    Public EdadFin4 As Date ' Campo: fFin4
    
' End Class
' ====================================================================


Private Sub Class_Initialize()

    Pina1 = ""
    Pina2 = ""
    Pina3 = ""
    Pina4 = ""

    Desa1 = ""
    Desa2 = ""
    Desa3 = ""
    Desa4 = ""

    EdadIni1 = 0
    EdadIni2 = 0
    EdadIni3 = 0
    EdadIni4 = 0

    EdadFin1 = 0
    EdadFin2 = 0
    EdadFin3 = 0
    EdadFin4 = 0

End Sub

