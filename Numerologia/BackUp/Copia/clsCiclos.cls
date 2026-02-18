VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCiclos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
 '====================================================================
 'Clase generada automáticamente desde la tabla: tbuCiclos
 '====================================================================

Option Compare Database
Option Explicit

' Public Class clsCiclos

    Public idCiclo As Long      ' Campo: IDCiclo
    Public IDResultado As Long  ' Campo: IDResultado
    Public IDPersona As Long    ' Campo: IDPersona
    Public IDFonetica As Long
    
    Public NumCiclos As Byte    ' Campo: NumCiclos
    Public MetodoCiclos As Byte ' Campo: MetodoCiclos
    
    Public Ciclo1 As Byte   ' Campo: Ciclo1
    Public EdadIni1 As Byte ' Campo: EdadIni1
    Public EdadFin1 As Byte ' Campo: EdadFin1
    
    Public Ciclo2 As Byte   ' Campo: Ciclo2
    Public EdadIni2 As Byte ' Campo: EdadIni2
    Public EdadFin2 As Byte ' Campo: EdadFin2
    
    Public Ciclo3 As Byte   ' Campo: Ciclo3
    Public EdadIni3 As Byte ' Campo: EdadIni3
    Public EdadFin3 As Byte ' Campo: EdadFin3
    
    Public Ciclo4 As Byte   ' Campo: Ciclo4
    Public EdadIni4 As Byte ' Campo: EdadIni4
    Public EdadFin4 As Byte ' Campo: EdadFin4

' End Class
' ====================================================================

Private Sub Class_Initialize()

    NumCiclos = 0
    MetodoCiclos = 0

    Ciclo1 = 0
    Ciclo2 = 0
    Ciclo3 = 0
    Ciclo4 = 0
    
    EdadIni1 = 0
    EdadIni2 = 0
    EdadIni3 = 0
    EdadIni4 = 0
    
    EdadFin1 = 0
    EdadFin2 = 0
    EdadFin3 = 0
    EdadFin4 = 0
    
End Sub
