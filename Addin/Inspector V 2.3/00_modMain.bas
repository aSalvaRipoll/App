Attribute VB_Name = "00_modMain"

Option Compare Database
Option Explicit

'===============================================================
' 00_Main – Punto de entrada del InspectorVBA
' Solo delega en el Core y en los subsistemas
'===============================================================

Public Sub Inspector_Inicializar()
    Call Core_Inicializar
End Sub

Public Function Inspector_Analizar() As EstadoAnalisis
    Inspector_Analizar = Core_Analizar()
End Function

Public Function Inspector_Reparar() As EstadoReparacion
    Inspector_Reparar = Core_Reparar()
End Function

Public Function Inspector_Resumen() As String
    Inspector_Resumen = Core_Resumen()
End Function

Public Sub Inspector_Reset(Optional reiniciarMotor As Boolean = False)
    Call Core_Reset(reiniciarMotor)
End Sub

