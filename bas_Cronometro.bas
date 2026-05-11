Attribute VB_Name = "bas_Cronometro"

Option Compare Database
Option Explicit

Sub TestCronometro()

    Dim C As New clsCronometro
    Dim i As Long

    C.Inicio
    For i = 1 To 2000000: Next i
    Debug.Print "Lap 1: "; C.Lap

    For i = 1 To 2000000: Next i
    Debug.Print "Lap 2: "; C.Lap

    MsgBox "Pausa del usuario…"

    For i = 1 To 2000000: Next i
    C.Parar

    Debug.Print "Tiempo total: "; C.Tiempo

    Dim L As Variant
    For Each L In C.ListaLaps
        Debug.Print "Lap registrado: "; L
    Next L

End Sub

