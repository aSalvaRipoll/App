VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCronometro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

' Declaraciones para compatibilidad con 32 y 64 bits
#If VBA7 Then
    Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
    Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
#Else
    Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
    Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
#End If

Private Tstart As Currency
Private Tfreq As Currency
Private TiempoAcumulado As Double
Private CronoActivo As Boolean

Private Laps As Collection
Private UltimoLap As Double

' Tiempo en milisegundos
Public Property Get TiempoMS() As Double
    TiempoMS = Tiempo * 1000#
End Property

' Tiempo en microsegundos
Public Property Get TiempoUS() As Double
    TiempoUS = Tiempo * 1000000#
End Property

' Tiempo en nanosegundos
Public Property Get TiempoNS() As Double
    TiempoNS = Tiempo * 1000000000#
End Property

' Formato automático según magnitud
Public Property Get TiempoAuto() As String
    Dim t As Double
    t = Tiempo

    Select Case t
        Case Is >= 1
            FormatoInteligente = Format(t, "0.000000") & " s"
        Case Is >= 0.001
            FormatoInteligente = Format(t * 1000#, "0.000") & " ms"
        Case Is >= 0.000001
            FormatoInteligente = Format(t * 1000000#, "0.000") & " µs"
        Case Else
            FormatoInteligente = Format(t * 1000000000#, "0.000") & " ns"
    End Select
End Property

' Constructor
Private Sub Class_Initialize()
    Set Laps = New Collection
    UltimoLap = 0
    ' Si quieres que el cronómetro empiece automáticamente:
    ' Inicio
End Sub

' Inicia el cronómetro
Public Sub Inicio()
    If Not CronoActivo Then
        QueryPerformanceFrequency Tfreq
        QueryPerformanceCounter Tstart
        CronoActivo = True
    End If
End Sub

' Detiene el cronómetro y acumula el tiempo
Public Sub Parar()
    Dim Tend As Currency
    Dim Intervalo As Double

    If CronoActivo Then
        QueryPerformanceCounter Tend
        Intervalo = (Tend - Tstart) / Tfreq
        TiempoAcumulado = TiempoAcumulado + Intervalo
        CronoActivo = False
    End If
End Sub

' Registra un tiempo parcial (lap)
Public Function Lap() As Double
    Dim Tend As Currency
    Dim Intervalo As Double
    Dim LapTime As Double

    If CronoActivo Then
        QueryPerformanceCounter Tend
        Intervalo = (Tend - Tstart) / Tfreq
        LapTime = TiempoAcumulado + Intervalo - UltimoLap
    Else
        LapTime = TiempoAcumulado - UltimoLap
    End If

    UltimoLap = UltimoLap + LapTime
    Laps.Add LapTime

    Lap = LapTime
End Function

' Devuelve el tiempo acumulado total
Public Property Get Tiempo() As Double
    Dim Tend As Currency
    Dim Intervalo As Double

    If CronoActivo Then
        QueryPerformanceCounter Tend
        Intervalo = (Tend - Tstart) / Tfreq
        Tiempo = TiempoAcumulado + Intervalo
    Else
        Tiempo = TiempoAcumulado
    End If
End Property

' Devuelve la colección de laps
Public Property Get ListaLaps() As Collection
    Set ListaLaps = Laps
End Property

' Reinicia todo
Public Sub Reiniciar()
    TiempoAcumulado = 0
    CronoActivo = False
    UltimoLap = 0
    Set Laps = New Collection
End Sub


