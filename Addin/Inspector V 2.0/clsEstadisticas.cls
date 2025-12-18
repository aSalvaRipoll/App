VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEstadisticas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

'===============================================================
' Clase: clsEstadisticas
' Representa estadísticas básicas sobre símbolos
'===============================================================

Private mTotal As Long
Private mUsados As Long
Private mNoUsados As Long

Public Property Get Total() As Long
    Total = mTotal
End Property
Public Property Let Total(val As Long)
    mTotal = val
End Property

Public Property Get Usados() As Long
    Usados = mUsados
End Property
Public Property Let Usados(val As Long)
    mUsados = val
End Property

Public Property Get NoUsados() As Long
    NoUsados = mNoUsados
End Property
Public Property Let NoUsados(val As Long)
    mNoUsados = val
End Property

''===============================================================
'' Clase: clsEstadisticas
'' Representa estadísticas básicas sobre símbolos
''===============================================================
'
'Private mTotal As Long
'Private mUsados As Long
'Private mNoUsados As Long
'
'Public Property Get Total() As Long
'    Total = mTotal
'End Property
'
'Public Property Let Total(val As Long)
'    mTotal = val
'End Property
'
'Public Property Get Usados() As Long
'    Usados = mUsados
'End Property
'
'Public Property Let Usados(val As Long)
'    mUsados = val
'End Property
'
'Public Property Get NoUsados() As Long
'    NoUsados = mNoUsados
'End Property
'
'Public Property Let NoUsados(val As Long)
'    mNoUsados = val
'End Property
'
