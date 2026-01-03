Attribute VB_Name = "modCalculosNumerologicos"

Option Compare Database
Option Explicit

' ============================================================
'   MÓDULO DE CÁLCULOS NUMEROLÓGICOS
'   (Puente entre datos y motor pitagórico)
' ============================================================


' ------------------------------------------------------------
'   1. Cálculo genérico para cualquier número
' ------------------------------------------------------------
Public Function CalcularNumero(ByVal valor As Integer) As tResultado
    CalcularNumero = AnalizarNumero(valor)
End Function


' ------------------------------------------------------------
'   2. Suma de fecha (día + mes + año)
' ------------------------------------------------------------
Public Function SumarFecha(ByVal d As Integer, ByVal m As Integer, ByVal a As Integer) As Integer
    SumarFecha = d + m + a
End Function


' ------------------------------------------------------------
'   3. Cálculo numerológico de una fecha completa
' ------------------------------------------------------------
Public Function CalcularFecha(ByVal d As Integer, ByVal m As Integer, ByVal a As Integer) As tResultado
    Dim suma As Integer
    suma = SumarFecha(d, m, a)
    CalcularFecha = AnalizarNumero(suma)
End Function


' ------------------------------------------------------------
'   4. Cálculo numerológico de un nombre
'      (requiere tu motor fonético)
' ------------------------------------------------------------
Public Function CalcularNombre(ByVal Nombre As String) As tResultado
    Dim valor As Integer
    
    ' Motor fonético externo
    Call ParseNombre(Nombre)
    valor = ValorTotalFonemas()
    
    CalcularNombre = AnalizarNumero(valor)
End Function

