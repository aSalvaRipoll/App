Attribute VB_Name = "modMotor_V_2_2_Aux"
Option Compare Database
Option Explicit

Public Function MarcarTonicas(ByVal silabas As String, ByVal tonicas As String) As String

    Dim partes() As String
    Dim idx() As String
    Dim i As Long, j As Long

    partes = Split(silabas, "-")
    idx = Split(tonicas, ",")

    For j = LBound(idx) To UBound(idx)
        i = CLng(idx(j)) - 1
        If i >= 0 And i <= UBound(partes) Then
            partes(i) = "*" & partes(i) & "*"
        End If
    Next j

    MarcarTonicas = Join(partes, "-")

End Function

