
Public Type Resultados
    Cadena As String
    Inicial As Byte
    Medio As Byte
    Final As Byte
    Maestro As Byte
    Karma As Byte
End Type




Function NumerCalc(ByVal xNum) As Resultados

    Dim NumCnv, arrNum
    Dim i

    NumCnv = ""

    NumerCalc.Inicial = 0
    NumerCalc.Maestro = 0
    NumerCalc.Final = 0

    If Trim$(xNum) = "" Then
        Exit Function
    End If

    On Error GoTo NumerCalc_Error

    xNum = Replace(xNum, " ", "")

    While Len(xNum) > 2    'xNum > 99 '78
        xNum = SumaCadena(xNum)
    Wend

    NumerCalc.Cadena = xNum

    While xNum > 9
        xNum = SumaCadena(xNum)
        NumerCalc.Cadena = NumerCalc.Cadena & "/" & xNum
    Wend

    arrNum = Split(NumerCalc.Cadena, "/")

    NumerCalc.Inicial = arrNum(0)    
    NumerCalc.Final = arrNum(UBound(arrNum))

    For i = 0 To UBound(arrNum)
        Select Case arrNum(i)
        Case 11, 22, 33, 44
            NumerCalc.Maestro = arrNum(i)
        Case 13, 14, 16, 19
            NumerCalc.Karma = arrNum(i)
        End Select
    Next

    On Error GoTo 0
    Exit Function

NumerCalc_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento NumerCalc del Módulo FuncionesConversor"
    Resume Next
End Function


Function SumaCadena(ByVal strCadena As String) As Integer

    Dim sSum As Integer
    Dim car As String * 1

    For n = 1 To Len(strCadena)
        car = Mid(strCadena, n, 1)
        If IsNumeric(car) And car <> " " Then
            sSum = sSum + CInt(car)
        End If
    Next
    SumaCadena = sSum

End Function

Function MiJoin(Lista As Resultados) As String
    If Lista.Inicial = Lista.Maestro And Lista.Inicial = Lista.Final Then
        MiJoin = Lista.Final
    ElseIf Lista.Inicial = Lista.Maestro Then
        MiJoin = Lista.Maestro & "/" & Lista.Final
    ElseIf Lista.Maestro > 0 Then
        MiJoin = Lista.Inicial & "/" & Lista.Maestro & "/" & Lista.Final
    ElseIf Lista.Inicial = Lista.Final Then
        MiJoin = Lista.Final
    Else
        MiJoin = Lista.Inicial & "/" & Lista.Final
    End If
End Function
