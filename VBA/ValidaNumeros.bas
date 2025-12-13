Select Case tipo
    Case tiDiaNacimiento
        ValidarNumero = (numero >= 1 And numero <= 31)  ' ⭐ 1-31 para días
    Case Else
        ValidarNumero = (numero >= 1 And numero <= 9) Or _
                        numero = 11 Or numero = 22 Or numero = 33 Or numero = 44
End Select