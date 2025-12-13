Dim calc As clsCalculoDiaNacimiento
Set calc = New clsCalculoDiaNacimiento

calc.FechaNacimiento = #3/15/1985#
calc.Calcular

Debug.Print calc.ObtenerResumen
' Salida: "Día de Nacimiento: 15 - El Armonizador Versátil (15/6)"

If calc.EsNumeroMaestro Then
    Debug.Print "¡Es un número maestro!"
End If

If calc.EsNumeroKarmico Then
    Debug.Print "Número kármico: " & calc.NumeroKarmico
End If

Set calc = Nothing