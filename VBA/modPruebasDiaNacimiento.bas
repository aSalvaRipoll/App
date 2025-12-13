Option Compare Database
Option Explicit

' =============================================================================
' Módulo: modPruebasDiaNacimiento
' Descripción: Pruebas para el cálculo del Día de Nacimiento
' =============================================================================

Public Sub PruebaDiaNacimiento()
    ' Prueba general del cálculo del Día de Nacimiento
    
    Debug.Print "=========================================="
    Debug.Print "PRUEBA: DÍA DE NACIMIENTO"
    Debug.Print "=========================================="
    Debug.Print ""
    
    Dim calculo As clsCalculoDiaNacimiento
    Set calculo = New clsCalculoDiaNacimiento
    
    ' Ejemplo 1: Día básico (número 1-9)
    Debug.Print "--- Ejemplo 1: Día básico ---"
    calculo.FechaNacimiento = #3/5/1985#
    calculo.Calcular
    Debug.Print calculo.DetalleCalculo
    Debug.Print "Resumen: " & calculo.ObtenerResumen
    Debug.Print ""
    
    ' Ejemplo 2: Número maestro 11
    Debug.Print "--- Ejemplo 2: Número Maestro 11 ---"
    Set calculo = New clsCalculoDiaNacimiento
    calculo.FechaNacimiento = #11/15/1980#
    calculo.Calcular
    Debug.Print calculo.DetalleCalculo
    Debug.Print "Resumen: " & calculo.ObtenerResumen
    Debug.Print "¿Es Maestro? " & calculo.EsNumeroMaestro
    Debug.Print ""
    
    ' Ejemplo 3: Número maestro 22
    Debug.Print "--- Ejemplo 3: Número Maestro 22 ---"
    Set calculo = New clsCalculoDiaNacimiento
    calculo.FechaNacimiento = #10/22/1975#
    calculo.Calcular
    Debug.Print calculo.DetalleCalculo
    Debug.Print "Resumen: " & calculo.ObtenerResumen
    Debug.Print "¿Es Maestro? " & calculo.EsNumeroMaestro
    Debug.Print ""
    
    ' Ejemplo 4: Número maestro 29
    Debug.Print "--- Ejemplo 4: Número Maestro 29 ---"
    Set calculo = New clsCalculoDiaNacimiento
    calculo.FechaNacimiento = #1/29/1990#
    calculo.Calcular
    Debug.Print calculo.DetalleCalculo
    Debug.Print "Resumen: " & calculo.ObtenerResumen
    Debug.Print "¿Es Maestro? " & calculo.EsNumeroMaestro
    Debug.Print ""
    
    ' Ejemplo 5: Número kármico 13
    Debug.Print "--- Ejemplo 5: Número Kármico 13 ---"
    Set calculo = New clsCalculoDiaNacimiento
    calculo.FechaNacimiento = #5/13/1988#
    calculo.Calcular
    Debug.Print calculo.DetalleCalculo
    Debug.Print "Resumen: " & calculo.ObtenerResumen
    Debug.Print "¿Es Kármico? " & calculo.EsNumeroKarmico
    Debug.Print "Número Kármico: " & calculo.NumeroKarmico
    Debug.Print ""
    
    ' Ejemplo 6: Número kármico 14
    Debug.Print "--- Ejemplo 6: Número Kármico 14 ---"
    Set calculo = New clsCalculoDiaNacimiento
    calculo.FechaNacimiento = #7/14/1992#
    calculo.Calcular
    Debug.Print calculo.DetalleCalculo
    Debug.Print "Resumen: " & calculo.ObtenerResumen
    Debug.Print "¿Es Kármico? " & calculo.EsNumeroKarmico
    Debug.Print "Número Kármico: " & calculo.NumeroKarmico
    Debug.Print ""
    
    ' Ejemplo 7: Número kármico 16
    Debug.Print "--- Ejemplo 7: Número Kármico 16 ---"
    Set calculo = New clsCalculoDiaNacimiento
    calculo.FechaNacimiento = #9/16/1995#
    calculo.Calcular
    Debug.Print calculo.DetalleCalculo
    Debug.Print "Resumen: " & calculo.ObtenerResumen
    Debug.Print "¿Es Kármico? " & calculo.EsNumeroKarmico
    Debug.Print "Número Kármico: " & calculo.NumeroKarmico
    Debug.Print ""
    
    ' Ejemplo 8: Número kármico 19
    Debug.Print "--- Ejemplo 8: Número Kármico 19 ---"
    Set calculo = New clsCalculoDiaNacimiento
    calculo.FechaNacimiento = #12/19/1987#
    calculo.Calcular
    Debug.Print calculo.DetalleCalculo
    Debug.Print "Resumen: " & calculo.ObtenerResumen
    Debug.Print "¿Es Kármico? " & calculo.EsNumeroKarmico
    Debug.Print "Número Kármico: " & calculo.NumeroKarmico
    Debug.Print ""
    
    ' Ejemplo 9: Número compuesto regular (28)
    Debug.Print "--- Ejemplo 9: Número Compuesto 28 ---"
    Set calculo = New clsCalculoDiaNacimiento
    calculo.FechaNacimiento = #2/28/1993#
    calculo.Calcular
    Debug.Print calculo.DetalleCalculo
    Debug.Print "Resumen: " & calculo.ObtenerResumen
    Debug.Print ""
    
    Set calculo = Nothing
    
    Debug.Print "=========================================="
    Debug.Print "PRUEBA COMPLETADA"
    Debug.Print "=========================================="
End Sub

Public Sub PruebaTodosDiasMes()
    ' Prueba todos los días del mes (1-31)
    ' Útil para verificar que todos los archivos de interpretación existen
    
    Debug.Print "=========================================="
    Debug.Print "PRUEBA: TODOS LOS DÍAS DEL MES (1-31)"
    Debug.Print "=========================================="
    Debug.Print ""
    
    Dim calculo As clsCalculoDiaNacimiento
    Dim dia As Integer
    
    For dia = 1 To 31
        Set calculo = New clsCalculoDiaNacimiento
        
        ' Usar cualquier mes/año válido
        calculo.FechaNacimiento = DateSerial(1990, 1, dia)
        calculo.Calcular
        
        Debug.Print "Día " & Format(dia, "00") & ": " & calculo.ObtenerResumen
        
        Set calculo = Nothing
    Next dia
    
    Debug.Print ""
    Debug.Print "=========================================="
    Debug.Print "PRUEBA COMPLETADA - 31 DÍAS VERIFICADOS"
    Debug.Print "=========================================="
End Sub

Public Sub PruebaIntegracionCompleta()
    ' Prueba integración completa con una persona real
    
    Debug.Print "=========================================="
    Debug.Print "PRUEBA: INTEGRACIÓN COMPLETA"
    Debug.Print "=========================================="
    Debug.Print ""
    
    Dim nombre As String
    Dim fecha As Date
    
    nombre = "María Carmen García López"
    fecha = #3/15/1985#
    
    Debug.Print "DATOS DE LA PERSONA:"
    Debug.Print "-------------------"
    Debug.Print "Nombre: " & nombre
    Debug.Print "Fecha de nacimiento: " & Format(fecha, "dd/mm/yyyy")
    Debug.Print ""
    
    ' Día de Nacimiento
    Debug.Print "DÍA DE NACIMIENTO:"
    Debug.Print "-------------------"
    Dim diaNac As clsCalculoDiaNacimiento
    Set diaNac = New clsCalculoDiaNacimiento
    diaNac.FechaNacimiento = fecha
    diaNac.Calcular
    
    Debug.Print diaNac.DetalleCalculo
    Debug.Print ""
    Debug.Print "RESUMEN: " & diaNac.ObtenerResumen
    
    If diaNac.EsNumeroMaestro Then
        Debug.Print ">>> NÚMERO MAESTRO <<<"
    ElseIf diaNac.EsNumeroKarmico Then
        Debug.Print ">>> NÚMERO KÁRMICO: " & diaNac.NumeroKarmico & " <<<"
    End If
    
    Set diaNac = Nothing
    
    Debug.Print ""
    Debug.Print "=========================================="
    Debug.Print "INTEGRACIÓN COMPLETADA"
    Debug.Print "=========================================="
End Sub

Public Sub VerificarArchivosInterpretacion()
    ' Verifica que existan los 31 archivos de interpretación
    ' NOTA: Esta función requiere acceso al sistema de archivos
    
    Debug.Print "=========================================="
    Debug.Print "VERIFICACIÓN DE ARCHIVOS DE INTERPRETACIÓN"
    Debug.Print "=========================================="
    Debug.Print ""
    
    Dim fso As Object
    Dim rutaBase As String
    Dim dia As Integer
    Dim nombreArchivo As String
    Dim existe As Boolean
    Dim contadorExistentes As Integer
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Ruta base donde deberían estar los archivos
    ' AJUSTAR según tu configuración
    rutaBase = CurrentProject.Path & "\Interpretaciones\DiaNacimiento\"
    
    Debug.Print "Ruta base: " & rutaBase
    Debug.Print ""
    
    contadorExistentes = 0
    
    For dia = 1 To 31
        nombreArchivo = Format(dia, "00") & "_DiaNacimiento.md"
        existe = fso.FileExists(rutaBase & nombreArchivo)
        
        If existe Then
            Debug.Print "✓ " & nombreArchivo & " - EXISTE"
            contadorExistentes = contadorExistentes + 1
        Else
            Debug.Print "✗ " & nombreArchivo & " - NO ENCONTRADO"
        End If
    Next dia
    
    Debug.Print ""
    Debug.Print "=========================================="
    Debug.Print "RESUMEN: " & contadorExistentes & " de 31 archivos encontrados"
    
    If contadorExistentes = 31 Then
        Debug.Print "¡TODOS LOS ARCHIVOS ESTÁN PRESENTES!"
    Else
        Debug.Print "FALTAN " & (31 - contadorExistentes) & " ARCHIVOS"
    End If
    
    Debug.Print "=========================================="
    
    Set fso = Nothing
End Sub
