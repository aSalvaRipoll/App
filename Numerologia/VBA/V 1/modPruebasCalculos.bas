Attribute VB_Name = "modPruebasCalculos"
Option Compare Database
Option Explicit

' ============================================================================
' Módulo: modPruebasCalculos
' Descripción: Procedimientos de prueba para verificar los cálculos
' Autor: Sistema de Numerología
' Fecha: 2024
' ============================================================================

Public Sub PruebaCompletaNumerologia()
    Dim nombrePrueba As String
    Dim fechaPrueba As Date
    
    nombrePrueba = "MARIA CARMEN RODRIGUEZ GARCIA"
    fechaPrueba = #1/15/1985#
    
    Debug.Print "=== PRUEBA COMPLETA DE NUMEROLOGÍA ==="
    Debug.Print "Nombre: " & nombrePrueba
    Debug.Print "Fecha: " & Format(fechaPrueba, "dd/mm/yyyy")
    Debug.Print ""
    
    ' Probar todos los cálculos básicos
    Call PruebaCaminoVida(fechaPrueba)
    Call PruebaDestino(nombrePrueba)
    Call PruebaAlma(nombrePrueba)
    Call PruebaPersonalidad(nombrePrueba)
    Call PruebaMadurez(nombrePrueba, fechaPrueba)
    
    ' Probar cálculos temporales
    Call PruebaAnoPersonal(fechaPrueba, 2024)
    Call PruebaEdadPersonal(fechaPrueba)
    
    ' Probar ciclos
    Call PruebaCiclos(fechaPrueba)
    
    ' Probar pináculos
    Call PruebaPinaculos(fechaPrueba)
    
    ' Probar desafíos
    Call PruebaDesafios(fechaPrueba)
    
    ' Probar números especiales
    Call PruebaNumeroEspeciales(nombrePrueba)
    
    Debug.Print ""
    Debug.Print "=== FIN DE PRUEBAS ==="
End Sub

Private Sub PruebaCaminoVida(ByVal fecha As Date)
    Dim obj As clsCalculoCaminoVida
    Set obj = New clsCalculoCaminoVida
    
    obj.FechaNacimiento = fecha
    Debug.Print "Camino de Vida: " & obj.Calcular()
    
    Set obj = Nothing
End Sub

Private Sub PruebaDestino(ByVal nombre As String)
    Dim obj As clsCalculoDestino
    Set obj = New clsCalculoDestino
    
    obj.NombreCompleto = nombre
    Debug.Print "Destino: " & obj.Calcular()
    
    Set obj = Nothing
End Sub

Private Sub PruebaAlma(ByVal nombre As String)
    Dim obj As clsCalculoAlma
    Set obj = New clsCalculoAlma
    
    obj.NombreCompleto = nombre
    Debug.Print "Alma: " & obj.Calcular()
    
    Set obj = Nothing
End Sub

Private Sub PruebaPersonalidad(ByVal nombre As String)
    Dim obj As clsCalculoPersonalidad
    Set obj = New clsCalculoPersonalidad
    
    obj.NombreCompleto = nombre
    Debug.Print "Personalidad: " & obj.Calcular()
    
    Set obj = Nothing
End Sub

Private Sub PruebaMadurez(ByVal nombre As String, ByVal fecha As Date)
    Dim obj As clsCalculoMadurez
    Set obj = New clsCalculoMadurez
    
    obj.NombreCompleto = nombre
    obj.FechaNacimiento = fecha
    Debug.Print "Madurez: " & obj.Calcular()
    
    Set obj = Nothing
End Sub

Private Sub PruebaAnoPersonal(ByVal fecha As Date, ByVal ano As Integer)
    Dim obj As clsCalculoAnoPersonal
    Set obj = New clsCalculoAnoPersonal
    
    obj.FechaNacimiento = fecha
    obj.AnoCalculo = ano
    Debug.Print "Año Personal " & ano & ": " & obj.Calcular()
    
    Set obj = Nothing
End Sub

Private Sub PruebaEdadPersonal(ByVal fecha As Date)
    Dim obj As clsCalculoEdadPersonal
    Set obj = New clsCalculoEdadPersonal
    
    obj.FechaNacimiento = fecha
    Debug.Print "Edad Personal: " & obj.Calcular() & " (Edad: " & obj.EdadActual & ")"
    
    Set obj = Nothing
End Sub

Private Sub PruebaCiclos(ByVal fecha As Date)
    Dim obj1 As clsCalculoCiclo1
    Dim obj2 As clsCalculoCiclo2
    Dim obj3 As clsCalculoCiclo3
    
    Debug.Print ""
    Debug.Print "--- CICLOS ---"
    
    Set obj1 = New clsCalculoCiclo1
    obj1.FechaNacimiento = fecha
    Debug.Print "Ciclo 1: " & obj1.Calcular() & " (" & obj1.ObtenerRangoEdades() & ")"
    
    Set obj2 = New clsCalculoCiclo2
    obj2.FechaNacimiento = fecha
    Debug.Print "Ciclo 2: " & obj2.Calcular() & " (" & obj2.ObtenerRangoEdades() & ")"
    
    Set obj3 = New clsCalculoCiclo3
    obj3.FechaNacimiento = fecha
    Debug.Print "Ciclo 3: " & obj3.Calcular() & " (" & obj3.ObtenerRangoEdades() & ")"
    
    Set obj1 = Nothing
    Set obj2 = Nothing
    Set obj3 = Nothing
End Sub

Private Sub PruebaPinaculos(ByVal fecha As Date)
    Dim obj1 As clsCalculoPinaculo1
    Dim obj2 As clsCalculoPinaculo2
    Dim obj3 As clsCalculoPinaculo3
    Dim obj4 As clsCalculoPinaculo4
    
    Debug.Print ""
    Debug.Print "--- PINÁCULOS ---"
    
    Set obj1 = New clsCalculoPinaculo1
    obj1.FechaNacimiento = fecha
    Debug.Print "Pináculo 1: " & obj1.Calcular() & " (" & obj1.ObtenerRangoEdades() & ")"
    
    Set obj2 = New clsCalculoPinaculo2
    obj2.FechaNacimiento = fecha
    Debug.Print "Pináculo 2: " & obj2.Calcular() & " (" & obj2.ObtenerRangoEdades() & ")"
    
    Set obj3 = New clsCalculoPinaculo3
    obj3.FechaNacimiento = fecha
    Debug.Print "Pináculo 3: " & obj3.Calcular() & " (" & obj3.ObtenerRangoEdades() & ")"
    
    Set obj4 = New clsCalculoPinaculo4
    obj4.FechaNacimiento = fecha
    Debug.Print "Pináculo 4: " & obj4.Calcular() & " (" & obj4.ObtenerRangoEdades() & ")"
    
    Set obj1 = Nothing
    Set obj2 = Nothing
    Set obj3 = Nothing
    Set obj4 = Nothing
End Sub

Private Sub PruebaDesafios(ByVal fecha As Date)
    Dim obj1 As clsCalculoDesafio1
    Dim obj2 As clsCalculoDesafio2
    Dim obj3 As clsCalculoDesafio3
    Dim obj4 As clsCalculoDesafio4
    
    Debug.Print ""
    Debug.Print "--- DESAFÍOS ---"
    
    Set obj1 = New clsCalculoDesafio1
    obj1.FechaNacimiento = fecha
    Debug.Print "Desafío 1: " & obj1.Calcular() & " (" & obj1.ObtenerRangoEdades() & ")"
    
    Set obj2 = New clsCalculoDesafio2
    obj2.FechaNacimiento = fecha
    Debug.Print "Desafío 2: " & obj2.Calcular() & " (" & obj2.ObtenerRangoEdades() & ")"
    
    Set obj3 = New clsCalculoDesafio3
    obj3.FechaNacimiento = fecha
    Debug.Print "Desafío 3: " & obj3.Calcular() & " (" & obj3.ObtenerRangoEdades() & ")"
    
    Set obj4 = New clsCalculoDesafio4
    obj4.FechaNacimiento = fecha
    Debug.Print "Desafío 4: " & obj4.Calcular() & " (" & obj4.ObtenerRangoEdades() & ")"
    
    Set obj1 = Nothing
    Set obj2 = Nothing
    Set obj3 = Nothing
    Set obj4 = Nothing
End Sub

Private Sub PruebaNumeroEspeciales(ByVal nombre As String)
    Dim objPoder As clsCalculoNumeroPoder
    Dim objPrimeraLetra As clsCalculoPrimeraLetra
    Dim objDominante As clsCalculoNumeroDominante
    Dim objFaltante As clsCalculoNumeroFaltante
    Dim objPlano As clsCalculoPlanoExpresion
    
    Debug.Print ""
    Debug.Print "--- NÚMEROS ESPECIALES ---"
    
    Set objPoder = New clsCalculoNumeroPoder
    objPoder.NombreCompleto = nombre
    Debug.Print "Número de Poder (Primera Vocal): " & objPoder.Calcular() & " (" & objPoder.PrimeraVocal & ")"
    
    Set objPrimeraLetra = New clsCalculoPrimeraLetra
    objPrimeraLetra.NombreCompleto = nombre
    Debug.Print "Primera Letra: " & objPrimeraLetra.Calcular() & " (" & objPrimeraLetra.PrimeraLetra & ")"
    
    Set objDominante = New clsCalculoNumeroDominante
    objDominante.NombreCompleto = nombre
    Debug.Print "Número Dominante: " & objDominante.Calcular() & " (frecuencia: " & objDominante.MaximaFrecuencia & ")"
    
    Set objFaltante = New clsCalculoNumeroFaltante
    objFaltante.NombreCompleto = nombre
    objFaltante.Calcular
    Debug.Print "Números Faltantes: " & objFaltante.ObtenerNumerosFaltantes()
    
    Set objPlano = New clsCalculoPlanoExpresion
    objPlano.NombreCompleto = nombre
    objPlano.Calcular
    Debug.Print "Planos de Expresión: " & objPlano.ObtenerResumen()
    
    Set objPoder = Nothing
    Set objPrimeraLetra = Nothing
    Set objDominante = Nothing
    Set objFaltante = Nothing
    Set objPlano = Nothing
End Sub