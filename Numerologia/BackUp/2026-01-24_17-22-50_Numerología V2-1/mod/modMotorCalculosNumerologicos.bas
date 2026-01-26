Attribute VB_Name = "modMotorCalculosNumerologicos"

Option Compare Database
Option Explicit

' ------------------------------------------------------
' Nombre:    modCalculosNumerologicos
' Tipo:      Módulo
' Propósito: Realiza los cálculos Numerológicos
' Autor:     Alba Salvá
' Fecha:     15/01/2026
' Versión:   2.0
' ------------------------------------------------------

Private mAcumNombre As tAcumuladores
Private mAcumApe1 As tAcumuladores
Private mAcumApe2 As tAcumuladores

Private colFon As Collection

Dim t As Variant
Dim colTrans As Collection

Private Const mVersionMotor As String = "2.0"

'Dim Incl As clsInclusion
'Dim pd As clsPinaDes
'Dim c As clsCiclos
'Dim tr As clsTransitos

'Public Sub CalcularResultado(ByRef r As clsResultado, ByRef P As clsPersona, ByRef f As clsFonetica)
'Public Sub CalcularResultado(ByRef r As clsResultado, ByRef D As clsDTOcalculos)

Public Sub CalcularResultado(ByRef D As clsDTOcalculos, _
                             ByRef r As clsResultado, _
                             ByRef Incl As clsInclusion, _
                             ByRef c As clsCiclos, _
                             ByRef pd As clsPinaDes)

    '
    r.VersionMotor = mVersionMotor
    
    
    
    Call InicializarDTOs(D, r, Incl, c, pd)
    Call EjecutarCalculos(D, r, Incl, c, pd)
    Call PersistirResultados(r, Incl, c, pd, colTrans)

End Sub

Private Sub InicializarDTOs(ByRef D As clsDTOcalculos, _
                           ByRef r As clsResultado, _
                           ByRef Incl As clsInclusion, _
                           ByRef c As clsCiclos, _
                           ByRef pd As clsPinaDes)

    ' ============================
    '  RESULTADO PRINCIPAL
    ' ============================
    With r
        .IDPersona = D.IDPersona
        .IDFonetica = D.IDFonetica
        '.IDResultado = AutoNext("IDResultado", "tbuResultados", _
                                "IDPersona = " & D.IDPersona & _
                                " AND IDFonetica = " & D.IDFonetica)
        .Activo = True
        .SistemaCalculo = D.SistemaCalculo
        .SistemaTarot = D.SistemaTarot
        .NumCiclos = D.NumCiclos
        .SistemaCiclos = D.SistemaCiclos
    End With


    ' ============================
    '  INCLUSIÓN
    ' ============================
    With Incl
        .IDPersona = D.IDPersona
        .IDFonetica = D.IDFonetica
        '.IDResultado = r.IDResultado
        ' Los contadores N1..N9 ya están a 0 por Class_Initialize
    End With


    ' ============================
    '  CICLOS
    ' ============================
    With c
        .IDPersona = D.IDPersona
        '.IDResultado = r.IDResultado
        .NumCiclos = D.NumCiclos
        .MetodoCiclos = D.SistemaCiclos
        ' Los valores de Ciclo1..4 y edades ya están a 0
    End With


    ' ============================
    '  PINÁCULOS Y DESAFÍOS
    ' ============================
    With pd
        .IDPersona = D.IDPersona
        '.IDResultado = r.IDResultado
        ' Pina1..4, Desa1..4 y fechas ya están vacíos por Class_Initialize
    End With

End Sub

Private Sub EjecutarCalculos(ByRef D As clsDTOcalculos, _
                             ByRef r As clsResultado, _
                             ByRef Incl As clsInclusion, _
                             ByRef c As clsCiclos, _
                             ByRef pd As clsPinaDes)

    ' ============================
    '  1. ACUMULADORES FONÉTICOS
    ' ============================
    Call CargarAcumuladoresNombre(Incl, D)


    ' ============================
    '  2. NÚMEROS BÁSICOS
    ' ============================
    Call CalcularBasicos(r, D)


    ' ============================
    '  3. NÚMEROS DERIVADOS
    ' ============================
    Call CalcularDerivados(r, Incl)


    ' ============================
    '  4. LETRAS Y SÍMBOLOS
    ' ============================
    Call CalcularLetras(r, D)


    ' ============================
    '  5. FRECUENCIAS (INCLUSIÓN)
    ' ============================
    Call CalcularFrecuencias(r, Incl)


    ' ============================
    '  6. TEMPORALES
    ' ============================
    Call CalcularTemporal(r, D)


    ' ============================
    '  7. CICLOS
    ' ============================
    Call CalcularCiclos(r, D, c)


    ' ============================
    '  8. PINÁCULOS Y DESAFÍOS
    ' ============================
    Call CalcularPinaDes(r, D, pd)


    ' ============================
    '  9. TRÁNSITOS (cuando lo integres)
    ' ============================
    'Call CalcularTransitos(r, D, ColeccionTransitos)

End Sub


'Public Sub CalcularResultado(ByRef D As clsDTOcalculos, _
'                             ByRef r As clsResultado, _
'                             ByRef Incl As clsInclusion, _
'                             ByRef c As clsCiclos, _
'                             ByRef pd As clsPinaDes) ', _
'                             ByRef tr As clsTransitos) ', _
'                             'ByRef pr As clsProgresiones)
'
''    Set Progres = New clsProgresiones
'
'
'    'Inicializamos clases de objetos
'    With r
'        .IDPersona = D.IDPersona
'        .IDFonetica = D.IDFonetica
'        .idResultado = AutoNext("IDResultado", "tbuResultados", "idPersona = " & D.IDPersona & " AND IDFonetica = " & D.IDFonetica)
'    End With
'
''    Set Incl = New clsInclusion
'    With Incl
'        .IDPersona = D.IDPersona
'        .IDFonetica = D.IDFonetica
'        .idResultado = r.idResultado
'        '.IDInclusion = AutoNext("IDInclusion", "tbuInclusion")
'    End With
'
''    Set pd = New clsPinaDes
'    With pd
'        .IDPersona = D.IDPersona
'        .idResultado = r.idResultado
'        '.idPinaDes = AutoNext("IDPinaDes", "tbuPinaDes")
'    End With
'
''    Set c = New clsCiclos
'    With c
'        .IDPersona = D.IDPersona
'        .NumCiclos = D.NumCiclos
'        .MetodoCiclos = D.MetodoCiclos
'        .idResultado = r.idResultado
'        '.idCiclo = AutoNext("IDCiclo", "tbuCiclos")
'
'    End With
'
''    With Progres
''        .IDPersona = D.ID_Persona
''        '.idResultado = r.idResultado
''        '.IDProg = AutoNext("IDProg", "tbuProgresiones")
''    End With
'
''    Set tr = New clsTransitos
''    With tr
''        .IDPersona = D.IDPersona
''        .idResultado = r.idResultado
''        '.IDTransito = AutoNext("IDResultado", "tbuResultados")
''    End With
'
'
'    ' CÁLCULOS BASE DE LETRAS
'    'Call CargarAcumuladoresNombre(P, f)
'    Call CargarAcumuladoresNombre(D)
'
'
'    ' ORDEN RECOMENDADO DE CÁLCULO
'    Call CalcularBasicos(r, D)
'    Call CalcularDerivados(r)
'
'    Call CalcularLetras(r, D)
'
'    Call CalcularFrecuencias(r)
'
'    Call CalcularTemporal(r, D)
'
''   +- Ciclos
'    Call CalcularCiclos(r, D, c)
''   +- Pináculos
'    Call CalcularPinaDes(r, D, pd)
''   +- Desafíos
'
'
'    ' GUARDAR RESULTADOS (Persistencia)
'    Call GuardaDatosResultados(r, Incl, pd, c) ', Progres)
'
''    Call CalcularProgTransit(r, P, Progres, Transit, f)
'
'End Sub

'Sub CalcularBasicos(r As clsResultado, P As clsPersona) ', F As clsFonetica)
Sub CalcularBasicos(r As clsResultado, D As clsDTOcalculos) ', F As clsFonetica)

'1. FASE BÁSICA (Del nombre y fecha base):
'   +- Camino de Vida
    r.NumeroCaminoVida = CalcularCaminoVida(D)
    'Debug.Print "CaminoVida: "; r.NumeroCaminoVida
'   +- Destino

    r.NumeroDestino = CalcularDestino 'ReducirSimbolico(mAcumNombre.Completo + mAcumApe1.Completo + mAcumApe2.Completo) 'mSumaLetras)
    'Debug.Print "Destino: "; r.NumeroDestino
'   +- Alma
    r.NumeroAlma = CalcularAlma ' ReducirSimbolico(mAcumNombre.Vocales + mAcumApe1.Vocales + mAcumApe2.Vocales) 'mSumaVocales)
    'Debug.Print "Alma: "; r.NumeroAlma
'   +- Personalidad
    r.NumeroPersonalidad = CalcularPersonalidad 'ReducirSimbolico(mAcumNombre.Consonantes + mAcumApe1.Consonantes + mAcumApe2.Consonantes) 'mSumaConsonantes)
    'Debug.Print "Personalidad: "; r.NumeroPersonalidad
'   +- Día de Nacimiento
    r.NumeroDiaNacimiento = Day(D.FechaNacimiento)

End Sub

Sub CalcularDerivados(r As clsResultado, Incl As clsInclusion)

'2. FASE DERIVADA (Requieren otros números):
'   +- Madurez (necesita CV + Destino)
    r.NumeroMadurez = CalcularMadurez(r)
    'Debug.Print "Madurez: "; r.NumeroMadurez
'   +- Número de Poder
    r.NumeroPoder = CalculaPoder(Incl)
   ' Debug.Print "Poder: "; r.NumeroPoder
'   +- Respuesta Subconsciente
    r.RespuestaSubconsciente = CalculaRespuestaSubconsciente(Incl)
    'Debug.Print "Subconsciente: "; r.NumeroRespuestaSubconsciente
'   +- Plano de Expresión
'    PlanoFisico
'    PlanoEmocional
'    PlanoMental
'    PlanoIntuitivo
    Call CalculaPlanos(r, Incl)
'   +- Árbol de Vida
    Call CalcularArbolVida(r)
'   +- Deuda Kármica (compilar hallazgos)No se hace


End Sub
    
'Sub CalcularLetras(r As clsResultado, P As clsPersona, f As clsFonetica)
Sub CalcularLetras(r As clsResultado, D As clsDTOcalculos)

'3. FASE DE ANÁLISIS DE LETRAS:
'   +- Piedra Angular
    r.PiedraAngular = CalcularPiedraAngular(D)
'   +- Piedra de Toque
    r.PiedraToque = CalcularPiedraToque(D)
'   +- Primera Vocal
    r.PrimeraVocal = ObtenerPrimeraVocal(D)
'   +- Primera Consonante
    r.PrimeraConsonante = ObtenerPrimeraConsonante(D)
'   +- Primera Letra
    r.PrimeraLetra = ObtenerPrimeraLetra(D)

End Sub
    
Sub CalcularFrecuencias(r As clsResultado, Incl As clsInclusion)
'4. FASE DE FRECUENCIAS:
'   +- Dominantes y Ausentes
    r.Dominantes = CalculaDominantes(Incl)
    r.Ausentes = CalculaAusentes(Incl)

End Sub
    
'Sub CalcularTemporal(r As clsResultado, P As clsPersona, c As clsCiclos, pd As clsPinaDes)
Sub CalcularTemporal(r As clsResultado, D As clsDTOcalculos)
'5. FASE TEMPORAL:
'   +- Edad de la persona
    r.Edad = GetEdad(D.FechaNacimiento)
'   +- Dia Nacimiento
    r.NumeroDiaNacimiento = CalculaDiaNacimiento(D.FechaNacimiento)
'   +- Año Personal
    r.AnioPersonal = CalculaAnioPersonal(D.FechaNacimiento)
'   +- Edad Personal
    'r.NumeroEdadPersonal = CalculaEdadPersonal(D.FechaNacimiento)
    r.EdadPersonal = ReducirSimbolico(r.Edad)

End Sub
    
'Sub CalcularProgTransit(r As clsResultado, P As clsPersona, PR As clsProgresiones, t As clsTransitos, f As clsFonetica)
'Sub CalcularProgTransit(r As clsResultado, pr As clsProgresiones, tr As clsTransitos, D As clsDTOcalculos)
Sub CalcularProgTransit(r As clsResultado, tr As clsTransito, D As clsDTOcalculos)
'6. FASE Progresiones:
    
    Dim t As Variant

'   +- Progresiones
    'Call CalcularProgresiones(r, D, pr)
'   +- Esencia
    Call CalcularEsencia(r)
    
'   +- Transitos (Array Decenal)
    'T = GenerarTablaDecadas(R, P, f)
    'Call ExportarTablaDecadasHTML(T, "N:\decadas.html")

    ' 1. Calcular matriz de tránsitos
    tr = GenerarTablaDecadas(D, r)

    ' 2. Crear objetos clsTransitos
    Set colTrans = CrearTransitosDesdeTabla(tr, D, r)


    ' 3. Guardar en lote
'    Call GuardarTransitosEnLote(colTrans)

    'Call GenerarTablaDecadas(P, f, r)
'   +- Transito Físico
'   +- Transito Mental
'   +- Transito Emocional
'   +- Transito Espiritual
    'Call CalcularTransitos(T, P, f, R)
    'Call CalcularTransitosTradicional(T, P, F, R)
    'Call CalcularTransitosFonetico(T, P, f, R)
    



End Sub

'Sub GuardaDatosResultados(r As clsResultado, Inclusion As clsInclusion, PinaDes As clsPinaDes, Ciclos As clsCiclos) ', Progres As clsProgresiones)
'
'    'r.idResultado = AutoNext("IDResultado", "tbuResultados")
'    Call GuardarResultado(r)
'
'    'Inclusion.idResultado = r.idResultado
'    Call GuardarInclusion(Inclusion)
'
'    'PinaDes.idResultado = r.idResultado
'    Call GuardarPinaDes(PinaDes)
'
'    'Ciclos.idResultado = r.idResultado
'    Call GuardarCiclos(Ciclos)
'
'End Sub
    
    
'===================================================================================
' CALCULOS
'===================================================================================
    
Public Function CalcularCaminoVida(D As clsDTOcalculos) As String
    Dim anioRed As Integer
    Dim suma As Integer

    ' Reducir el año primero
    anioRed = SumarDigitos(Year(D.FechaNacimiento))

    ' Sumar día + mes + año reducido
    suma = Day(D.FechaNacimiento) + Month(D.FechaNacimiento) + anioRed

    ' Aplicar reducción simbólica elegante
    CalcularCaminoVida = ReducirSimbolico(suma)
End Function


Public Function CalcularDestino() As String
    ' Se asume que CargarAcumuladoresNombre ya se ha llamado antes
    CalcularDestino = ReducirSimbolico(mAcumNombre.Completo + mAcumApe1.Completo + mAcumApe2.Completo)
End Function

Public Function CalcularAlma() As String
    ' Se asume que CargarAcumuladoresNombre ya se ha llamado antes
    CalcularAlma = ReducirSimbolico(mAcumNombre.Vocales + mAcumApe1.Vocales + mAcumApe2.Vocales)
End Function

Public Function CalcularPersonalidad() As String
    ' Se asume que CargarAcumuladoresNombre ya se ha llamado antes
    CalcularPersonalidad = ReducirSimbolico(mAcumNombre.Consonantes + mAcumApe1.Consonantes + mAcumApe2.Consonantes)
End Function


Public Function CalcularMadurez(r As clsResultado) As String
    Dim cv As String, dest As String
    Dim rcv As Long, rdest As Long

    cv = r.NumeroCaminoVida
    dest = r.NumeroDestino

    rcv = CLng(Split(cv, "/")(UBound(Split(cv, "/"))))
    rdest = CLng(Split(dest, "/")(UBound(Split(dest, "/"))))

    CalcularMadurez = ReducirSimbolico(rcv + rdest)
End Function

'=====================================================================================
' ÁRBOL DE VIDA (Árbol Genealógico Numerológico)
'
'**QUÉ REPRESENTA:** Herencia familiar, patrones ancestrales.
'
'**ENTRADA:** Apellidos completos (sin nombre de pila)
'
Public Sub CalcularArbolVida(r As clsResultado)

    r.ArbolPaterno = ReducirSimbolico(mAcumApe1.Completo)
    r.ArbolMaterno = ReducirSimbolico(mAcumApe2.Completo)
    r.ArbolDeVida = ReducirSimbolico(mAcumApe1.Completo + mAcumApe2.Completo)

End Sub

Public Function CalculaPoder(i As clsInclusion) As String
'**QUÉ REPRESENTA:** Tu pasión oculta, el talento que más utilizas, la energía más fuerte en tu nombre.
'
'**ENTRADA:** Nombre completo
'
'**PROCESO:**
'
'1. Convertir todas las letras a valores (1-9)
'   - Reducir cualquier maestro (11-->2, 22-->4)
'   - Dígrafos: CH, LL, RR
'
    CalculaPoder = CollectionToString(i.GetMaximos)

End Function

Function CalculaRespuestaSubconsciente(i As clsInclusion) As Byte
'**QUÉ REPRESENTA:** Capacidad de respuesta en crisis, recursos automáticos ante emergencias, confianza en uno mismo bajo presión.
'
'**ENTRADA:** Nombre completo
'
'**PROCESO:**
'1. Convertir cada letra a valor (1-9)
'   - Reducir maestros (11?2, 22?4)
'   - Dígrafos reducidos
'2. Contar cuántos números diferentes del 1 al 9 están presentes
'3. **Fórmula:**
'   Respuesta Subconsciente = Cantidad de números presentes
'   O alternativamente:
'   Respuesta Subconsciente = 9 - (cantidad de números ausentes)

    CalculaRespuestaSubconsciente = i.GetCuenta

End Function

Sub CalculaPlanos(r As clsResultado, i As clsInclusion)
'**QUÉ REPRESENTA:** Cómo te expresas en cuatro niveles: Físico, Mental, Emocional, Intuitivo.
'
'**ENTRADA:** Nombre completo
'
'**PROCESO:**
'**PASO 1: Clasificar cada letra según su plano**
'Las letras se agrupan según su valor numerológico:
    With i
'**PLANO FÍSICO (hacer, actuar):**
'- Letras con valor 4 o 5
'- Letras: **D, E, M, N, W, Ñ**
        r.PlanoFisico = .N4 + .N5
'**PLANO MENTAL (pensar, analizar):**
'- Letras con valor 1 u 8
'- Letras: **A, H, J, Q, S**
        r.PlanoMental = .N1 + .N8
'**PLANO EMOCIONAL (sentir, relacionarse):**
'- Letras con valor 2, 3 o 6
'- Letras: **B, C, F, K, L, O, T, U, X**
        r.PlanoEmocional = .N2 + .N3 + .N6
'**PLANO INTUITIVO (intuir, percibir):**
'- Letras con valor 7 o 9
'- Letras: **G, I, P, R, Z**
        r.PlanoIntuitivo = .N7 + .N9
    
    End With
End Sub


Function CalculaDominantes(i As clsInclusion) As String

    CalculaDominantes = CollectionToString(i.GetDominantes)

End Function

Function CalculaAusentes(i As clsInclusion) As String

    CalculaAusentes = CollectionToString(i.GetAusentes)

End Function

Function CalculaDiaNacimiento(FechaNac As Date) As Byte

    CalculaDiaNacimiento = Day(FechaNac)

End Function

Function CalculaAnioPersonal(FechaNac As Date) As String

    Dim dia As Byte
    Dim mes As Byte
    Dim redAnio As Integer
    
    dia = Day(FechaNac)
    mes = Month(FechaNac)
    
    redAnio = Year(Date)
    
    While redAnio > 9
        redAnio = SumarDigitos(redAnio)
    Wend
    
    CalculaAnioPersonal = ReducirSimbolico(dia + mes + redAnio)
    
    'Debug.Print "Año Personal: "; CalculaAnioPersonal
    
End Function

'-------------------------------------------------------------------
'Grupo Letras
'-------------------------------------------------------------------

Public Function ObtenerPrimeraLetra(D As clsDTOcalculos) As String
    
    Dim col As Collection
    Dim txt As String
    Dim letra As String
    
    letra = ""
    
    txt = Trim$(D.Nombre & " " & D.Ape1 & " " & D.Ape2)
    
    Select Case D.SistemaFonetico
        ' ============================================================
        ' TRADICIONAL (grafía)
        ' ============================================================
        Case mfTradicional
            If Len(txt) > 0 Then
                letra = Left$(txt, 1)
            End If

        ' ============================================================
        ' FONÉTICO (fonemas)
        ' ============================================================
        Case mfFonetico
            Set col = ExtraerFonemasFinales(txt)
            If col.Count > 0 Then
                letra = col(1)  ' primer fonema
            End If

    End Select
    
    ObtenerPrimeraLetra = letra
    
End Function

Public Function ObtenerPrimeraVocal(D As clsDTOcalculos) As String
    Dim col As Collection
    Dim txt As String
    Dim i As Long
    Dim letra As String

    letra = ""
    
    txt = Trim$(D.Nombre & " " & D.Ape1 & " " & D.Ape2)
    
    Select Case D.SistemaFonetico

        Case mfTradicional ' Tradicional
            For i = 1 To Len(txt)
                If EsVocal(Mid$(txt, i, 1)) Then
                    letra = Mid$(txt, i, 1)
                    Exit For
                End If
            Next i

        Case mfFonetico ' Fonético
            Set col = ExtraerFonemasFinales(txt)
            For i = 1 To col.Count
                If EsVocal(col(i)) Then
                    letra = col(i)
                    Exit For
                End If
            Next i

    End Select
    
    ObtenerPrimeraVocal = letra
    
End Function

Public Function ObtenerPrimeraConsonante(D As clsDTOcalculos) As String
    
    Dim i As Long
    Dim txt As String
    Dim col As Collection
    Dim letra As String
    
    letra = ""
    
    txt = Trim$(D.Nombre & " " & D.Ape1 & " " & D.Ape2)
    
    Select Case D.SistemaFonetico
        ' ============================================================
        ' TRADICIONAL (grafía)
        ' ============================================================
        Case mfTradicional
            For i = 1 To Len(txt)
                If Not EsVocal(Mid$(txt, i, 1)) Then
                    letra = Mid$(txt, i, 1)
                    Exit For
                End If
            Next i

        ' ============================================================
        ' FONÉTICO (fonemas)
        ' ============================================================
        Case mfFonetico
            Set col = ExtraerFonemasFinales(txt)

            For i = 1 To col.Count
                If Not EsVocal(col(i)) Then
                    letra = col(i)
                    Exit For
                End If
            Next i

    End Select
    
    ObtenerPrimeraConsonante = letra
    
End Function

Public Function CalcularPiedraAngular(D As clsDTOcalculos) As String
    Dim letra As String
    Dim txt As String
    Dim col As Collection

    txt = Trim$(D.Nombre & " " & D.Ape1 & " " & D.Ape2)
    
    Select Case D.SistemaFonetico
        ' ============================================================
        ' TRADICIONAL (grafía)
        ' ============================================================
        Case mfTradicional
            If Len(txt) >= 2 And UCase$(Left$(txt, 2)) = "LL" Then
                letra = "LL"
            ElseIf Len(txt) >= 2 And UCase$(Left$(txt, 2)) = "CH" Then
                letra = "CH"
            Else
                letra = Left$(txt, 1)
            End If

        ' ============================================================
        ' FONÉTICO (fonemas)
        ' ============================================================
        Case mfFonetico
            Set col = ExtraerFonemasFinales(txt)
            letra = col(1)                     ' primer fonema
    End Select
    
    CalcularPiedraAngular = letra
    
End Function

Public Function CalcularPiedraToque(D As clsDTOcalculos) As String
    Dim letra As String
    Dim txt As String
    Dim col As Collection
    Dim n As Long

    txt = Trim$(D.Nombre & " " & D.Ape1 & " " & D.Ape2)
    
    Select Case D.SistemaFonetico
        ' ============================================================
        ' TRADICIONAL (grafía)
        ' ============================================================
        Case mfTradicional
            n = Len(txt)

            If n >= 2 And UCase$(Right$(txt, 2)) = "LL" Then
                letra = "LL"
            ElseIf n >= 2 And UCase$(Right$(txt, 2)) = "CH" Then
                letra = "CH"
            Else
                letra = Right$(txt, 1)
            End If

        ' ============================================================
        ' FONÉTICO (fonemas)
        ' ============================================================
        Case mfFonetico
            Set col = ExtraerFonemasFinales(txt)
            letra = col(col.Count)             ' último fonema
    End Select
    
    CalcularPiedraToque = letra

End Function

'------------------------------------------------------------------------
' Calcular ciclos

Public Sub CalcularCiclos(r As clsResultado, D As clsDTOcalculos, c As clsCiclos)

    Select Case D.NumCiclos
        Case 1 'Tres Ciclos
            Call CalcularCiclos3(r, D, c)
        Case 2 'Cuatro Ciclos
            Call CalcularCiclos4(r, D, c)
        Case Else
            'Error: tipo no definido
    End Select

End Sub

Private Sub CalcularCiclos3(r As clsResultado, D As clsDTOcalculos, c As clsCiclos)

    Dim cv As Integer
    Dim primera As Byte, segunda As Byte
    Dim anioRed As Integer


    '-----------------------------------------
    ' 1. NÚMEROS DE CICLO (valores crudos)
    '-----------------------------------------

    ' Ciclo 1 = mes reducido
    c.Ciclo1 = CByte(SumarDigitos(Month(D.FechaNacimiento)))

    ' Ciclo 2 = día reducido
    c.Ciclo2 = CByte(SumarDigitos(Day(D.FechaNacimiento)))

    ' Ciclo 3 = año reducido
    anioRed = SumarDigitos(Year(D.FechaNacimiento))
    c.Ciclo3 = CByte(SumarDigitos(anioRed))

    '-----------------------------------------
    ' 2. TRANSICIONES
    '-----------------------------------------

    ' Extraer el valor crudo del Camino de Vida (ej: "11/2" ? 2)
    cv = ValorFinal(r.NumeroCaminoVida)

    primera = 36 - cv
    segunda = primera + 27

    '-----------------------------------------
    ' 3. EDADES
    '-----------------------------------------

    ' Ciclo 1
    c.EdadIni1 = 0
    c.EdadFin1 = CByte(primera)

    ' Ciclo 2
    c.EdadIni2 = CByte(primera)
    c.EdadFin2 = CByte(segunda)

    ' Ciclo 3
    c.EdadIni3 = CByte(segunda)
    c.EdadFin3 = 120   ' vida completa simbólica

    ' Ciclo 4 no se usa
    c.Ciclo4 = 0
    c.EdadIni4 = 0
    c.EdadFin4 = 0

End Sub

Private Sub CalcularCiclos4(r As clsResultado, D As clsDTOcalculos, c As clsCiclos)

    Dim trans As Integer
    Dim primera As Byte
    Dim segunda As Byte
    Dim tercera As Byte
    Dim anioRed As Integer

    'Pasar método de cálculo a la clase
    c.MetodoCiclos = D.SistemaCiclos
    '-----------------------------------------
    ' 1. NÚMEROS DE CICLO (valores crudos)
    '-----------------------------------------

    c.Ciclo1 = CByte(SumarDigitos(Month(D.FechaNacimiento)))
    c.Ciclo2 = CByte(SumarDigitos(Day(D.FechaNacimiento)))

    anioRed = SumarDigitos(Year(D.FechaNacimiento))
    c.Ciclo3 = CByte(SumarDigitos(anioRed))

    ' Invierno = mes + año reducido
    c.Ciclo4 = CByte(SumarDigitos(c.Ciclo1 + c.Ciclo3))

    '-----------------------------------------
    ' 2. TRANSICIONES
    '-----------------------------------------
    trans = ValorFinal(r.NumeroCaminoVida)
    
    If c.MetodoCiclos = ccFijo Then
        primera = 27 - trans
        segunda = 36
        tercera = 45

    ElseIf c.MetodoCiclos = ccClasico Then
        primera = 36 - trans
        segunda = primera + 9
        tercera = segunda + 9

    ElseIf c.MetodoCiclos = ccModerno Then
        trans = ValorFinal(r.NumeroDestino)
        primera = 36 - trans
        segunda = primera + 9
        tercera = segunda + 9
    Else
        'Error: método no definido
    End If

    '-----------------------------------------
    ' 3. EDADES
    '-----------------------------------------

    c.EdadIni1 = 0
    c.EdadFin1 = CByte(primera)

    c.EdadIni2 = CByte(primera)
    c.EdadFin2 = CByte(segunda)

    c.EdadIni3 = CByte(segunda)
    c.EdadFin3 = CByte(tercera)

    c.EdadIni4 = CByte(tercera)
    c.EdadFin4 = 120

End Sub

Public Sub CalcularPinaDes(r As clsResultado, D As clsDTOcalculos, pd As clsPinaDes)

    Dim mes As Byte
    Dim dia As Byte
    Dim anio As Integer
    Dim cv As Byte

    Dim primera As Byte
    Dim segunda As Byte
    Dim tercera As Byte

    '-----------------------------------------
    ' 1. EXTRAER FECHA EN VALORES CRUDOS
    '-----------------------------------------

    mes = SumarDigitos(Month(D.FechaNacimiento))
    dia = SumarDigitos(Day(D.FechaNacimiento))
    anio = SumarDigitos(SumarDigitos(Year(D.FechaNacimiento)))

    '-----------------------------------------
    ' 2. PINÁCULOS (valores crudos)
    '-----------------------------------------

    pd.Pina1 = ReducirSimbolico(mes + dia)
    pd.Pina2 = ReducirSimbolico(dia + anio)
    pd.Pina3 = ReducirSimbolico((mes + dia) + (dia + anio)) '(PD.Pina1 + PD.Pina2)
    pd.Pina4 = ReducirSimbolico(mes + anio)

    '-----------------------------------------
    ' 3. DESAFÍOS (valores crudos)
    '-----------------------------------------

    pd.Desa1 = ReducirSimbolico(Abs(CInt(mes) - CInt(dia)))
    pd.Desa2 = ReducirSimbolico(Abs(CInt(dia) - CInt(anio)))
    pd.Desa3 = ReducirSimbolico(Abs(CInt((Abs(CInt(mes) - CInt(dia)) - Abs(CInt(dia) - CInt(anio))))))  '(PD.Desa1 - PD.Desa2))
    pd.Desa4 = ReducirSimbolico(Abs(CInt(mes) - CInt(anio)))

    '-----------------------------------------
    ' 4. TRANSICIONES
    '-----------------------------------------

    cv = ValorFinal(r.NumeroCaminoVida)

    primera = 36 - CByte(cv)
    segunda = primera + 9
    tercera = segunda + 9

    '-----------------------------------------
    ' 5. EDADES
    '-----------------------------------------

    pd.EdadIni1 = 0
    pd.EdadFin1 = primera

    pd.EdadIni2 = primera
    pd.EdadFin2 = segunda

    pd.EdadIni3 = segunda
    pd.EdadFin3 = tercera

    pd.EdadIni4 = tercera
    pd.EdadFin4 = 120

End Sub

Public Sub CalcularProgresiones(r As clsResultado, pr As clsProgresiones)

    Dim cv As Integer
    Dim AP As Integer
    Dim ep As Integer
    Dim prog As Integer

    '-----------------------------------------
    ' 1. EXTRAER VALORES CRUDOS
    '-----------------------------------------

    cv = ValorFinal(r.NumeroCaminoVida)
    AP = ValorFinal(r.AnioPersonal)
    ep = ValorFinal(r.EdadPersonal)

    '-----------------------------------------
    ' 2. PROGRESIÓN ACTUAL
    '-----------------------------------------

    prog = (ep + AP) Mod 9
    If prog = 0 Then prog = 9

    pr.ProgActual = prog

    '-----------------------------------------
    ' 3. PROGRESIÓN SIGUIENTE
    '-----------------------------------------

    prog = ((ep + 1) + AP) Mod 9
    If prog = 0 Then prog = 9

    pr.ProgSiguiente = prog

    '-----------------------------------------
    ' 4. PROGRESIÓN ANTERIOR
    '-----------------------------------------

    prog = ((ep - 1) + AP) Mod 9
    If prog <= 0 Then prog = prog + 9

    pr.ProgAnterior = prog

End Sub


Public Sub CalcularEsencia(r As clsResultado)

    Dim cv As Integer
    Dim AP As Integer
    Dim Esencia As Integer

    '-----------------------------------------
    ' 1. EXTRAER VALORES CRUDOS
    '-----------------------------------------

    cv = ValorFinal(r.NumeroCaminoVida)
    AP = ValorFinal(r.AnioPersonal)

    '-----------------------------------------
    ' 2. CÁLCULO DE LA ESENCIA
    '-----------------------------------------

    Esencia = cv + AP          ' no reducimos aquí
    'R.EsenciaCruda = Esencia   ' guardas el valor entero

    ' Forma simbólica para guardar / mostrar
    r.Esencia = ReducirSimbolico(Esencia)

    Debug.Print "Esencia: "; r.Esencia
End Sub

''------------------------------------------------------------
'' TRÁNSITOS UNIFICADOS (Moderno)
''------------------------------------------------------------
'Public Sub CalcularTransitos(ByRef T As clsTransitos, _
'                             ByVal P As clsPersona, _
'                             ByVal F As clsFonetica, _
'                             ByVal R As clsResultado)
'
''    Dim i As Integer
'    Dim sumaEspiritual As Integer
'    Dim txt As String
'    Dim colFon As Collection
'    Dim objFon As Variant
'    Dim esFon As Boolean
'
'    ' ¿Estamos en sistema fonético?
'    esFon = (F.ModoFonetico = 2)
'
'    ' Texto base (fonético o tradicional)
'    If esFon Then
'        txt = Trim$(F.FonNombre & " " & F.FonApe1 & " " & F.FonApe2)
'    Else
'        txt = Trim$(P.Nombre & " " & P.Ape1 & " " & P.Ape2)
'    End If
'
'    '-----------------------------------------
'    ' 1. TRÁNSITO FÍSICO
'    '-----------------------------------------
'    If esFon Then
'        T.Fisico = ReducirSimbolico(ConvertirFonemaANumero(R.primeraLetra))
'    Else
'        T.Fisico = ReducirSimbolico(ConvertirLetraANumero(R.primeraLetra, F.ModoFonetico))
'    End If
'    T.LetraFisico = R.primeraLetra
'
'    '-----------------------------------------
'    ' 2. TRÁNSITO MENTAL
'    '-----------------------------------------
'    If esFon Then
'        T.Mental = ReducirSimbolico(ConvertirFonemaANumero(R.primeraVocal))
'    Else
'        T.Mental = ReducirSimbolico(ConvertirLetraANumero(R.primeraVocal, F.ModoFonetico))
'    End If
'    T.LetraMental = R.primeraVocal
'    '-----------------------------------------
'    ' 3. TRÁNSITO EMOCIONAL
'    '-----------------------------------------
'    If esFon Then
'        T.Emocional = ReducirSimbolico(ConvertirFonemaANumero(R.primeraConsonante))
'    Else
'        T.Emocional = ReducirSimbolico(ConvertirLetraANumero(R.primeraConsonante, F.ModoFonetico))
'    End If
'    T.LetraEmocional = R.primeraConsonante
'    '-----------------------------------------
'    ' 4. TRÁNSITO ESPIRITUAL
'    '-----------------------------------------
'    If esFon Then
'        ' Fonético ? sumar fonemas
'        Set colFon = ExtraerFonemasFinales(txt)
'        For Each objFon In colFon
'            sumaEspiritual = sumaEspiritual + ConvertirFonemaANumero(CStr(objFon))
'        Next
'    Else
'        ' Tradicional --> sumar letras
'        Dim j As Integer, letra As String
'        For j = 1 To Len(txt)
'            letra = Mid$(txt, j, 1)
'            If letra Like "[A-ZÁÉÍÓÚÜÑ]" Then
'                sumaEspiritual = sumaEspiritual + ConvertirLetraANumero(letra, F.ModoFonetico)
'            End If
'        Next j
'    End If
'
'    T.Espiritual = ReducirSimbolico(sumaEspiritual)
'
'    With T
'        Debug.Print
'
'        Debug.Print "T. Físico: "; .LetraFisico; " "; .Fisico
'        Debug.Print "T. Mental: "; .LetraMental; " "; .Mental
'        Debug.Print "T. Emocional: "; .LetraEmocional; " "; .Emocional
'        Debug.Print "T. Espiritual: "; .Espiritual
'
'    End With
'
'
'End Sub

'--------------------------------------------------------------------------
' TRANSITOS SECUENCIALES (Tradicional)
'--------------------------------------------------------------------------
'--------------------------------------------------------------------------
' Bloque 3 (Final)
'--------------------------------------------------------------------------
Public Function GenerarTablaDecadas(ByVal D As clsDTOcalculos, r As clsResultado) As Variant
    
    Dim t(1 To 7, 1 To 11) As Variant
    Dim i As Integer
    Dim EdadTr As Byte
    Dim edadInicial As Integer
    Dim anio As Integer
    Dim AP As Integer
    Dim TE As Integer
    
    '-----------------------------------------
    ' Ventana centrada en la edad actual
    '-----------------------------------------
    'edad = GetEdad(FechaNac)
    edadInicial = r.Edad - 5
    
    If edadInicial < 0 Then edadInicial = 0
    
    For i = 1 To 11
        
        EdadTr = edadInicial + (i - 1)
        anio = Year(D.FechaNacimiento) + EdadTr
        
        '-----------------------------------------
        ' 1. Edad
        '-----------------------------------------
        t(1, i) = EdadTr
        
        '-----------------------------------------
        ' 2. Tránsitos secuenciales
        '-----------------------------------------
        t(2, i) = TransitoFisico(r, D, EdadTr)
        t(3, i) = TransitoMental(r, D, EdadTr)
        t(4, i) = TransitoEmocional(r, D, EdadTr)
        t(5, i) = TransitoEspiritual(r, D, EdadTr)
        
        ' Verificación emocional (si está vacío)
        If t(4, i) = "" Then t(4, i) = t(5, i)
        
        '-----------------------------------------
        ' 3. Año Personal
        '-----------------------------------------
        AP = Day(D.FechaNacimiento) + Month(D.FechaNacimiento) + SumarDigitos(anio)
        t(7, i) = AP
        
        '-----------------------------------------
        ' 4. Esencia = Año Personal + Tránsito Espiritual
        '-----------------------------------------
        TE = t(5, i)
        t(6, i) = SumarDigitos(AP + TE) ' ReducirSimbolico(AP + TE)
        
    Next i
    
    GenerarTablaDecadas = t
End Function

'--------------------------------------------------------------------------
' Bloque 2 (Cálculos)
'--------------------------------------------------------------------------
Public Function TransitoFisico(ByVal r As clsResultado, ByVal D As clsDTOcalculos, _
                               ByVal Edad As Byte) As Byte

    Dim texto As String
    Dim sec As Collection
    Dim elem As String
    Dim Valor As Integer
    
    texto = D.Nombre
    
    Set sec = ConstruirSecuencia(r, texto, D)
    elem = ElementoActivo(sec, Edad, D.SistemaFonetico, D.SistemaCalculo)
    Valor = ConvertirElementoANumero(elem, D.SistemaFonetico, D.SistemaCalculo)
    
    TransitoFisico = ReducirSimbolico(Valor)
End Function

Public Function TransitoMental(ByVal r As clsResultado, ByVal D As clsDTOcalculos, _
                               ByVal Edad As Byte) As Byte

    Dim texto As String
    Dim sec As Collection
    Dim elem As String
    Dim Valor As Integer
    
    texto = D.Ape1
    
    Set sec = ConstruirSecuencia(r, texto, D)
    elem = ElementoActivo(sec, Edad, D.SistemaFonetico, D.SistemaCalculo)
    Valor = ConvertirElementoANumero(elem, D.SistemaFonetico, D.SistemaCalculo)
    
    TransitoMental = ReducirSimbolico(Valor)
End Function

Public Function TransitoEmocional(ByVal r As clsResultado, ByVal D As clsDTOcalculos, _
                                  ByVal Edad As Byte) As Byte

    Dim texto As String
    Dim sec As Collection
    Dim elem As String
    Dim Valor As Integer
    
    If Trim$(D.Ape2) = "" Then
        TransitoEmocional = ""
        Exit Function
    End If

    If Trim$(D.Ape2) <> "" Then
        texto = D.Ape2
    Else
        texto = D.Nombre & " " & D.Ape1
    End If
    
    Set sec = ConstruirSecuencia(r, texto, D)
    elem = ElementoActivo(sec, Edad, D.SistemaFonetico, D.SistemaCalculo)
    Valor = ConvertirElementoANumero(elem, D.SistemaFonetico, D.SistemaCalculo)
    
    TransitoEmocional = ReducirSimbolico(Valor)
End Function

Public Function TransitoEspiritual(ByVal r As clsResultado, ByVal D As clsDTOcalculos, _
                                   ByVal Edad As Byte) As Byte

    Dim texto As String
    Dim sec As Collection
    Dim elem As String
    Dim Valor As Integer
    
    texto = Trim$(D.Nombre & " " & D.Ape1 & " " & D.Ape2)
    
    Set sec = ConstruirSecuencia(r, texto, D)
    elem = ElementoActivo(sec, Edad, D.SistemaFonetico, D.SistemaCalculo)
    Valor = ConvertirElementoANumero(elem, D.SistemaFonetico, D.SistemaCalculo)
    
    TransitoEspiritual = ReducirSimbolico(Valor)
End Function

Public Function CrearTransitosDesdeTabla(ByVal t As Variant, _
                                         ByVal D As clsDTOcalculos, _
                                         ByVal r As clsResultado) As Collection
    Dim col As New Collection
    Dim tr As clsTransito
    Dim i As Long
    Dim edadIni As Byte
    Dim anioNac As Integer
    
    If IsEmpty(t) Then
        Set CrearTransitosDesdeTabla = col
        Exit Function
    End If
    
    anioNac = Year(D.FechaNacimiento)
    
    For i = 1 To 11
        
        Set tr = New clsTransito
        
        edadIni = CByte(t(1, i))
        
        tr.orden = CByte(i)
        tr.Edad = edadIni
'        tr.EdadFin = edadIni + 1
        
        tr.anio = anioNac + edadIni
'        tr.AnioFin = anioNac + edadIni + 1
        
        tr.Fisico = NzByte(t(2, i))
        tr.Mental = NzByte(t(3, i))
        tr.Emocional = NzByte(t(4, i))
        tr.Espiritual = NzByte(t(5, i))
        tr.Esencia = NzByte(t(6, i))
        tr.AnioPersonal = NzByte(t(7, i))
        
        tr.SistemaFonetico = D.SistemaFonetico
        tr.SistemaCalculo = D.SistemaCalculo
        
        tr.IDPersona = D.IDPersona
        tr.IDFonetica = r.IDFonetica
        
        ' Marcar tránsito actual
        tr.EsActual = (tr.Edad = r.Edad)
        
        col.Add tr
    Next i
    
    Set CrearTransitosDesdeTabla = col
End Function

'Public Function CrearTransitosDesdeTabla(ByVal t As Variant, _
'                                         ByVal D As clsDTOcalculos, _
'                                         ByVal r As clsResultado) As Collection
'    Dim col As New Collection
'    Dim tr As clsTransito
'    Dim I As Long
'
'    ' Validación mínima
'    If IsEmpty(t) Then
'        Set CrearTransitosDesdeTabla = col
'        Exit Function
'    End If
'
'    ' t(1 To 7, 1 To 11)
'    For I = 1 To 11
'
'        Set tr = New clsTransito
'
'        ' -------------------------
'        ' Datos temporales
'        ' -------------------------
'        tr.Orden = CByte(I)
'        tr.anio = r.AnioInicio + (I - 1)
'        tr.Edad = CByte(t(1, I))
'
'        ' -------------------------
'        ' Valores reducidos
'        ' -------------------------
'        tr.Fisico = NzByte(t(2, I))
'        tr.Mental = NzByte(t(3, I))
'        tr.Emocional = NzByte(t(4, I))
'        tr.Espiritual = NzByte(t(5, I))
'        tr.Esencia = NzByte(t(6, I))
'        tr.AnioPersonal = NzByte(t(7, I))
'
'        ' -------------------------
'        ' Metadatos del cálculo
'        ' -------------------------
'        tr.MetodoFonetico = D.SistemaFonetico
'        tr.SistemaCalculo = D.SistemaCalculo
'
'        ' -------------------------
'        ' Trazabilidad fonética
'        ' -------------------------
'        tr.LetraFisico = r.FuenteFisico(I)
'        tr.LetraMental = r.FuenteMental(I)
'        tr.LetraEmocional = r.FuenteEmocional(I)
'        tr.LetraEspiritual = r.FuenteEspiritual(I)
'
'        tr.ValorFisicoBruto = r.ValorFisicoBruto(I)
'        tr.ValorMentalBruto = r.ValorMentalBruto(I)
'        tr.ValorEmocionalBruto = r.ValorEmocionalBruto(I)
'        tr.ValorEspiritualBruto = r.ValorEspiritualBruto(I)
'
'        tr.CicloFisico = r.CicloFisico(I)
'        tr.CicloMental = r.CicloMental(I)
'        tr.CicloEmocional = r.CicloEmocional(I)
'        tr.CicloEspiritual = r.CicloEspiritual(I)
'
'        ' -------------------------
'        ' Identificadores (si procede)
'        ' -------------------------
'        tr.IDPersona = D.IDPersona
''        tr.idResultado = r.idResultado
'        tr.IDFonetica = r.IDFonetica
'
'        ' -------------------------
'        ' Añadir a la colección
'        ' -------------------------
'        col.Add tr
'
'    Next I
'
'    Set CrearTransitosDesdeTabla = col
'End Function

'Public Function CrearTransitosDesdeTabla(ByVal t As Variant, _
'                                         ByVal D As clsDTOcalculos, _
'                                         ByVal r As clsResultado) As Collection
'    Dim col As New Collection
'    Dim i As Long
'    Dim tr As clsTransito
'    Dim Edad As Integer
'
'    ' Validación básica de la tabla
'    If IsEmpty(t) Then
'        Set CrearTransitosDesdeTabla = col
'        Exit Function
'    End If
'
'    ' Asumimos t(1 To 7, 1 To 11)
'    For i = 1 To 11
'
'        Set tr = New clsTransito
'
'        ' Edad
'        Edad = t(1, i)
'        tr.Edad = Edad
'
'        ' Valores principales
'        tr.Fisico = NzByte(t(2, i))
'        tr.Mental = NzByte(t(3, i))
'        tr.Emocional = NzByte(t(4, i))
'        tr.Espiritual = NzByte(t(5, i))
'        tr.Esencia = NzByte(t(6, i))
'        tr.AnioPersonal = NzByte(t(7, i))
'
'        ' Metadatos de contexto
'        tr.MetodoFonetico = D.ModoFonetico
'        tr.SistemaCalculo = D.SistemaCalculo
'
'        ' Trazabilidad simbólica mínima
'        tr.FuenteFisico = D.Nombre
'        tr.FuenteMental = D.Ape1
'        tr.FuenteEmocional = D.Ape2
'        tr.FuenteEspiritual = Trim$(D.Nombre & " " & D.Ape1 & " " & D.Ape2)
'
'        ' Validación suave: si todo está vacío, no lo añadimos
'        If Not EsTransitoVacio(tr) Then
'            col.Add tr
'        End If
'
'    Next i
'
'    Set CrearTransitosDesdeTabla = col
'End Function

Private Function NzByte(v As Variant) As Byte
    If IsNull(v) Or v = "" Then
        NzByte = 0
    Else
        NzByte = CByte(v)
    End If
End Function

Private Function EsTransitoVacio(ByVal tr As clsTransito) As Boolean
    EsTransitoVacio = (tr.Fisico = 0 And _
                       tr.Mental = 0 And _
                       tr.Emocional = 0 And _
                       tr.Espiritual = 0 And _
                       tr.Esencia = 0 And _
                       tr.AnioPersonal = 0)
End Function

'Public Function CrearTransitosDesdeTabla(t As Variant, _
'                                         D As clsDTOcalculos, _
'                                         r As clsResultado) As Collection
'    Dim col As New Collection
'    Dim tr As clsTransitos
'    Dim i As Integer
'    Dim Edad As Integer
'    Dim anio As Integer
'    Dim IDPersona As Long
'    Dim idResultado  As Long
'
'    IDPersona = D.ID_Persona
'    idResultado = r.idResultado
'
'    For i = 1 To 11
'
'        Set tr = New clsTransitos
'
'        ' Identificadores
'        tr.IDPersona = IDPersona
'        tr.idResultado = idResultado
'        tr.Orden = i
'
'        ' Edad y año
'        Edad = t(1, i)
'        tr.Edad = Edad
'        tr.anio = Year(D.FechaNacimiento) + Edad
'
'        ' Valores numéricos
'        tr.Fisico = t(2, i)
'        tr.Mental = t(3, i)
'        tr.Emocional = t(4, i)
'        tr.Espiritual = t(5, i)
'
'        tr.Esencia = t(6, i)
'        tr.AñoPersonal = t(7, i)
'
'        ' Letras asociadas
'        tr.LetraFisico = LetraTransito(D.Nombre, D, Edad)
'        tr.LetraMental = LetraTransito(D.Ape1, D, Edad)
'
'        If Trim$(D.Ape2) <> "" Then
'            tr.LetraEmocional = LetraTransito(D.Ape2, D, Edad)
'        Else
'            tr.LetraEmocional = LetraTransito(D.Nombre & " " & D.Ape1, D, Edad)
'        End If
'
'        tr.LetraEspiritual = LetraTransito(D.Nombre & " " & D.Ape1 & " " & D.Ape2, D, Edad)
'
'        ' Añadir a la colección
'        col.Add tr
'    Next i
'
'    Set CrearTransitosDesdeTabla = col
'End Function

Public Function LetraTransito(r As clsResultado, texto As String, D As clsDTOcalculos, Edad As Integer) As String
    Dim sec As Collection
    Dim elem As String
    
    Set sec = ConstruirSecuencia(r, texto, D)
    elem = ElementoActivo(sec, Edad, D.SistemaFonetico, D.SistemaCalculo)
    
    LetraTransito = elem
End Function

'--------------------------------------------------------------------------
' Bloque 1 Auxiliares
'--------------------------------------------------------------------------
Private Function ConstruirSecuencia(r As clsResultado, ByVal texto As String, _
                                   ByVal D As clsDTOcalculos) As Collection
    Dim col As New Collection
    Dim base As New Collection
    Dim i As Long
    Dim c As String
    Dim Valor As Integer
    Dim rep As Integer
    Dim maxPos As Long
    Dim ciclo As Long

    ' 1. Construir el patrón base expandido (una sola vez)
    For i = 1 To Len(texto)
        c = Mid$(texto, i, 1)

        If c = " " Then GoTo ContinueLetter

        Valor = ConvertirElementoANumero(c, D.SistemaFonetico, D.SistemaCalculo)

        For rep = 1 To Valor
            base.Add c
        Next rep

ContinueLetter:
    Next i

    ' 2. Determinar cuántas posiciones necesitamos
    maxPos = r.Edad + 5   ' edad actual + 5

    ' 3. Repetir el patrón base hasta cubrir maxPos
    ciclo = 1
    Do While col.Count < maxPos
        For i = 1 To base.Count
            col.Add base(i)
            If col.Count >= maxPos Then Exit Do
        Next i
        ciclo = ciclo + 1
    Loop

    Set ConstruirSecuencia = col
End Function


'Private Function ConstruirSecuencia(ByVal Texto As String, _
'                                   ByVal D As clsDTOcalculos) As Collection
'    Dim col As New Collection
'    Dim i As Long
'    Dim c As String
'    Dim c2 As String
'
'    If D.ModoFonetico = Fonetico Then
'        '-----------------------------------------
'        ' MÉTODO FONÉTICO – usar fonemas
'        '-----------------------------------------
'        Dim fon As Collection
'        Set fon = ExtraerFonemasFinales(Texto)
'
'        For Each c In fon
'            col.Add c
'        Next c
'
'    Else
'        '-----------------------------------------
'        ' MÉTODO TRADICIONAL – usar letras y dígrafos
'        '-----------------------------------------
'        i = 1
'        Do While i <= Len(Texto)
'
'            c = Mid$(Texto, i, 1)
'
'            ' Saltar espacios
'            If c = " " Then
'                i = i + 1
'                GoTo ContinueLoop
'            End If
'
'            ' Detectar dígrafos CH, LL, RR
'            If i < Len(Texto) Then
'                c2 = Mid$(Texto, i, 2)
'
'                If c2 = "CH" Or c2 = "LL" Or c2 = "RR" Then
'                    col.Add c2
'                    i = i + 2
'                    GoTo ContinueLoop
'                End If
'            End If
'
'            ' Añadir letra individual
'            col.Add c
'            i = i + 1
'
'ContinueLoop:
'        Loop
'    End If
'
'    Set ConstruirSecuencia = col
'End Function


'Private Function ConstruirSecuencia(ByVal Texto As String, _
'                                   ByVal D As clsDTOcalculos) As Collection
'    Dim col As New Collection
'    'Dim f As Variant
'    Dim i As Variant 'Integer
'
'    If D.ModoFonetico = 2 Then
'        '-----------------------------------------
'        ' MÉTODO FONÉTICO ? usar fonemas
'        '-----------------------------------------
'        Dim fon As Collection
'        Set fon = ExtraerFonemasFinales(Texto)
'
'        For Each i In fon
'            col.Add i
'        Next i
'
'    Else
'        '-----------------------------------------
'        ' MÉTODO TRADICIONAL ? usar letras
'        '-----------------------------------------
'        For i = 1 To Len(Texto)
'            Dim c As String
'            c = Mid$(Texto, i, 1)
'
'            If c Like "[A-ZÁÉÍÓÚÜÑ]" Then
'                col.Add c
'            End If
'        Next i
'    End If
'
'    Set ConstruirSecuencia = col
'End Function


Public Function ConvertirElementoANumero(ByVal Elemento As String, _
                                         ByVal Sistema As ModoFonetico, modoCalc As ModoCalculo) As Integer
    If Sistema = mfFonetico Then
        ConvertirElementoANumero = ConvertirFonemaANumero(Elemento)
    Else
        ConvertirElementoANumero = ConvertirLetraANumero(Elemento, Sistema, modoCalc)
    End If
End Function

Public Function ElementoActivo(ByVal secuencia As Collection, _
                               ByVal Edad As Byte, _
                               ByVal metodo As ModoFonetico, _
                               ByVal Modo As ModoCalculo) As String
    Dim total As Integer
    Dim i As Integer
    Dim Valor As Integer
    
    total = 0
    
    Do
        For i = 1 To secuencia.Count
            Valor = ConvertirElementoANumero(secuencia(i), metodo, Modo)
            
            If Edad < total + Valor Then
                ElementoActivo = secuencia(i)
                Exit Function
            End If
            
            total = total + Valor
        Next i
    Loop
End Function

'--------------------------------------------------------------------------
'--------------------------------------------------------------------------

Public Sub ExportarTablaDecadasHTML(ByVal Tabla As Variant, ByVal Ruta As String)

    Dim f As Integer
    Dim i As Integer, j As Integer
    
    f = FreeFile
    
    Open Ruta For Output As #f
    
    Print #f, "<html><head><meta charset='UTF-8'></head><body>"
    Print #f, "<table border='1' cellspacing='0' cellpadding='4'>"
    
    '-----------------------------------------
    ' Encabezados de columnas (Años 1..10)
    '-----------------------------------------
    Print #f, "<tr><th></th>"
    For j = 1 To 10
        Print #f, "<th>Año " & j & "</th>";
    Next j
    Print #f, "</tr>"
    
    '-----------------------------------------
    ' Filas (7 conceptos)
    '-----------------------------------------
    Dim etiquetas As Variant
    etiquetas = Array("Edad", "Tránsito Físico", "Tránsito Mental", _
                      "Tránsito Emocional", "Tránsito Espiritual", _
                      "Esencia", "Año Personal")
    
    For i = 1 To 7
        Print #f, "<tr><td><b>" & etiquetas(i - 1) & "</b></td>";
        
        For j = 1 To 10
            Print #f, "<td>" & Tabla(i, j) & "</td>";
        Next j
        
        Print #f, "</tr>"
    Next i
    
    Print #f, "</table></body></html>"
    
    Close #f

End Sub


'=====================================================================================
' FUNCIONES PUBLICAS ADICIONALES
'=====================================================================================

Public Function CollectionToString(col As Collection, Optional sep As String = ",") As String
    Dim item As Variant
    Dim s As String
    
    For Each item In col
        If s = "" Then
            s = CStr(item)
        Else
            s = s & sep & CStr(item)
        End If
    Next item
    
    CollectionToString = s
End Function

Public Function GetEdad(FechaNac As Date) As Byte

    Dim hoy As Date
    Dim cumple As Date
    Dim anio As Integer
    Dim mes As Integer
    Dim dia As Integer
    Dim ultimoDiaMes As Integer

    hoy = Date
    anio = Year(hoy)
    mes = Month(FechaNac)
    dia = Day(FechaNac)

    ' Último día del mes del cumpleaños este año
    ultimoDiaMes = Day(DateSerial(anio, mes + 1, 0))

    If dia > ultimoDiaMes Then
        dia = ultimoDiaMes
    End If

    cumple = DateSerial(anio, mes, dia)

    ' Si aún no ha llegado el cumpleaños este año, restar 1
    If cumple > hoy Then
        anio = anio - 1
    End If

    GetEdad = anio - Year(FechaNac)

End Function

Public Function ValorFinal(cadena As String) As Byte

    ValorFinal = CByte(Split(cadena, "/")(UBound(Split(cadena, "/"))))

End Function


'=====================================================================================
' FUNCIONES PRIVADAS INTERNAS
'=====================================================================================

Private Function ZeroAcum() As tAcumuladores
    Dim a As tAcumuladores
    a.Vocales = 0
    a.Consonantes = 0
    a.Completo = 0
    ZeroAcum = a
End Function

'Private Sub CargarAcumuladoresNombre(P As clsPersona, f As clsFonetica)
Private Sub CargarAcumuladoresNombre(Incl As clsInclusion, D As clsDTOcalculos)

    Dim Sistema As ModoFonetico
    Dim Modo As ModoCalculo

    ' Inicializar acumuladores
    mAcumNombre = ZeroAcum()
    mAcumApe1 = ZeroAcum()
    mAcumApe2 = ZeroAcum()

    Modo = D.SistemaFonetico
    Sistema = D.SistemaFonetico
    
    mAcumNombre = CalculadorNumeros(Incl, D.Nombre, Sistema, Modo)
    mAcumApe1 = CalculadorNumeros(Incl, D.Ape1, Sistema, Modo)
    mAcumApe2 = CalculadorNumeros(Incl, D.Ape2, Sistema, Modo)

End Sub

Private Function CalculadorNumeros(Incl As clsInclusion, Parte As String, Sistema As ModoFonetico, Modo As ModoCalculo) As tAcumuladores

    Dim i As Long
    Dim letra As String
    Dim Valor As Integer
    Dim colFonemas As Collection
    Dim f As Variant

    Dim Ac As tAcumuladores   ' ? este será el resultado final

    ' ============================================================================
    ' SISTEMA 2 ? FONÉTICO
    ' ============================================================================
    If Sistema = mfFonetico Then
        
        ' Obtener fonemas ya normalizados
        Set colFonemas = ExtraerFonemasFinales(Parte)

        For Each f In colFonemas

            Valor = ConvertirFonemaANumero(CStr(f))

            If Valor > 0 Then
            
                Call Incl.AddNumero(Valor)
                
                Ac.Completo = Ac.Completo + Valor

                If EsVocal(CStr(f)) Then
                    Ac.Vocales = Ac.Vocales + Valor
                Else
                    Ac.Consonantes = Ac.Consonantes + Valor
                End If

            End If

        Next f

        CalculadorNumeros = Ac
        Exit Function
    End If

    ' ============================================================================
    ' SISTEMAS 1 ? TRADICIONAL
    ' ============================================================================
    For i = 1 To Len(Parte)

        letra = NormalizarLetraTradicional(Mid$(Parte, i, 1))
        Valor = ConvertirLetraANumero(letra, Sistema, Modo)

        If Valor > 0 Then
            
            Call Incl.AddNumero(Valor)
            
            Ac.Completo = Ac.Completo + Valor

            If EsVocal(letra) Then
                Ac.Vocales = Ac.Vocales + Valor
            Else
                Ac.Consonantes = Ac.Consonantes + Valor
            End If

        End If

    Next i

    CalculadorNumeros = Ac

End Function

' ============================================================================
'  VOCAL FONÉTICA Y TRADICIONAL
' ============================================================================

Private Function EsVocal(letra As String) As Boolean

    ' Normalizar a mayúscula para unificar
    letra = UCase$(letra)

    ' Vocales tradicionales (ya normalizadas)
    Select Case letra
        Case "A", "E", "I", "O", "U"
            EsVocal = True
        Case Else
            EsVocal = False
    End Select

End Function

