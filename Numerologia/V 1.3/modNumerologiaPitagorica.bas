Attribute VB_Name = "modNumerologiaPitagorica"

Option Compare Database
Option Explicit

Public Type tInterCadena
    numero As Byte
    Final As Byte
    EsMaestro As Boolean
    EsKarmico As Boolean
End Type

Dim res As tInterCadena



' ============================================================
'   DETECCIÓN DE MAESTROS
' ============================================================
Public Function EsMaestro(ByVal valor As Byte) As Boolean
    Select Case valor
        Case 11, 22, 33, 44
            EsMaestro = True
    End Select
End Function


' ============================================================
'   DETECCIÓN DE KÁRMICOS
' ============================================================
Public Function EsKarmico(ByVal valor As Byte) As Boolean
    Select Case valor
        Case 13, 14, 16, 19
            EsKarmico = True
    End Select
End Function


' ============================================================
'   SUMA DE DÍGITOS (reducción básica)
' ============================================================
Private Function SumaDigitos(ByVal n As Integer) As Integer
    Dim s As Integer
    Do While n > 0
        s = s + (n Mod 10)
        n = n \ 10
    Loop
    SumaDigitos = s
End Function


' ============================================================
'   MOTOR PRINCIPAL DE ANÁLISIS NUMEROLÓGICO
' ============================================================
Public Function AnalizarNumero(ByVal valor As Integer) As tResultado
    Dim r As tResultado
    Dim actual As Integer
    Dim cadenaInterna As String
    Dim partes() As String
    Dim i As Long
    Dim v As Byte
    
    ' Guardar inicial real
    r.Inicial = valor
    actual = valor
    cadenaInterna = CStr(actual)
    
    ' ----------------------------------------------------------
    ' 1) Reducir hasta 2 dígitos ? valor medio
    ' ----------------------------------------------------------
    Do While actual > 99
        actual = SumaDigitos(actual)
        cadenaInterna = cadenaInterna & "/" & CStr(actual)
    Loop
    
    If actual > 9 Then
        r.Medio = CByte(actual)
    End If
    
    ' ----------------------------------------------------------
    ' 2) Reducir hasta 1 dígito ? valor final
    ' ----------------------------------------------------------
    Do While actual > 9
        actual = SumaDigitos(actual)
        cadenaInterna = cadenaInterna & "/" & CStr(actual)
    Loop
    
    r.Final = CByte(actual)
    
    ' ----------------------------------------------------------
    ' 3) Detectar Maestro y Kármico en cualquier etapa
    ' ----------------------------------------------------------
    partes = Split(cadenaInterna, "/")
    
    For i = 0 To UBound(partes)
        v = CByte(partes(i))
        
        If EsMaestro(v) Then r.Maestro = v
        If EsKarmico(v) Then r.Karmico = v
    Next i
    
    ' ----------------------------------------------------------
    ' 4) Generar presentación final
    ' ----------------------------------------------------------
    Call FormatearResultado(r)
    
    AnalizarNumero = r
End Function


' ============================================================
'   PRESENTACIÓN FINAL (rellena r.Cadena)
' ============================================================
Public Sub FormatearResultado(r As tResultado)
    Dim mostrarInicial As Boolean
    mostrarInicial = (r.Inicial <= 999)
    
    Dim tieneEspecial As Boolean
    Dim valorEspecial As Byte
    
    tieneEspecial = (r.Maestro > 0 Or r.Karmico > 0)
    
    If tieneEspecial Then
        If r.Maestro > 0 Then
            valorEspecial = r.Maestro
        Else
            valorEspecial = r.Karmico
        End If
    End If
    
    ' ----------------------------------------------------------
    ' CASO A: Inicial <= 999 ? se muestra el inicial
    ' ----------------------------------------------------------
    If mostrarInicial Then
        
        r.Cadena = CStr(r.Inicial)
        
        ' Si hay maestro o kármico ? inicial/especial/final
        If tieneEspecial Then
            r.Cadena = r.Cadena & "/" & valorEspecial & "/" & r.Final
            Exit Sub
        End If
        
        ' Si no hay especial ? inicial/medio/final
        If r.Medio > 0 Then
            r.Cadena = r.Cadena & "/" & r.Medio
        End If
        
        r.Cadena = r.Cadena & "/" & r.Final
        Exit Sub
    End If
    
    ' ----------------------------------------------------------
    ' CASO B: Inicial >= 1000 ? NO se muestra el inicial
    ' ----------------------------------------------------------
    If tieneEspecial Then
        r.Cadena = valorEspecial & "/" & r.Final
        Exit Sub
    End If
    
    If r.Medio > 0 Then
        r.Cadena = r.Medio & "/" & r.Final
    Else
        r.Cadena = CStr(r.Final)
    End If
End Sub



Public Function InterpretarCadenaResultado(Cadena As String) As tInterCadena
    Dim arr() As String
    Dim res As tInterCadena
    Dim n As Integer
    
    arr = Split(Cadena, "/")
    n = UBound(arr)
    
    ' Final siempre es el último
'    res.Final = Trim(arr(n))
    
    Select Case n
        Case 0
            ' Solo un número ? simple
            res.Original = res.Final
            res.numero = arr(0)
        Case 1
            ' Dos elementos --> maestro o kármico + final
            res.MaestroOKarmico = Trim(arr(0))
'            res.Final = Trim(arr(1))
'            res.Original = res.MaestroOKarmico
        
        Case 2
            ' Tres elementos --> original / maestro-kármico / final
            res.Original = Trim(arr(0))
            res.MaestroOKarmico = Trim(arr(1))
            res.Final = Trim(arr(2))
    End Select
    
    ' Detectar maestro
    Select Case res.MaestroOKarmico
        Case "11", "22", "33", "44"
            res.EsMaestro = True
    End Select
    
    ' Detectar kármico
    Select Case res.MaestroOKarmico
        Case "13", "14", "16", "19"
            res.EsKarmico = True
    End Select
    
    InterpretarCadenaResultado = res
End Function


'------------------------------------------------------------------------
'Public Type tInterCadena
'    Numero As Byte
'    Final As Byte
'    EsMaestro As Boolean
'    EsKarmico As Boolean
'End Type
'
'Dim res As tInterCadena

Public Sub InterCadRes(ByVal Cadena As String)
    Dim arrCad As Variant
    Dim n As Integer
    
    arrCad = Split(Cadena, "/")
    
    With res
        .EsMaestro = False
        .EsKarmico = False
        If UBound(arrCad) > 0 Then ' Tiene 2 o 3 elementos
            For n = 0 To UBound(arrCad) - 1
                Select Case CByte(arrCad)
                    Case 11, 22, 33, 44
                        .numero = CByte(arrCad)
                        .EsMaestro = True
                        Exit For
                    Case 13, 14, 16, 19
                        .numero = CByte(arrCad)
                        .EsKarmico = True
                        Exit For
                    Case Else
                        .numero = CByte(arrCad)
                End Select
            Next n
        End If
        .Final = CByte(arrCad(UBound(arrCad)))
    End With
End Sub

Public Function NumeroInforme(ByVal Cadena As String) As Byte
    
    Call InterCadRes(Cadena)
    
    NumeroInforme = res.Final
    
    If res.numero > 0 Then
        NumeroInforme = res.numero
    End If

End Function

Public Function NumeroSinastria(Cadena As String) As Byte
    
    Call InterCadRes(Cadena)
    
    NumeroSinastria = res.Final
    
    If res.EsMaestro Then
        NumeroSinastria = res.numero
    End If
    
End Function

'-----------------------------------------------------
Public Function ObtenerNumeroParaInforme(Cadena As String) As String
    Dim r As tInterpretacionCadena
    r = InterpretarCadenaResultado(Cadena)
    
    If r.EsMaestro Then
        ObtenerNumeroParaInforme = r.MaestroOKarmico
        Exit Function
    End If
    
    If r.EsKarmico Then
        ObtenerNumeroParaInforme = r.MaestroOKarmico
        Exit Function
    End If
    
    ObtenerNumeroParaInforme = r.Final
End Function

Public Function ObtenerNumeroParaSinastria(Cadena As String) As String
    Dim r As tInterpretacionCadena
    r = InterpretarCadenaResultado(Cadena)
    
    If r.EsMaestro Then
        ObtenerNumeroParaSinastria = r.MaestroOKarmico
        Exit Function
    End If
    
    ' Kármico ? usar final
    ObtenerNumeroParaSinastria = r.Final
End Function






'' ============================================================
''   TIPO DE RESULTADO NUMEROLÓGICO
'' ============================================================
'Public Type tResultado
'    cadena As String      ' Ej: "78/15/6"
'    Inicial As Integer    ' Valor inicial bruto
'    Medio As Byte         ' Primera reducción a 2 dígitos (si existe)
'    Final As Byte         ' Reducción final a 1 dígito
'    Maestro As Byte       ' 11,22,33,44 o 0
'    Karmico As Byte       ' 13,14,16,19 o 0
'End Type
'
'
'' ============================================================
''   DETECCIÓN DE MAESTROS
'' ============================================================
'Public Function EsMaestro(ByVal valor As Byte) As Boolean
'    Select Case valor
'        Case 11, 22, 33, 44
'            EsMaestro = True
'    End Select
'End Function
'
'
'' ============================================================
''   DETECCIÓN DE KÁRMICOS
'' ============================================================
'Public Function EsKarmico(ByVal valor As Byte) As Boolean
'    Select Case valor
'        Case 13, 14, 16, 19
'            EsKarmico = True
'    End Select
'End Function
'
'
'' ============================================================
''   SUMA DE DÍGITOS (reducción básica)
'' ============================================================
'Private Function SumaDigitos(ByVal n As Integer) As Integer
'    Dim s As Integer
'    Do While n > 0
'        s = s + (n Mod 10)
'        n = n \ 10
'    Loop
'    SumaDigitos = s
'End Function
'
'
'' ============================================================
''   MOTOR PRINCIPAL DE ANÁLISIS NUMEROLÓGICO
'' ============================================================
'Public Function AnalizarNumero(ByVal valor As Integer) As tResultado
'    Dim res As tResultado
'    Dim cadena As String
'    Dim actual As Integer
'    Dim partes() As String
'    Dim i As Long
'    Dim v As Byte
'
'    ' Guardar inicial real
'    res.Inicial = valor
'    actual = valor
'    cadena = CStr(actual)
'
'    ' ----------------------------------------------------------
'    ' 1) Reducir hasta 2 dígitos ? valor medio
'    ' ----------------------------------------------------------
'    Do While actual > 99
'        actual = SumaDigitos(actual)
'        cadena = cadena & "/" & CStr(actual)
'    Loop
'
'    ' Si el resultado es de 2 dígitos, es el valor medio
'    If actual > 9 Then
'        res.Medio = CByte(actual)
'    End If
'
'    ' ----------------------------------------------------------
'    ' 2) Reducir hasta 1 dígito ? valor final
'    ' ----------------------------------------------------------
'    Do While actual > 9
'        actual = SumaDigitos(actual)
'        cadena = cadena & "/" & CStr(actual)
'    Loop
'
'    res.Final = CByte(actual)
'    res.cadena = cadena
'
'    ' ----------------------------------------------------------
'    ' 3) Detectar Maestro y Kármico en cualquier etapa
'    ' ----------------------------------------------------------
'    partes = Split(cadena, "/")
'
'    For i = 0 To UBound(partes)
'        v = CByte(partes(i))
'
'        If EsMaestro(v) Then res.Maestro = v
'        If EsKarmico(v) Then res.Karmico = v
'    Next i
'
'    AnalizarNumero = res
'End Function
'
'Public Function FormatearResultado(r As tResultado) As String
'    Dim mostrarInicial As Boolean
'    mostrarInicial = (r.Inicial <= 999)
'
'    Dim tieneEspecial As Boolean
'    Dim valorEspecial As Byte
'
'    tieneEspecial = (r.Maestro > 0 Or r.Karmico > 0)
'
'    If tieneEspecial Then
'        If r.Maestro > 0 Then
'            valorEspecial = r.Maestro
'        Else
'            valorEspecial = r.Karmico
'        End If
'    End If
'
'    ' ----------------------------------------------------------
'    ' CASO A: Inicial <= 999 ? se muestra el inicial
'    ' ----------------------------------------------------------
'    If mostrarInicial Then
'
'        ' Siempre empezamos por el inicial
'        FormatearResultado = CStr(r.Inicial)
'
'        ' Si hay maestro o kármico ? inicial/especial/final
'        If tieneEspecial Then
'            FormatearResultado = FormatearResultado & "/" & valorEspecial & "/" & r.Final
'            Exit Function
'        End If
'
'        ' Si no hay especial ? inicial/medio/final
'        If r.Medio > 0 Then
'            FormatearResultado = FormatearResultado & "/" & r.Medio
'        End If
'
'        FormatearResultado = FormatearResultado & "/" & r.Final
'        Exit Function
'    End If
'
'    ' ----------------------------------------------------------
'    ' CASO B: Inicial >= 1000 ? NO se muestra el inicial
'    ' ----------------------------------------------------------
'    If tieneEspecial Then
'        FormatearResultado = valorEspecial & "/" & r.Final
'        Exit Function
'    End If
'
'    If r.Medio > 0 Then
'        FormatearResultado = r.Medio & "/" & r.Final
'    Else
'        FormatearResultado = r.Final
'    End If
'End Function
'
'
'
'
''' ============================================================
'''   PRESENTACIÓN FINAL SEGÚN TUS REGLAS ORIGINALES
''' ============================================================
''Public Function FormatearResultado(r As tResultado) As String
''
''    ' Caso 1: inicial maestro ? mostrar inicial/medio/final
''    If EsMaestro(CByte(r.Inicial)) Then
''        If r.Medio > 0 Then
''            FormatearResultado = r.Inicial & "/" & r.Medio & "/" & r.Final
''        Else
''            FormatearResultado = r.Inicial & "/" & r.Final
''        End If
''        Exit Function
''    End If
''
''    ' Caso 2: inicial kármico ? mostrar inicial/medio/final
''    If EsKarmico(CByte(r.Inicial)) Then
''        If r.Medio > 0 Then
''            FormatearResultado = r.Inicial & "/" & r.Medio & "/" & r.Final
''        Else
''            FormatearResultado = r.Inicial & "/" & r.Final
''        End If
''        Exit Function
''    End If
''
''    ' Caso 3: inicial de 2 dígitos ? mostrar inicial/medio/final
''    If r.Inicial >= 10 And r.Inicial <= 99 Then
''        If r.Medio > 0 Then
''            FormatearResultado = r.Inicial & "/" & r.Medio & "/" & r.Final
''        Else
''            FormatearResultado = r.Inicial & "/" & r.Final
''        End If
''        Exit Function
''    End If
''
''    ' Caso 4: inicial de 3+ dígitos ? mostrar medio/final
''    If r.Inicial > 99 Then
''        If r.Medio > 0 Then
''            FormatearResultado = r.Medio & "/" & r.Final
''        Else
''            FormatearResultado = r.Final
''        End If
''        Exit Function
''    End If
''
''    ' Caso 5: inicial de 1 dígito ? solo final
''    FormatearResultado = CStr(r.Final)
''End Function
''
