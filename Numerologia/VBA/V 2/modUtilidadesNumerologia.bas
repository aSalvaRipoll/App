Attribute VB_Name = "modUtilidadesNumerologia"
Option Compare Database
Option Explicit

' ============================================================================
' Proyecto:     Sistema de Numerología Tradicional y Fonético
' Módulo: modUtilidadesNumerologia
' Descripción: Funciones de utilidad para cálculos numerológicos
' Autor: Sistema de Numerología
' Fecha: 2024
' ============================================================================

' ============================================================================
' Función: ReducirADigito
' Descripción: Reduce un número a un dígito simple (1-9) o número maestro
' Parámetros:
'   - numero: El número a reducir
'   - permitirMaestros: Si True, preserva números maestros (11, 22, 33, 44)
' Retorna: Número reducido
' ============================================================================
Public Function ReducirADigito(ByVal numero As Long, Optional ByVal permitirMaestros As Boolean = True) As Integer
    Dim suma As Long
    Dim digito As String
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    ' Si ya es un dígito simple, retornar
    If numero >= 1 And numero <= 9 Then
        ReducirADigito = CInt(numero)
        Exit Function
    End If
    
    ' Si es número maestro y se permiten, retornar
    If permitirMaestros Then
        If numero = NUM_MAESTRO_11 Or numero = NUM_MAESTRO_22 Or _
           numero = NUM_MAESTRO_33 Or numero = NUM_MAESTRO_44 Then
            ReducirADigito = CInt(numero)
            Exit Function
        End If
    End If
    
    ' Reducir sumando dígitos
    Do While numero > 9
        suma = 0
        For i = 1 To Len(CStr(numero))
            digito = Mid(CStr(numero), i, 1)
            suma = suma + CLng(digito)
        Next i
        
        numero = suma
        
        ' Verificar si es número maestro
        If permitirMaestros Then
            If numero = NUM_MAESTRO_11 Or numero = NUM_MAESTRO_22 Or _
               numero = NUM_MAESTRO_33 Or numero = NUM_MAESTRO_44 Then
                Exit Do
            End If
        End If
    Loop
    
    ReducirADigito = CInt(numero)
    Exit Function
    
ErrorHandler:
    err.Raise vbObjectError + 1001, "ReducirADigito", ERR_CALCULO_FALLIDO & ": " & err.Description
End Function

' ============================================================================
' Función: EsNumeroMaestro
' Descripción: Verifica si un número es un número maestro
' Parámetros:
'   - numero: El número a verificar
' Retorna: True si es número maestro, False en caso contrario
' ============================================================================
Public Function EsNumeroMaestro(ByVal numero As Integer) As Boolean
    EsNumeroMaestro = (numero = NUM_MAESTRO_11 Or numero = NUM_MAESTRO_22 Or _
                       numero = NUM_MAESTRO_33 Or numero = NUM_MAESTRO_44)
End Function

' ============================================================================
' Función: EsNumeroKarmico
' Descripción: Verifica si un número es un número kármico
' Parámetros:
'   - numero: El número a verificar
' Retorna: True si es número kármico, False en caso contrario
' ============================================================================
Public Function EsNumeroKarmico(ByVal numero As Integer) As Boolean
    EsNumeroKarmico = (numero = NUM_KARMICO_13 Or numero = NUM_KARMICO_14 Or _
                       numero = NUM_KARMICO_16 Or numero = NUM_KARMICO_19)
End Function

' ============================================================================
' Función: ObtenerValorLetra
' Descripción: Obtiene el valor numérico de una letra según el sistema Pitagórico
' Parámetros:
'   - letra: La letra a convertir
' Retorna: Valor numérico de la letra (1-9)
' ============================================================================
Public Function ObtenerValorLetra(ByVal letra As String) As Integer
    Dim letraUpper As String
    
    On Error GoTo ErrorHandler
    
    ' Convertir a mayúscula y tomar solo el primer carácter
    letraUpper = UCase(Trim(Left(letra, 1)))
    
    ' Verificar que sea una letra
    If Len(letraUpper) = 0 Then
        ObtenerValorLetra = 0
        Exit Function
    End If
    
    ' Asignar valor según el sistema Pitagórico
    Select Case letraUpper
        Case "A", "J", "S": ObtenerValorLetra = 1
        Case "B", "K", "T": ObtenerValorLetra = 2
        Case "C", "L", "U": ObtenerValorLetra = 3
        Case "D", "M", "V": ObtenerValorLetra = 4
        Case "E", "N", "W", "Ñ": ObtenerValorLetra = 5
        Case "F", "O", "X": ObtenerValorLetra = 6
        Case "G", "P", "Y": ObtenerValorLetra = 7
        Case "H", "Q", "Z": ObtenerValorLetra = 8
        Case "I", "R": ObtenerValorLetra = 9
        Case "Ç": ObtenerValorLetra = 3
        Case Else: ObtenerValorLetra = 0
    End Select
    
    Exit Function
    
ErrorHandler:
    ObtenerValorLetra = 0
End Function

' ============================================================================
' Función: EsVocal
' Descripción: Verifica si una letra es vocal
' Parámetros:
'   - letra: La letra a verificar
'   - contexto: Palabra completa para determinar si Y es vocal
' Retorna: True si es vocal, False en caso contrario
' ============================================================================
Public Function EsVocal(ByVal letra As String, Optional ByVal contexto As String = "") As Boolean
    Dim letraUpper As String
    
    letraUpper = UCase(Trim(Left(letra, 1)))
    
    ' Vocales estándar
    If InStr(1, "AEIOUÁÉÍÓÚÄËÏÖÜ", letraUpper) > 0 Then
        EsVocal = True
        Exit Function
    End If
    
    ' Caso especial: Y
    If letraUpper = "Y" Then
        ' Si hay contexto, determinar si Y actúa como vocal
        If Len(contexto) > 0 Then
            EsVocal = YEsVocal(contexto, InStr(1, UCase(contexto), "Y"))
        Else
            ' Sin contexto, asumir que es consonante
            EsVocal = False
        End If
        Exit Function
    End If
    
    EsVocal = False
End Function

' ============================================================================
' Función: YEsVocal
' Descripción: Determina si Y actúa como vocal en una palabra específica
' Parámetros:
'   - palabra: La palabra completa
'   - posicion: Posición de Y en la palabra (1-based)
' Retorna: True si Y es vocal en ese contexto
' ============================================================================
Private Function YEsVocal(ByVal palabra As String, ByVal posicion As Integer) As Boolean
    Dim palabraUpper As String
    Dim letraAnterior As String
    Dim letraSiguiente As String
    
    palabraUpper = UCase(Trim(palabra))
    
    ' Y al final de palabra suele ser vocal (ej: rey, ley)
    If posicion = Len(palabraUpper) Then
        YEsVocal = True
        Exit Function
    End If
    
    ' Y al inicio de palabra es consonante (ej: yate)
    If posicion = 1 Then
        YEsVocal = False
        Exit Function
    End If
    
    ' Obtener letras adyacentes
    If posicion > 1 Then letraAnterior = Mid(palabraUpper, posicion - 1, 1)
    If posicion < Len(palabraUpper) Then letraSiguiente = Mid(palabraUpper, posicion + 1, 1)
    
    ' Y entre consonantes suele ser vocal (ej: myth -> mit)
    If Not EsVocal(letraAnterior) And Not EsVocal(letraSiguiente) Then
        YEsVocal = True
        Exit Function
    End If
    
    ' Y seguida de vocal es consonante (ej: yate, maya)
    If EsVocal(letraSiguiente) Then
        YEsVocal = False
        Exit Function
    End If
    
    ' Por defecto, considerar vocal
    YEsVocal = True
End Function

' ============================================================================
' Función: EsLetraMuda
' Descripción: Verifica si una letra es muda en el contexto dado
' Parámetros:
'   - letra: La letra a verificar
'   - palabra: Palabra completa para contexto
'   - posicion: Posición de la letra en la palabra
' Retorna: True si la letra es muda
' ============================================================================
Public Function EsLetraMuda(ByVal letra As String, ByVal palabra As String, ByVal posicion As Integer) As Boolean
    Dim letraUpper As String
    Dim palabraUpper As String
    Dim letraAnterior As String
    
    letraUpper = UCase(Trim(Left(letra, 1)))
    palabraUpper = UCase(Trim(palabra))
    
    ' H siempre es muda en español
    If letraUpper = "H" Then
        EsLetraMuda = True
        Exit Function
    End If
    
    ' U muda después de Q o G (antes de E o I)
    If letraUpper = "U" And posicion > 1 Then
        letraAnterior = Mid(palabraUpper, posicion - 1, 1)
        
        If letraAnterior = "Q" Then
            EsLetraMuda = True
            Exit Function
        End If
        
        If letraAnterior = "G" And posicion < Len(palabraUpper) Then
            Dim letraSiguiente As String
            letraSiguiente = Mid(palabraUpper, posicion + 1, 1)
            If letraSiguiente = "E" Or letraSiguiente = "I" Then
                EsLetraMuda = True
                Exit Function
            End If
        End If
    End If
    
    EsLetraMuda = False
End Function

' ============================================================================
' Función: NormalizarTexto
' Descripción: Normaliza texto eliminando caracteres especiales y espacios extra
' Parámetros:
'   - texto: El texto a normalizar
' Retorna: Texto normalizado
' ============================================================================
Public Function NormalizarTexto(ByVal texto As String) As String
    Dim Resultado As String
    Dim i As Integer
    Dim caracter As String
    
    Resultado = ""
    texto = Trim(texto)
    
    For i = 1 To Len(texto)
        caracter = Mid(texto, i, 1)
        
        ' Mantener solo letras y espacios
        If (caracter >= "A" And caracter <= "Z") Or _
           (caracter >= "a" And caracter <= "z") Or _
           caracter = " " Or _
           InStr(1, "ÑÇÁÉÍÓÚÄËÏÖÜñçáéíóúäëïöü", caracter) > 0 Then
            Resultado = Resultado & caracter
        End If
    Next i
    
    ' Eliminar espacios múltiples
    Do While InStr(1, Resultado, "  ") > 0
        Resultado = Replace(Resultado, "  ", " ")
    Loop
    
    NormalizarTexto = UCase(Trim(Resultado))
End Function

' ============================================================================
' Función: ValidarFecha
' Descripción: Valida si una fecha es correcta
' Parámetros:
'   - fecha: La fecha a validar
' Retorna: True si la fecha es válida
' ============================================================================
Public Function ValidarFecha(ByVal fecha As Date) As Boolean
    On Error Resume Next
    
    ValidarFecha = False
    
    ' Verificar que la fecha sea válida y no futura
    If IsDate(fecha) And fecha <= Date Then
        ValidarFecha = True
    End If
End Function

' ============================================================================
' Función: ObtenerAnoUniversal
' Descripción: Calcula el año universal para un año dado
' Parámetros:
'   - ano: El año a calcular
' Retorna: Número del año universal (1-9 o maestro)
' ============================================================================
Public Function ObtenerAnoUniversal(ByVal ano As Integer) As Integer
    Dim suma As Long
    Dim i As Integer
    Dim digito As String
    
    suma = 0
    
    ' Sumar dígitos del año
    For i = 1 To Len(CStr(ano))
        digito = Mid(CStr(ano), i, 1)
        suma = suma + CLng(digito)
    Next i
    
    ' Reducir a dígito
    ObtenerAnoUniversal = ReducirADigito(suma, True)
End Function

' ============================================================================
' Función: FormatearNumeroResultado
' Descripción: Formatea un número de resultado para mostrar
' Parámetros:
'   - numero: El número a formatear
'   - mostrarTipo: Si True, incluye etiqueta del tipo de número
' Retorna: Cadena formateada
' ============================================================================
Public Function FormatearNumeroResultado(ByVal numero As Integer, Optional ByVal mostrarTipo As Boolean = True) As String
    Dim Resultado As String
    
    Resultado = CStr(numero)
    
    If mostrarTipo Then
        If EsNumeroMaestro(numero) Then
            Resultado = Resultado & " (Maestro)"
        ElseIf EsNumeroKarmico(numero) Then
            Resultado = Resultado & " (Kármico)"
        End If
    End If
    
    FormatearNumeroResultado = Resultado
End Function
