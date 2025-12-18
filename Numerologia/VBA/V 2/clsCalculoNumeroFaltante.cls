VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCalculoNumeroFaltante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ============================================================================
' Clase: clsCalculoNumeroFaltante
' Descripción: Calcula los Números Faltantes (ausentes en el nombre)
' Autor: Sistema de Numerología
' Fecha: 2024
' ============================================================================

Private objBase As clsCalculoBase
Private pNumerosFaltantes() As Integer
Private pCantidadFaltantes As Integer

' ============================================================================
' PROPIEDADES PÚBLICAS
' ============================================================================

Public Property Get NombreCompleto() As String
    NombreCompleto = objBase.NombreCompleto
End Property

Public Property Let NombreCompleto(ByVal valor As String)
    objBase.NombreCompleto = valor
End Property

Public Property Get CantidadFaltantes() As Integer
    CantidadFaltantes = pCantidadFaltantes
End Property

' ============================================================================
' MÉTODOS PÚBLICOS
' ============================================================================

Public Function Calcular() As Integer
    Dim i As Integer
    Dim letra As String
    Dim valor As Integer
    Dim nombreNormalizado As String
    Dim numerosPresentes() As Boolean
    Dim indice As Integer
    
    On Error GoTo ErrorHandler
    
    ' Validar nombre
    If Len(objBase.NombreCompleto) = 0 Then
        err.Raise vbObjectError + 2033, "clsCalculoNumeroFaltante", ERR_NOMBRE_VACIO
    End If
    
    ' Inicializar array de presencia
    ReDim numerosPresentes(1 To 9)
    For i = 1 To 9
        numerosPresentes(i) = False
    Next i
    
    nombreNormalizado = objBase.NombreCompleto
    
    ' Marcar números presentes
    For i = 1 To Len(nombreNormalizado)
        letra = Mid(nombreNormalizado, i, 1)
        
        ' Saltar espacios
        If letra = " " Then GoTo SiguienteLetra
        
        ' Saltar letras mudas
        If EsLetraMuda(letra, nombreNormalizado, i) Then GoTo SiguienteLetra
        
        ' Obtener valor
        valor = ObtenerValorLetra(letra)
        
        ' Marcar como presente
        If valor >= 1 And valor <= 9 Then
            numerosPresentes(valor) = True
        End If
        
SiguienteLetra:
    Next i
    
    ' Contar y almacenar números faltantes
    pCantidadFaltantes = 0
    For i = 1 To 9
        If Not numerosPresentes(i) Then
            pCantidadFaltantes = pCantidadFaltantes + 1
        End If
    Next i
    
    ' Dimensionar array de faltantes
    If pCantidadFaltantes > 0 Then
        ReDim pNumerosFaltantes(1 To pCantidadFaltantes)
        indice = 1
        
        For i = 1 To 9
            If Not numerosPresentes(i) Then
                pNumerosFaltantes(indice) = i
                indice = indice + 1
            End If
        Next i
    Else
        ReDim pNumerosFaltantes(0)
    End If
    
    Calcular = pCantidadFaltantes
    Exit Function
    
ErrorHandler:
    err.Raise err.Number, "clsCalculoNumeroFaltante.Calcular", err.Description
End Function

Public Function ObtenerNumerosFaltantes() As String
    Dim i As Integer
    Dim Resultado As String
    
    If pCantidadFaltantes = 0 Then
        ObtenerNumerosFaltantes = "Ninguno (todos los números están presentes)"
        Exit Function
    End If
    
    Resultado = ""
    For i = 1 To pCantidadFaltantes
        If i > 1 Then Resultado = Resultado & ", "
        Resultado = Resultado & CStr(pNumerosFaltantes(i))
    Next i
    
    ObtenerNumerosFaltantes = Resultado
End Function

Public Function ObtenerInterpretacion(Optional ByVal numeroFaltante As Integer = 0) As String
    Dim nombreArchivo As String
    Dim interpretacionCompleta As String
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    ' Verificar que se haya calculado
    If pCantidadFaltantes < 0 Then
        Calcular
    End If
    
    ' Si no hay números faltantes
    If pCantidadFaltantes = 0 Then
        ObtenerInterpretacion = "No hay números faltantes. Tu nombre contiene todos los números del 1 al 9, " & _
                                "lo que indica una personalidad completa y versátil con acceso a todas las energías numerológicas."
        Exit Function
    End If
    
    ' Si se especifica un número concreto
    If numeroFaltante > 0 And numeroFaltante <= 9 Then
        nombreArchivo = objBase.ObtenerNombreArchivoInterpretacion("Ausente", numeroFaltante)
        ObtenerInterpretacion = objBase.CargarInterpretacionDesdeArchivo(nombreArchivo)
        Exit Function
    End If
    
    ' Si no se especifica, cargar interpretaciones de todos los faltantes
    interpretacionCompleta = "NÚMEROS FALTANTES EN TU NOMBRE" & vbCrLf & vbCrLf
    interpretacionCompleta = interpretacionCompleta & "Números ausentes: " & ObtenerNumerosFaltantes() & vbCrLf & vbCrLf
    
    For i = 1 To pCantidadFaltantes
        nombreArchivo = objBase.ObtenerNombreArchivoInterpretacion("Ausente", pNumerosFaltantes(i))
        interpretacionCompleta = interpretacionCompleta & "=== NÚMERO " & pNumerosFaltantes(i) & " AUSENTE ===" & vbCrLf
        interpretacionCompleta = interpretacionCompleta & objBase.CargarInterpretacionDesdeArchivo(nombreArchivo) & vbCrLf & vbCrLf
    Next i
    
    ObtenerInterpretacion = interpretacionCompleta
    Exit Function
    
ErrorHandler:
    err.Raise err.Number, "clsCalculoNumeroFaltante.ObtenerInterpretacion", err.Description
End Function

Public Function NumeroEstaFaltante(ByVal numero As Integer) As Boolean
    Dim i As Integer
    
    If numero < 1 Or numero > 9 Then
        NumeroEstaFaltante = False
        Exit Function
    End If
    
    For i = 1 To pCantidadFaltantes
        If pNumerosFaltantes(i) = numero Then
            NumeroEstaFaltante = True
            Exit Function
        End If
    Next i
    
    NumeroEstaFaltante = False
End Function

' ============================================================================
' EVENTOS DE CLASE
' ============================================================================

Private Sub Class_Initialize()
    Set objBase = New clsCalculoBase
    pCantidadFaltantes = -1  ' Indica que no se ha calculado
End Sub

Private Sub Class_Terminate()
    Set objBase = Nothing
    If pCantidadFaltantes >= 0 Then
        Erase pNumerosFaltantes
    End If
End Sub
