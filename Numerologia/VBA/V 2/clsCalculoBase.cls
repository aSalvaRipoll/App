VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCalculoBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ============================================================================
' Proyecto:     Sistema de Numerología Tradicional y Fonético
' Clase: clsCalculoBase
' Descripción: Clase base abstracta para todos los cálculos numerológicos (Friend)
' Autor: Sistema de Numerología
' Fecha: 2024
' Nota: Esta clase NO debe instanciarse directamente
' ============================================================================

' ============================================================================
' CLASE BASE PARA CÁLCULOS NUMEROLÓGICOS
' ============================================================================

' Propiedades privadas comunes
Private pFechaNacimiento As Date
Private pNombreCompleto As String
Private pResultado As Integer
Private pInterpretacion As String
Private pEsKarmico As Boolean
Private pEsMaestro As Boolean


' ============================================================================
' PROPIEDADES PÚBLICAS
' ============================================================================

Public Property Get FechaNacimiento() As Date
    FechaNacimiento = pFechaNacimiento
End Property

Public Property Let FechaNacimiento(ByVal valor As Date)
    If Not ValidarFecha(valor) Then
        err.Raise vbObjectError + 1001, "clsCalculoBase", ERR_FECHA_INVALIDA
    End If
    pFechaNacimiento = valor
End Property

Public Property Get NombreCompleto() As String
    NombreCompleto = pNombreCompleto
End Property

Public Property Let NombreCompleto(ByVal valor As String)
    If Len(Trim(valor)) = 0 Then
        err.Raise vbObjectError + 1002, "clsCalculoBase", ERR_NOMBRE_VACIO
    End If
    pNombreCompleto = NormalizarTexto(valor)
End Property

Public Property Get Resultado() As Integer
    Resultado = pResultado
End Property

Public Property Let Resultado(ByVal valor As Integer)
    pResultado = valor
    pEsMaestro = EsNumeroMaestro(valor)
    pEsKarmico = EsNumeroKarmico(valor)
End Property

Public Property Get Interpretacion() As String
    Interpretacion = pInterpretacion
End Property

Public Property Let Interpretacion(ByVal valor As String)
    pInterpretacion = valor
End Property

Public Property Get EsKarmico() As Boolean
    EsKarmico = pEsKarmico
End Property

Public Property Get EsMaestro() As Boolean
    EsMaestro = pEsMaestro
End Property

' ============================================================================
' MÉTODOS PÚBLICOS ABSTRACTOS (deben ser implementados por clases derivadas)
' ============================================================================

Public Function Calcular() As Integer
    ' Método abstracto - debe ser implementado en clases derivadas
    err.Raise vbObjectError + 1003, "clsCalculoBase", "Método Calcular debe ser implementado en clase derivada"
End Function

Public Function ObtenerInterpretacion() As String
    ' Método abstracto - debe ser implementado en clases derivadas
    err.Raise vbObjectError + 1004, "clsCalculoBase", "Método ObtenerInterpretacion debe ser implementado en clase derivada"
End Function

' ============================================================================
' MÉTODOS PROTEGIDOS (para uso de clases derivadas)
' ============================================================================

Friend Function CargarInterpretacionDesdeArchivo(ByVal nombreArchivo As String) As String
    Dim rutaCompleta As String
    Dim contenido As String
    Dim numArchivo As Integer
    
    On Error GoTo ErrorHandler
    
    ' Construir ruta completa
    rutaCompleta = CurrentProject.Path & "\" & RUTA_INTERPRETACIONES & nombreArchivo & EXTENSION_MARKDOWN
    
    ' Verificar que el archivo existe
    If Dir(rutaCompleta) = "" Then
        err.Raise vbObjectError + 1005, "clsCalculoBase", ERR_ARCHIVO_NO_ENCONTRADO & ": " & rutaCompleta
    End If
    
    ' Leer contenido del archivo
    numArchivo = FreeFile
    Open rutaCompleta For Input As #numArchivo
    contenido = Input$(LOF(numArchivo), numArchivo)
    Close #numArchivo
    
    CargarInterpretacionDesdeArchivo = contenido
    Exit Function
    
ErrorHandler:
    If numArchivo > 0 Then Close #numArchivo
    err.Raise err.Number, "clsCalculoBase.CargarInterpretacionDesdeArchivo", err.Description
End Function

Friend Function ObtenerNombreArchivoInterpretacion(ByVal tipoCalculo As String, ByVal numero As Integer) As String
    ' Formato: TipoCalculo_Numero.md
    ' Ejemplo: CaminoVida_1.md, Destino_11.md
    ObtenerNombreArchivoInterpretacion = tipoCalculo & "_" & CStr(numero)
End Function

' ============================================================================
' MÉTODOS DE UTILIDAD PROTEGIDOS
' ============================================================================

Friend Function SumarDigitosFecha(ByVal fecha As Date, Optional ByVal permitirMaestros As Boolean = True) As Integer
    Dim dia As Integer
    Dim mes As Integer
    Dim ano As Integer
    Dim suma As Long
    
    dia = Day(fecha)
    mes = Month(fecha)
    ano = Year(fecha)
    
    ' Reducir cada componente
    dia = ReducirADigito(dia, permitirMaestros)
    mes = ReducirADigito(mes, permitirMaestros)
    ano = ReducirADigito(ano, permitirMaestros)
    
    ' Sumar y reducir total
    suma = CLng(dia) + CLng(mes) + CLng(ano)
    SumarDigitosFecha = ReducirADigito(suma, permitirMaestros)
End Function

Friend Function SumarValoresNombre(ByVal nombre As String, Optional ByVal soloVocales As Boolean = False, Optional ByVal soloConsonantes As Boolean = False) As Integer
    Dim suma As Long
    Dim i As Integer
    Dim letra As String
    Dim valor As Integer
    Dim nombreNormalizado As String
    
    nombreNormalizado = NormalizarTexto(nombre)
    suma = 0
    
    For i = 1 To Len(nombreNormalizado)
        letra = Mid(nombreNormalizado, i, 1)
        
        ' Saltar espacios
        If letra = " " Then GoTo SiguienteLetra
        
        ' Saltar letras mudas
        If EsLetraMuda(letra, nombreNormalizado, i) Then GoTo SiguienteLetra
        
        ' Aplicar filtros según parámetros
        If soloVocales And Not EsVocal(letra, nombreNormalizado) Then GoTo SiguienteLetra
        If soloConsonantes And EsVocal(letra, nombreNormalizado) Then GoTo SiguienteLetra
        
        ' Obtener valor y sumar
        valor = ObtenerValorLetra(letra)
        If valor > 0 Then
            suma = suma + valor
        End If
        
SiguienteLetra:
    Next i
    
    SumarValoresNombre = ReducirADigito(suma, True)
End Function

' ============================================================================
' MÉTODO DE DEBUGGING
' ============================================================================

Public Function ObtenerInformacionCompleta() As String
    Dim info As String
    
    info = "=== INFORMACIÓN DEL CÁLCULO ===" & vbCrLf
    info = info & "Tipo: " & TypeName(Me) & vbCrLf
    
    If pFechaNacimiento > 0 Then
        info = info & "Fecha de Nacimiento: " & Format(pFechaNacimiento, "dd/mm/yyyy") & vbCrLf
    End If
    
    If Len(pNombreCompleto) > 0 Then
        info = info & "Nombre Completo: " & pNombreCompleto & vbCrLf
    End If
    
    If pResultado > 0 Then
        info = info & "Resultado: " & FormatearNumeroResultado(pResultado, True) & vbCrLf
        info = info & "Es Maestro: " & IIf(pEsMaestro, "Sí", "No") & vbCrLf
        info = info & "Es Kármico: " & IIf(pEsKarmico, "Sí", "No") & vbCrLf
    End If
    
    ObtenerInformacionCompleta = info
End Function

' ============================================================================
' EVENTOS DE CLASE
' ============================================================================

Private Sub Class_Initialize()
    ' Inicializar propiedades
    pFechaNacimiento = 0
    pNombreCompleto = ""
    pResultado = 0
    pInterpretacion = ""
    pEsKarmico = False
    pEsMaestro = False
End Sub

Private Sub Class_Terminate()
    ' Limpiar recursos si es necesario
End Sub
