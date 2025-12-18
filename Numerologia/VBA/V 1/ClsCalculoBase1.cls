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

Friend  Property Get FechaNacimiento() As Date
    FechaNacimiento = pFechaNacimiento
End Property

Friend  Property Let FechaNacimiento(ByVal valor As Date)
    If Not ValidarFecha(valor) Then
        Err.Raise vbObjectError + 1001, "clsCalculoBase", ERR_FECHA_INVALIDA
    End If
    pFechaNacimiento = valor
End Property

Friend  Property Get NombreCompleto() As String
    NombreCompleto = pNombreCompleto
End Property

Friend  Property Let NombreCompleto(ByVal valor As String)
    If Len(Trim(valor)) = 0 Then
        Err.Raise vbObjectError + 1002, "clsCalculoBase", ERR_NOMBRE_VACIO
    End If
    pNombreCompleto = NormalizarTexto(valor)
End Property

Friend  Property Get Resultado() As Integer
    Resultado = pResultado
End Property

Friend Property Let Resultado(ByVal valor As Integer)
    pResultado = valor
    pEsMaestro = EsNumeroMaestro(valor)
    pEsKarmico = EsNumeroKarmico(valor)
End Property

Friend  Property Get Interpretacion() As String
    Interpretacion = pInterpretacion
End Property

Friend Property Let Interpretacion(ByVal valor As String)
    pInterpretacion = valor
End Property

Friend  Property Get EsKarmico() As Boolean
    EsKarmico = pEsKarmico
End Property

Friend  Property Get EsMaestro() As Boolean
    EsMaestro = pEsMaestro
End Property

' ============================================================================
' MÉTODOS PÚBLICOS ABSTRACTOS
' ============================================================================

Friend  Function Calcular() As Integer
    Err.Raise vbObjectError + 1003, "clsCalculoBase", _
              "Método Calcular debe ser implementado en clase derivada"
End Function

Friend  Function ObtenerInterpretacion() As String
    Err.Raise vbObjectError + 1004, "clsCalculoBase", _
              "Método ObtenerInterpretacion debe ser implementado en clase derivada"
End Function

' ============================================================================
' MÉTODOS PROTEGIDOS (Friend) PARA CLASES DERIVADAS
' ============================================================================

Friend Function CargarInterpretacionDesdeArchivo(ByVal nombreArchivo As String) As String
    Dim rutaCompleta As String
    Dim contenido As String
    Dim numArchivo As Integer

    On Error GoTo ErrorHandler

    rutaCompleta = CurrentProject.Path & "\" & RUTA_INTERPRETACIONES & _
                   nombreArchivo & EXTENSION_MARKDOWN

    If Dir(rutaCompleta) = "" Then
        Err.Raise vbObjectError + 1005, "clsCalculoBase", _
                  ERR_ARCHIVO_NO_ENCONTRADO & ": " & rutaCompleta
    End If

    numArchivo = FreeFile
    Open rutaCompleta For Input As #numArchivo
    contenido = Input$(LOF(numArchivo), numArchivo)
    Close #numArchivo

    CargarInterpretacionDesdeArchivo = contenido
    Exit Function

ErrorHandler:
    If numArchivo > 0 Then Close #numArchivo
    Err.Raise Err.Number, "clsCalculoBase.CargarInterpretacionDesdeArchivo", Err.Description
End Function

Friend Function ObtenerNombreArchivoInterpretacion(ByVal tipoCalculo As String, _
                                                   ByVal numero As Integer) As String
    ObtenerNombreArchivoInterpretacion = tipoCalculo & "_" & CStr(numero)
End Function

' ============================================================================
' MÉTODOS DE UTILIDAD (Friend)
' ============================================================================

Friend Function SumarDigitosFecha(ByVal fecha As Date, _
                                  Optional ByVal permitirMaestros As Boolean = True) As Integer
    Dim dia As Integer, mes As Integer, ano As Integer
    Dim suma As Long

    dia = ReducirADigito(Day(fecha), permitirMaestros)
    mes = ReducirADigito(Month(fecha), permitirMaestros)
    ano = ReducirADigito(Year(fecha), permitirMaestros)

    suma = CLng(dia) + CLng(mes) + CLng(ano)
    SumarDigitosFecha = ReducirADigito(suma, permitirMaestros)
End Function

Friend Function SumarValoresNombre(ByVal nombre As String, _
                                   Optional ByVal soloVocales As Boolean = False, _
                                   Optional ByVal soloConsonantes As Boolean = False) As Integer
    Dim suma As Long
    Dim i As Integer
    Dim letra As String
    Dim valor As Integer
    Dim nombreNormalizado As String

    nombreNormalizado = NormalizarTexto(nombre)
    suma = 0

    For i = 1 To Len(nombreNormalizado)
        letra = Mid(nombreNormalizado, i, 1)

        If letra = " " Then GoTo SiguienteLetra
        If EsLetraMuda(letra, nombreNormalizado, i) Then GoTo SiguienteLetra

        If soloVocales And Not EsVocal(letra, nombreNormalizado) Then GoTo SiguienteLetra
        If soloConsonantes And EsVocal(letra, nombreNormalizado) Then GoTo SiguienteLetra

        valor = ObtenerValorLetra(letra)
        If valor > 0 Then suma = suma + valor

SiguienteLetra:
    Next i

    SumarValoresNombre = ReducirADigito(suma, True)
End Function

' ============================================================================
' DEBUG
' ============================================================================

Friend  Function ObtenerInformacionCompleta() As String
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
    pFechaNacimiento = 0
    pNombreCompleto = ""
    pResultado = 0
    pInterpretacion = ""
    pEsKarmico = False
    pEsMaestro = False
End Sub

Private Sub Class_Terminate()
    ' Limpieza si fuera necesaria
End Sub
