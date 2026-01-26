Attribute VB_Name = "modMotorFonetico"
' ------------------------------------------------------
' Nombre:    modMotorFonetico
' Tipo:      Módulo
' Propósito:
' Autor:     asalv
' Fecha:     15/01/2026
' ------------------------------------------------------

'Option Compare Database
'Option Explicit
'
'' ============================================================================
''  MÓDULO: modMotorFonetico
''  Descripción:
''     - Carga colecciones de nombres y apellidos UNA sola vez.
''     - Carga diccionarios SOLO cuando cambia el idioma.
''     - Convierte palabras usando:
''           1) Diccionario por prioridad
''           2) Tokenizador por idioma (fallback)
''     - Expone funciones públicas para el formulario.
'' ============================================================================

Option Compare Database
Option Explicit

' ============================================================================
'  MÓDULO: modMotorFonetico
'  Descripción:
'     - Preprocesa partículas (de, del, la, los, y, i, etc.)
'     - Fonetiza partículas por tokenizador según idioma
'     - Fonetiza nombres por diccionario + tokenizador
'     - Mantiene trazabilidad completa
' ============================================================================

' ============================================================================
'  COLECCIONES RICAS (cargadas una sola vez)
' ============================================================================

'Private ColNombres As Collection
'Private ColApellidos As Collection

'Public ColNombres As Collection
'Public ColApellidos As Collection

' ============================================================================
'  DICCIONARIOS (índice rápido Palabra ? ID)
' ============================================================================

'Private DicNom As Scripting.Dictionary
'Private DicApe As Scripting.Dictionary


Private UltimoIdiomaNom As String
Private UltimoIdiomaApe As String

' ============================================================================
'  FUNCIÓN PRINCIPAL
' ============================================================================

'Public Function MF_Convertir( _
'    ByVal Texto As String, _
'    ByVal Idioma As clsIdioma, _
'    Optional ByVal EsApellido As Boolean = False, _
'    Optional ByRef Trazabilidad As String _
'    ) As String
'
'    Dim palabras() As String
'    Dim parte As String
'    Dim fonema As String
'    Dim linea As String
'    Dim resultado As String
'    Dim i As Long
'
'    If Len(Trim$(Texto)) = 0 Then Exit Function
'
'    ' ============================================================
'    ' 1. Separar en palabras
'    ' ============================================================
'    palabras = Split(UCase$(Trim$(Texto)), " ")

Public Const mFonVersionMotor As String = "1.3"

Public Function MF_Convertir( _
    ByVal texto As String, _
    ByVal Idioma As clsIdioma, _
    Optional ByVal EsApellido As Boolean = False, _
    Optional ByRef Trazabilidad As String _
    ) As String

    Dim palabras() As String
    Dim Parte As String
    Dim fonema As String
    Dim linea As String
    Dim Resultado As String
    Dim i As Long

    If Len(Trim$(texto)) = 0 Then Exit Function

    ' ============================================================
    ' 0. Cargar estructuras internas del motor
    ' ============================================================
    Call MF_CargarColeccionesSiEsNecesario
    
    Call MF_CargarDiccionarioSiCambia(Idioma.Abreviado, EsApellido)

    ' ============================================================
    ' 1. Separar en palabras
    ' ============================================================
    palabras = Split(UCase$(Trim$(texto)), " ")

    ' ============================================================
    ' 2. Procesar cada parte
    ' ============================================================
    For i = LBound(palabras) To UBound(palabras)

        Parte = palabras(i)

        Select Case Parte

            ' ====================================================
            ' PARTICULAS ESPAÑOLAS Y CATALANAS
            ' ====================================================
            Case "DE", "LA", "LOS", "LAS", "DEL", "DA", "DO", "Y", "I"
                Select Case Idioma.Abreviado
                    Case "CA-IB", "CA-VA", "CA"
                        fonema = MF_TokenizarParticula(Parte, "CA")
                    Case Else
                        fonema = MF_TokenizarParticula(Parte, Idioma.Abreviado)
                End Select
                
                linea = Parte & " ? " & fonema & _
                        "   [Partícula / Tokenizador " & Idioma.Abreviado & "]"

            ' ====================================================
            ' NOMBRE NORMAL ? DICCIONARIO + TOKENIZADOR
            ' ====================================================
            Case Else
                fonema = MF_ProcesarNombre(Parte, Idioma, EsApellido)

                linea = Parte & " --> " & fonema & _
                        "   [Nombre / Diccionario + Tokenizador]"
                
                Debug.Print linea
        
        End Select
        
        fonema = fonema & " "
        ' ====================================================
        ' Añadir trazabilidad
        ' ====================================================
        If Len(Trazabilidad) > 0 Then
            Trazabilidad = Trazabilidad & vbCrLf & linea
        Else
            Trazabilidad = linea
        End If

        ' ====================================================
        ' Añadir al resultado final
        ' ====================================================
        Resultado = Resultado & fonema

    Next i


    Resultado = Trim$(Resultado)

    Do While InStr(Resultado, "  ") > 0
        Resultado = Replace(Resultado, "  ", " ")
    Loop

    Resultado = MF_NormalizarVocalesPorIdioma(Resultado, Idioma.Abreviado)
    
    MF_Convertir = Resultado

End Function


' ============================================================================
'  TOKENIZADOR DE PARTICULAS SEGÚN IDIOMA
' ============================================================================

Private Function MF_TokenizarParticula(Parte As String, Abreviado As String) As String

    Select Case Abreviado
        Case "es": MF_TokenizarParticula = ObtenerFonemasCastellano(Parte)
        Case "ca": MF_TokenizarParticula = ObtenerFonemasCatalan(Parte)
        Case "eu": MF_TokenizarParticula = ObtenerFonemasEuskera(Parte)
        Case "gl": MF_TokenizarParticula = ObtenerFonemasGalego(Parte)
        Case Else: MF_TokenizarParticula = ObtenerFonemasCastellano(Parte)
    End Select

End Function




' ============================================================================
'  PROCESAMIENTO DE NOMBRES (diccionario + tokenizador)
' ============================================================================

Private Function MF_ProcesarNombre(Parte As String, Idioma As clsIdioma, EsApellido As Boolean) As String
    Dim entrada As clsEntradaFonema

    ' 1) Buscar en diccionario
    Set entrada = MF_BuscarEntrada(Parte, EsApellido)

    If Not entrada Is Nothing Then
        MF_ProcesarNombre = entrada.fonema
        Exit Function
    End If

    ' 2) Fallback: tokenizar por idioma
    MF_ProcesarNombre = MF_TokenizarParticula(Parte, Idioma.Abreviado)

End Function


' ============================================================
'   NuevaConversionFonetica — Inicializa el objeto público Fonetica
' ============================================================
Public Sub NuevaConversionFonetica(Optional ByVal IDPersona As Long = 0)

    ' --- Identificación ---
'    Fonetica.IDFonetica = 0          ' Nuevo registro
    Fonetica.IDPersona = IDPersona   ' Persona a la que pertenece

    ' --- Sistema por defecto ---
    ' 1 = Tradicional clásico
    ' 2 = Fonético
    ' 3 = Tradicional moderno
    Fonetica.ModoFonetico = 1             ' Por defecto: fonético (puedes cambiarlo)

    ' --- Idiomas (vacíos hasta que el conversor los determine) ---
    Fonetica.IdiomaNombre = 0
    Fonetica.IdiomaApe1 = 0
    Fonetica.IdiomaApe2 = 0

    ' --- Resultados fonéticos (vacíos hasta conversión) ---
    Fonetica.FonNombre = ""
    Fonetica.FonApe1 = ""
    Fonetica.FonApe2 = ""

    ' --- Gestión ---
    Fonetica.FechaCalculo = Now
    Fonetica.Activo = True

End Sub



' ============================================================================
'  BÚSQUEDA EN DICCIONARIO (ya existente en tu motor)
' ============================================================================

Private Function MF_BuscarEntrada( _
    ByVal palabra As String, _
    ByVal EsApellido As Boolean _
    ) As clsEntradaFonema

    Dim clave As String
    Dim idx As Long

    clave = UCase$(palabra)

    If EsApellido Then
        If DicApe.Exists(clave) Then
            idx = DicApe(clave)
            Set MF_BuscarEntrada = ColApellidos(idx)
        End If
    Else
        If DicNom.Exists(clave) Then
            idx = DicNom(clave)
            Set MF_BuscarEntrada = ColNombres(idx)
        End If
    End If
End Function


' ============================================================================
'  CARGA DE COLECCIONES (UNA SOLA VEZ)
' ============================================================================

Private Sub MF_CargarColeccionesSiEsNecesario()
    If ColNombres Is Nothing Then
        Set ColNombres = MF_CargarColeccionDesdeTabla("tbmDicFonemasNom")
    End If

    If ColApellidos Is Nothing Then
        Set ColApellidos = MF_CargarColeccionDesdeTabla("tbmDicFonemasApe")
    End If
End Sub

Private Function MF_CargarColeccionDesdeTabla(ByVal Tabla As String) As Collection
    Dim col As New Collection
    Dim rs As DAO.Recordset
    Dim entrada As clsEntradaFonema

    Set rs = CurrentDb.OpenRecordset("SELECT * FROM " & Tabla & " WHERE Activo = True")

    Do While Not rs.EOF
        Set entrada = New clsEntradaFonema
        entrada.palabra = UCase$(rs!palabra)
        entrada.fonema = rs!FonemaCompleto
        entrada.Idioma = rs!Idioma
        entrada.Tipo = rs!TipoEntrada
        entrada.Notas = Nz(rs!Notas, "")
        'entrada.FonemaIPA = Nz(rs!FonemaIPA, "")
        entrada.Fuente = "Diccionario " & rs!Idioma

        col.Add entrada
        rs.MoveNext
    Loop

    rs.Close
    Set MF_CargarColeccionDesdeTabla = col
End Function

' ============================================================================
'  CARGA DE DICCIONARIOS (SOLO SI CAMBIA EL IDIOMA)
' ============================================================================

Private Sub MF_CargarDiccionarioSiCambia( _
    ByVal Idioma As String, _
    ByVal EsApellido As Boolean _
    )

    Dim dic As Scripting.Dictionary
    Dim col As Collection
    Dim ultimo As String

    Dim i As Long, idx As Long
    Dim entrada As clsEntradaFonema
    Dim clave As String
    
    If EsApellido Then
        Set dic = DicApe
        Set col = ColApellidos
        ultimo = UltimoIdiomaApe
    Else
        Set dic = DicNom
        Set col = ColNombres
        ultimo = UltimoIdiomaNom
    End If

    ' Si el idioma no ha cambiado ? no recargar
    If Idioma = ultimo And Not dic Is Nothing Then Exit Sub

    ' Crear diccionario nuevo
    Set dic = New Scripting.Dictionary

    ' Orden de prioridad
    Dim idiomas As Variant
    
    Select Case Idioma
        Case "CA-IB"
            idiomas = Array(Idioma, "CA-VA", "CA", "ES", "GA", "EU")
        Case "CA-VA"
            idiomas = Array(Idioma, "CA-IB", "CA", "ES", "GA", "EU")
        Case "CA"
            idiomas = Array(Idioma, "CA-IB", "CA-VA", "ES", "GA", "EU")
        Case Else
            idiomas = Array(Idioma, "ES", "GA", "CA-IB", "CA-VA", "CA", "EU")
    End Select
    
    

    ' Cargar según prioridad
    For i = LBound(idiomas) To UBound(idiomas)
        For idx = 1 To col.Count
            Set entrada = col(idx)

            If entrada.Idioma = idiomas(i) Then
                clave = entrada.palabra
'                If Left(clave, 4) = "SALV" Then
'                    Stop
'                End If
                If Not dic.Exists(clave) Then
                    dic.Add clave, idx
                End If
            End If
        Next idx
    Next i

    ' Guardar estado
    If EsApellido Then
        Set DicApe = dic
        UltimoIdiomaApe = Idioma
    Else
        Set DicNom = dic
        UltimoIdiomaNom = Idioma
    End If
End Sub

Public Function MF_NormalizarVocalesPorIdioma( _
    ByVal texto As String, _
    ByVal Idioma As String _
    ) As String

    Select Case UCase$(Idioma)

        Case "ES"
            texto = MF_NormalizarVocales_ES(texto)

        Case "CA", "CA-IB", "CA-VA"
            texto = MF_NormalizarVocales_CA(texto)

'        Case "CA-IB"
'            Texto = MF_NormalizarVocales_CA_IB(Texto)
'
'        Case "CA-VA"
'            Texto = MF_NormalizarVocales_CA_VA(Texto)

        Case "GL"
            texto = MF_NormalizarVocales_GL(texto)

        Case "EU"
            texto = MF_NormalizarVocales_EU(texto)

        Case Else
            texto = MF_NormalizarVocales_General(texto)

    End Select

    MF_NormalizarVocalesPorIdioma = texto

End Function

Private Function MF_NormalizarVocales_ES(ByVal texto As String) As String

    texto = Replace(texto, "Á", "A")
    texto = Replace(texto, "É", "E")
    texto = Replace(texto, "Í", "I")
    texto = Replace(texto, "Ó", "O")
    texto = Replace(texto, "Ú", "U")
    texto = Replace(texto, "Ü", "U")

    MF_NormalizarVocales_ES = texto

End Function

Private Function MF_NormalizarVocales_CA(ByVal texto As String) As String

    texto = Replace(texto, "À", "A")
    texto = Replace(texto, "Á", "A")

    texto = Replace(texto, "È", "E")
    texto = Replace(texto, "É", "E")

    texto = Replace(texto, "Í", "I")
    texto = Replace(texto, "Ï", "I")

    texto = Replace(texto, "Ò", "O")
    texto = Replace(texto, "Ó", "O")

    texto = Replace(texto, "Ú", "U")
    texto = Replace(texto, "Ü", "U")

    MF_NormalizarVocales_CA = texto

End Function

Private Function MF_NormalizarVocales_CA_IB(ByVal texto As String) As String
    MF_NormalizarVocales_CA_IB = MF_NormalizarVocales_CA(texto)
End Function

Private Function MF_NormalizarVocales_CA_VA(ByVal texto As String) As String
    MF_NormalizarVocales_CA_VA = MF_NormalizarVocales_CA(texto)
End Function

Private Function MF_NormalizarVocales_GL(ByVal texto As String) As String

    texto = Replace(texto, "Á", "A")
    texto = Replace(texto, "É", "E")
    texto = Replace(texto, "Í", "I")
    texto = Replace(texto, "Ó", "O")
    texto = Replace(texto, "Ú", "U")

    MF_NormalizarVocales_GL = texto

End Function

Private Function MF_NormalizarVocales_EU(ByVal texto As String) As String

    texto = Replace(texto, "Á", "A")
    texto = Replace(texto, "É", "E")
    texto = Replace(texto, "Í", "I")
    texto = Replace(texto, "Ó", "O")
    texto = Replace(texto, "Ú", "U")

    MF_NormalizarVocales_EU = texto

End Function

Private Function MF_NormalizarVocales_General(ByVal texto As String) As String
    
    Dim i As Integer
    Dim C As String
    Dim Res As String

    C = ""
    Res = ""
    
    For i = 1 To Len(texto)
        C = Mid(texto, i, 1)
        Select Case C
            
            Case "Á", "À", "Ä"
                C = "A"
    
            Case "É", "È", "Ë"
                C = "E"
    
            Case "Í", "Ï"
                C = "I"
            
            Case "Ó", "Ò", "Ö"
                C = "O"
    
            Case "Ú", "Ü"
                C = "U"
        End Select
        
        Res = Res & C
    Next i
    
    MF_NormalizarVocales_General = Res
    
''---------------------------------
'    Texto = Replace(Texto, "Á", "A")
'    Texto = Replace(Texto, "À", "A")
'    Texto = Replace(Texto, "Ä", "A")
'
'    Texto = Replace(Texto, "É", "E")
'    Texto = Replace(Texto, "È", "E")
'    Texto = Replace(Texto, "Ë", "E")
'
'    Texto = Replace(Texto, "Í", "I")
'    Texto = Replace(Texto, "Ï", "I")
'
'    Texto = Replace(Texto, "Ó", "O")
'    Texto = Replace(Texto, "Ò", "O")
'    Texto = Replace(Texto, "Ö", "O")
'
'    Texto = Replace(Texto, "Ú", "U")
'    Texto = Replace(Texto, "Ü", "U")
'
'    MF_NormalizarVocales_General = Texto

End Function


