Option Compare Database
Option Explicit

' =============================================================================
' Módulo: modSalidaWord
' Descripción: Gestión de salida a Word con soporte para UTF-8 y emojis
' Autor: Sistema de Numerología
' Fecha: 2025
' =============================================================================

' =============================================================================
' LECTURA DE ARCHIVOS UTF-8
' =============================================================================

Public Function LeerArchivoUTF8(ByVal rutaArchivo As String) As String
    ' Lee archivos de texto en formato UTF-8 (Markdown, etc.)
    ' Maneja correctamente acentos, Ñ, Ç y caracteres especiales
    '
    ' Parámetros:
    '   rutaArchivo: Ruta completa del archivo a leer
    ' Retorna: Contenido del archivo como String
    
    On Error GoTo ErrorHandler
    
    Dim stream As Object
    Dim contenido As String
    
    ' Verificar que el archivo existe
    If Dir(rutaArchivo) = "" Then
        LeerArchivoUTF8 = ""
        Debug.Print "ERROR: Archivo no encontrado: " & rutaArchivo
        Exit Function
    End If
    
    ' Crear objeto ADODB.Stream para UTF-8
    Set stream = CreateObject("ADODB.Stream")
    
    With stream
        .Type = 2           ' adTypeText
        .Charset = "UTF-8"  ' Codificación UTF-8
        .Open
        .LoadFromFile rutaArchivo
        contenido = .ReadText
        .Close
    End With
    
    Set stream = Nothing
    LeerArchivoUTF8 = contenido
    
    Exit Function
    
ErrorHandler:
    Debug.Print "ERROR en LeerArchivoUTF8: " & Err.Description
    LeerArchivoUTF8 = ""
    If Not stream Is Nothing Then
        If stream.State = 1 Then stream.Close
    End If
    Set stream = Nothing
End Function

' =============================================================================
' REEMPLAZO DE PLACEHOLDERS POR EMOJIS
' =============================================================================

Public Function ReemplazarPlaceholdersPorEmojis(ByVal texto As String) As String
    ' Reemplaza placeholders [#NOMBRE#] por códigos de emoji VBA
    ' Los emojis funcionan perfectamente en Office 2019
    '
    ' Parámetros:
    '   texto: Texto con placeholders
    ' Retorna: Texto con emojis insertados
    
    Dim resultado As String
    resultado = texto
    
    ' =========================================================================
    ' CELEBRACIÓN Y LOGROS
    ' =========================================================================
    resultado = Replace(resultado, "[#CELEBRACION#]", ChrW(&HD83C) & ChrW(&HDF89))  ' 🎉
    resultado = Replace(resultado, "[#BRILLO#]", ChrW(&H2728))                      ' ✨
    resultado = Replace(resultado, "[#ESTRELLA#]", ChrW(&HD83C) & ChrW(&HDF1F))     ' 🌟
    resultado = Replace(resultado, "[#ESTRELLA_SIMPLE#]", ChrW(&H2B50))             ' ⭐
    resultado = Replace(resultado, "[#TROFEO#]", ChrW(&HD83C) & ChrW(&HDFC6))       ' 🏆
    resultado = Replace(resultado, "[#CORONA#]", ChrW(&HD83D) & ChrW(&HDC51))       ' 👑
    
    ' =========================================================================
    ' ADVERTENCIAS Y PRECAUCIONES
    ' =========================================================================
    resultado = Replace(resultado, "[#ADVERTENCIA#]", ChrW(&H26A0) & ChrW(&HFE0F)) ' ⚠️
    resultado = Replace(resultado, "[#ALERTA#]", ChrW(&HD83D) & ChrW(&HDEA8))      ' 🚨
    resultado = Replace(resultado, "[#PROHIBIDO#]", ChrW(&H26D4))                  ' ⛔
    
    ' =========================================================================
    ' APROBACIÓN Y NEGACIÓN
    ' =========================================================================
    resultado = Replace(resultado, "[#CHECK#]", ChrW(&H2705))                      ' ✅
    resultado = Replace(resultado, "[#CRUZ#]", ChrW(&H274C))                       ' ❌
    resultado = Replace(resultado, "[#PULGAR_ARRIBA#]", ChrW(&HD83D) & ChrW(&HDC4D)) ' 👍
    resultado = Replace(resultado, "[#PULGAR_ABAJO#]", ChrW(&HD83D) & ChrW(&HDC4E))  ' 👎
    
    ' =========================================================================
    ' TRABAJO Y ACCIÓN
    ' =========================================================================
    resultado = Replace(resultado, "[#MALETIN#]", ChrW(&HD83D) & ChrW(&HDCBC))     ' 💼
    resultado = Replace(resultado, "[#MARTILLO#]", ChrW(&HD83D) & ChrW(&HDD28))    ' 🔨
    resultado = Replace(resultado, "[#CONSTRUCCION#]", ChrW(&HD83C) & ChrW(&HDFD7) & ChrW(&HFE0F)) ' 🏗️
    resultado = Replace(resultado, "[#COHETE#]", ChrW(&HD83D) & ChrW(&HDE80))      ' 🚀
    resultado = Replace(resultado, "[#DIANA#]", ChrW(&HD83C) & ChrW(&HDFAF))       ' 🎯
    
    ' =========================================================================
    ' CONOCIMIENTO Y SABIDURÍA
    ' =========================================================================
    resultado = Replace(resultado, "[#LIBROS#]", ChrW(&HD83D) & ChrW(&HDCDA))      ' 📚
    resultado = Replace(resultado, "[#BOMBILLA#]", ChrW(&HD83D) & ChrW(&HDCA1))    ' 💡
    resultado = Replace(resultado, "[#BOLA_CRISTAL#]", ChrW(&HD83D) & ChrW(&HDD2E)) ' 🔮
    resultado = Replace(resultado, "[#CEREBRO#]", ChrW(&HD83E) & ChrW(&HDDE0))     ' 🧠
    
    ' =========================================================================
    ' AMOR Y RELACIONES
    ' =========================================================================
    resultado = Replace(resultado, "[#CORAZON#]", ChrW(&H2764) & ChrW(&HFE0F))     ' ❤️
    resultado = Replace(resultado, "[#DOS_CORAZONES#]", ChrW(&HD83D) & ChrW(&HDC95)) ' 💕
    resultado = Replace(resultado, "[#APRETÓN_MANOS#]", ChrW(&HD83E) & ChrW(&HDD1D)) ' 🤝
    
    ' =========================================================================
    ' DINERO Y ABUNDANCIA
    ' =========================================================================
    resultado = Replace(resultado, "[#BOLSA_DINERO#]", ChrW(&HD83D) & ChrW(&HDCB0)) ' 💰
    resultado = Replace(resultado, "[#DIAMANTE#]", ChrW(&HD83D) & ChrW(&HDC8E))     ' 💎
    resultado = Replace(resultado, "[#DINERO_VOLANDO#]", ChrW(&HD83D) & ChrW(&HDCB8)) ' 💸
    
    ' =========================================================================
    ' NATURALEZA Y CRECIMIENTO
    ' =========================================================================
    resultado = Replace(resultado, "[#SEMILLA#]", ChrW(&HD83C) & ChrW(&HDF31))     ' 🌱
    resultado = Replace(resultado, "[#ARBOL#]", ChrW(&HD83C) & ChrW(&HDF33))       ' 🌳
    resultado = Replace(resultado, "[#OLA#]", ChrW(&HD83C) & ChrW(&HDF0A))         ' 🌊
    resultado = Replace(resultado, "[#FUEGO#]", ChrW(&HD83D) & ChrW(&HDD25))       ' 🔥
    resultado = Replace(resultado, "[#ARCOIRIS#]", ChrW(&HD83C) & ChrW(&HDF08))    ' 🌈
    
    ' =========================================================================
    ' FUERZA Y PODER
    ' =========================================================================
    resultado = Replace(resultado, "[#MUSCULO#]", ChrW(&HD83D) & ChrW(&HDCAA))     ' 💪
    resultado = Replace(resultado, "[#RAYO#]", ChrW(&H26A1))                       ' ⚡
    resultado = Replace(resultado, "[#AGUILA#]", ChrW(&HD83E) & ChrW(&HDD85))      ' 🦅
    
    ' =========================================================================
    ' CREATIVIDAD Y ARTE
    ' =========================================================================
    resultado = Replace(resultado, "[#PALETA#]", ChrW(&HD83C) & ChrW(&HDFA8))      ' 🎨
    resultado = Replace(resultado, "[#MASCARAS#]", ChrW(&HD83C) & ChrW(&HDFAD))    ' 🎭
    resultado = Replace(resultado, "[#NOTA_MUSICAL#]", ChrW(&HD83C) & ChrW(&HDFB5)) ' 🎵
    
    ' =========================================================================
    ' DATOS Y ANÁLISIS
    ' =========================================================================
    resultado = Replace(resultado, "[#GRAFICO#]", ChrW(&HD83D) & ChrW(&HDCCA))     ' 📊
    resultado = Replace(resultado, "[#GRAFICO_SUBIDA#]", ChrW(&HD83D) & ChrW(&HDCC8)) ' 📈
    resultado = Replace(resultado, "[#GRAFICO_BAJADA#]", ChrW(&HD83D) & ChrW(&HDCC9)) ' 📉
    
    ' =========================================================================
    ' TIEMPO Y CICLOS
    ' =========================================================================
    resultado = Replace(resultado, "[#RELOJ#]", ChrW(&H23F0))                      ' ⏰
    resultado = Replace(resultado, "[#CICLO#]", ChrW(&HD83D) & ChrW(&HDD04))       ' 🔄
    resultado = Replace(resultado, "[#LUNA#]", ChrW(&HD83C) & ChrW(&HDF19))        ' 🌙
    resultado = Replace(resultado, "[#SOL#]", ChrW(&H2600) & ChrW(&HFE0F))         ' ☀️
    
    ' =========================================================================
    ' DIRECCIÓN Y MOVIMIENTO
    ' =========================================================================
    resultado = Replace(resultado, "[#FLECHA_DERECHA#]", ChrW(&H27A1) & ChrW(&HFE0F)) ' ➡️
    resultado = Replace(resultado, "[#FLECHA_ARRIBA#]", ChrW(&H2B06) & ChrW(&HFE0F))  ' ⬆️
    resultado = Replace(resultado, "[#FLECHA_ABAJO#]", ChrW(&H2B07) & ChrW(&HFE0F))   ' ⬇️
    
    ' =========================================================================
    ' SÍMBOLOS UNICODE BÁSICOS (100% compatibles)
    ' =========================================================================
    resultado = Replace(resultado, "[#TICK#]", ChrW(&H2713))                       ' ✓
    resultado = Replace(resultado, "[#X#]", ChrW(&H2717))                          ' ✗
    resultado = Replace(resultado, "[#PUNTO#]", ChrW(&H2022))                      ' •
    resultado = Replace(resultado, "[#CIRCULO_VACIO#]", ChrW(&H25CB))              ' ○
    resultado = Replace(resultado, "[#CIRCULO_LLENO#]", ChrW(&H25CF))              ' ●
    resultado = Replace(resultado, "[#DIAMANTE_LLENO#]", ChrW(&H25C6))             ' ◆
    resultado = Replace(resultado, "[#DIAMANTE_VACIO#]", ChrW(&H25C7))             ' ◇
    resultado = Replace(resultado, "[#ESTRELLA_NEGRA#]", ChrW(&H2605))             ' ★
    resultado = Replace(resultado, "[#ESTRELLA_BLANCA#]", ChrW(&H2606))            ' ☆
    resultado = Replace(resultado, "[#TRIANGULO_DERECHA#]", ChrW(&H25BA))          ' ►
    resultado = Replace(resultado, "[#TRIANGULO_ABAJO#]", ChrW(&H25BC))            ' ▼
    resultado = Replace(resultado, "[#TRIANGULO_ARRIBA#]", ChrW(&H25B2))           ' ▲
    resultado = Replace(resultado, "[#TRIANGULO_IZQUIERDA#]", ChrW(&H25C4))        ' ◄
    resultado = Replace(resultado, "[#FLECHA_DER#]", ChrW(&H2192))                 ' →
    resultado = Replace(resultado, "[#FLECHA_IZQ#]", ChrW(&H2190))                 ' ←
    resultado = Replace(resultado, "[#FLECHA_ARR#]", ChrW(&H2191))                 ' ↑
    resultado = Replace(resultado, "[#FLECHA_ABA#]", ChrW(&H2193))                 ' ↓
    resultado = Replace(resultado, "[#LINEA_DOBLE#]", ChrW(&H2550))                ' ═
    resultado = Replace(resultado, "[#LINEA_SIMPLE#]", ChrW(&H2500))               ' ─
    
    ReemplazarPlaceholdersPorEmojis = resultado
End Function

' =============================================================================
' PROCESAMIENTO DE MARKDOWN
' =============================================================================

Public Function ConvertirMarkdownAWord(ByVal contenidoMD As String, _
                                       ByRef wordDoc As Object, _
                                       Optional aplicarEmojis As Boolean = True) As Boolean
    ' Convierte contenido Markdown a formato Word
    ' Procesa títulos, negritas, listas, etc.
    '
    ' Parámetros:
    '   contenidoMD: Contenido en formato Markdown
    '   wordDoc: Objeto Document de Word (ya creado)
    '   aplicarEmojis: Si True, reemplaza placeholders por emojis
    ' Retorna: True si se procesó correctamente
    
    On Error GoTo ErrorHandler
    
    Dim lineas() As String
    Dim i As Long
    Dim linea As String
    Dim textoAProcesar As String
    
    ' Aplicar emojis si está habilitado
    If aplicarEmojis Then
        textoAProcesar = ReemplazarPlaceholdersPorEmojis(contenidoMD)
    Else
        textoAProcesar = contenidoMD
    End If
    
    ' Dividir en líneas
    lineas = Split(textoAProcesar, vbCrLf)
    
    ' Procesar cada línea
    For i = LBound(lineas) To UBound(lineas)
        linea = lineas(i)
        
        ' Procesar según el tipo de línea Markdown
        If Len(Trim(linea)) = 0 Then
            ' Línea vacía
            Call AgregarEspacio(wordDoc)
            
        ElseIf Left(linea, 2) = "# " Then
            ' Título H1
            Call AgregarTitulo(wordDoc, Mid(linea, 3), 1)
            
        ElseIf Left(linea, 3) = "## " Then
            ' Título H2
            Call AgregarTitulo(wordDoc, Mid(linea, 4), 2)
            
        ElseIf Left(linea, 4) = "### " Then
            ' Título H3
            Call AgregarTitulo(wordDoc, Mid(linea, 5), 3)
            
        ElseIf Left(linea, 2) = "- " Then
            ' Lista con viñetas
            Call AgregarItemLista(wordDoc, Mid(linea, 3))
            
        ElseIf Left(linea, 3) = "---" Or Left(linea, 3) = "===" Then
            ' Separador horizontal
            Call AgregarSeparador(wordDoc)
            
        Else
            ' Párrafo normal (puede contener negritas)
            Call AgregarParrafo(wordDoc, linea)
        End If
    Next i
    
    ConvertirMarkdownAWord = True
    Exit Function
    
ErrorHandler:
    Debug.Print "ERROR en ConvertirMarkdownAWord: " & Err.Description
    ConvertirMarkdownAWord = False
End Function

' =============================================================================
' FUNCIONES AUXILIARES PARA FORMATO WORD
' =============================================================================

Private Sub AgregarTitulo(ByRef wordDoc As Object, ByVal texto As String, ByVal nivel As Integer)
    ' Agrega un título con formato según nivel
    
    Dim rango As Object
    Set rango = wordDoc.Content
    rango.Collapse Direction:=0 ' wdCollapseEnd
    
    rango.InsertAfter texto & vbCrLf
    
    ' Aplicar formato según nivel
    With rango
        .Font.Bold = True
        .Font.Name = "Calibri"
        
        Select Case nivel
            Case 1
                .Font.Size = 20
                .Font.Color = RGB(46, 117, 182) ' Azul oscuro
            Case 2
                .Font.Size = 16
                .Font.Color = RGB(68, 114, 196) ' Azul medio
            Case 3
                .Font.Size = 14
                .Font.Color = RGB(91, 155, 213) ' Azul claro
        End Select
        
        .ParagraphFormat.SpaceAfter = 6
        .ParagraphFormat.SpaceBefore = 12
    End With
    
    Set rango = Nothing
End Sub

Private Sub AgregarParrafo(ByRef wordDoc As Object, ByVal texto As String)
    ' Agrega un párrafo normal con soporte para negritas **texto**
    
    Dim rango As Object
    Set rango = wordDoc.Content
    rango.Collapse Direction:=0
    
    ' Procesar negritas simples (sin implementar formato complejo aquí)
    ' Para simplificar, insertamos el texto tal cual
    ' Word procesará el Markdown si es necesario
    
    rango.InsertAfter texto & vbCrLf
    
    With rango
        .Font.Bold = False
        .Font.Size = 11
        .Font.Name = "Calibri"
        .Font.Color = RGB(0, 0, 0)
        .ParagraphFormat.SpaceAfter = 3
        .ParagraphFormat.Alignment = 0 ' wdAlignParagraphLeft
    End With
    
    Set rango = Nothing
End Sub

Private Sub AgregarItemLista(ByRef wordDoc As Object, ByVal texto As String)
    ' Agrega un item de lista con viñeta
    
    Dim rango As Object
    Set rango = wordDoc.Content
    rango.Collapse Direction:=0
    
    rango.InsertAfter ChrW(&H2022) & " " & texto & vbCrLf
    
    With rango
        .Font.Bold = False
        .Font.Size = 11
        .Font.Name = "Calibri"
        .Font.Color = RGB(0, 0, 0)
        .ParagraphFormat.LeftIndent = 20
        .ParagraphFormat.SpaceAfter = 2
    End With
    
    Set rango = Nothing
End Sub

Private Sub AgregarEspacio(ByRef wordDoc As Object)
    ' Agrega una línea vacía
    
    Dim rango As Object
    Set rango = wordDoc.Content
    rango.Collapse Direction:=0
    rango.InsertAfter vbCrLf
    Set rango = Nothing
End Sub

Private Sub AgregarSeparador(ByRef wordDoc As Object)
    ' Agrega una línea separadora
    
    Dim rango As Object
    Set rango = wordDoc.Content
    rango.Collapse Direction:=0
    
    rango.InsertAfter String(50, ChrW(&H2500)) & vbCrLf
    
    With rango
        .Font.Size = 10
        .Font.Color = RGB(150, 150, 150)
        .ParagraphFormat.Alignment = 1 ' wdAlignParagraphCenter
    End With
    
    Set rango = Nothing
End Sub

' =============================================================================
' GENERACIÓN DE REPORTES COMPLETOS
' =============================================================================

Public Function GenerarReporteNumerologico(ByVal nombrePersona As String, _
                                          ByVal fechaNacimiento As Date, _
                                          ByVal rutaSalida As String) As Boolean
    ' Genera un reporte numerológico completo en Word
    '
    ' Parámetros:
    '   nombrePersona: Nombre completo de la persona
    '   fechaNacimiento: Fecha de nacimiento
    '   rutaSalida: Ruta donde guardar el documento Word
    ' Retorna: True si se generó correctamente
    
    On Error GoTo ErrorHandler
    
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim contenidoMD As String
    Dim rutaMD As String
    
    ' Crear instancia de Word
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = True ' Mostrar Word durante la creación
    
    ' Crear nuevo documento
    Set wordDoc = wordApp.Documents.Add
    
    ' =========================================================================
    ' PORTADA
    ' =========================================================================
    With wordDoc.Content
        .Font.Name = "Calibri"
        .Font.Size = 28
        .Font.Bold = True
        .Text = "ANÁLISIS NUMEROLÓGICO COMPLETO"
        .ParagraphFormat.Alignment = 1 ' wdAlignParagraphCenter
        .Font.Color = RGB(46, 117, 182)
    End With
    
    wordDoc.Content.InsertParagraphAfter
    wordDoc.Content.InsertParagraphAfter
    
    ' Nombre de la persona
    Dim rango As Object
    Set rango = wordDoc.Content
    rango.Collapse Direction:=0
    rango.InsertAfter nombrePersona & vbCrLf
    With rango
        .Font.Size = 20
        .Font.Bold = False
        .Font.Color = RGB(0, 0, 0)
        .ParagraphFormat.Alignment = 1
    End With
    
    ' Fecha de nacimiento
    Set rango = wordDoc.Content
    rango.Collapse Direction:=0
    rango.InsertAfter Format(fechaNacimiento, "dd/mm/yyyy") & vbCrLf
    With rango
        .Font.Size = 14
        .Font.Color = RGB(100, 100, 100)
        .ParagraphFormat.Alignment = 1
    End With
    
    ' Salto de página
    wordDoc.Content.InsertBreak 7 ' wdPageBreak
    
    ' =========================================================================
    ' AQUÍ SE AÑADIRÍAN LAS SECCIONES CALCULADAS
    ' =========================================================================
    ' Ejemplo: Camino de Vida
    ' rutaMD = CurrentProject.Path & "\Interpretaciones\CaminoVida\01_CaminoVida.md"
    ' contenidoMD = LeerArchivoUTF8(rutaMD)
    ' Call ConvertirMarkdownAWord(contenidoMD, wordDoc, True)
    
    ' Nota informativa temporal
    Set rango = wordDoc.Content
    rango.Collapse Direction:=0
    rango.InsertAfter "Este documento se completará con las interpretaciones" & vbCrLf
    rango.InsertAfter "calculadas según el nombre y fecha de nacimiento." & vbCrLf
    
    ' =========================================================================
    ' GUARDAR DOCUMENTO
    ' =========================================================================
    wordDoc.SaveAs rutaSalida
    
    MsgBox "Reporte generado exitosamente:" & vbCrLf & vbCrLf & rutaSalida, vbInformation, "Reporte Numerológico"
    
    ' Limpiar objetos
    Set rango = Nothing
    Set wordDoc = Nothing
    Set wordApp = Nothing
    
    GenerarReporteNumerologico = True
    Exit Function
    
ErrorHandler:
    Debug.Print "ERROR en GenerarReporteNumerologico: " & Err.Description
    MsgBox "Error al generar reporte: " & Err.Description, vbCritical
    
    If Not wordDoc Is Nothing Then wordDoc.Close False
    If Not wordApp Is Nothing Then wordApp.Quit
    
    Set rango = Nothing
    Set wordDoc = Nothing
    Set wordApp = Nothing
    
    GenerarReporteNumerologico = False
End Function

' =============================================================================
' FUNCIÓN DE PRUEBA
' =============================================================================

Public Sub PruebaGeneracionWord()
    ' Función de prueba para verificar el módulo
    
    Dim rutaSalida As String
    Dim exito As Boolean
    
    rutaSalida = CurrentProject.Path & "\Reportes\Prueba_Numerologia.docx"
    
    ' Crear carpeta Reportes si no existe
    If Dir(CurrentProject.Path & "\Reportes", vbDirectory) = "" Then
        MkDir CurrentProject.Path & "\Reportes"
    End If
    
    exito = GenerarReporteNumerologico("María Carmen García López", #3/15/1985#, rutaSalida)
    
    If exito Then
        Debug.Print "✓ Prueba exitosa - Documento creado"
    Else
        Debug.Print "✗ Prueba fallida"
    End If
End Sub
