Sub CrearReporteNumerologico()
    On Error GoTo ErrorHandler
    
    Dim wordApp As Object
    Dim wordDoc As Object
    
    ' Crear instancia de Word
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = True
    
    ' Crear nuevo documento
    Set wordDoc = wordApp.Documents.Add
    
    ' Agregar contenido
    With wordDoc.Content
        .Text = "ANÁLISIS NUMEROLÓGICO COMPLETO" & vbCrLf & vbCrLf
        .Font.Name = "Calibri"
        .Font.Size = 16
        .Font.Bold = True
        .ParagraphFormat.Alignment = 1 ' Centrado
    End With
    
    ' Agregar más contenido
    wordDoc.Content.InsertAfter vbCrLf & "Nombre: Ana María Santos Varela"
    
    ' Guardar
    wordDoc.SaveAs CurrentProject.Path & "\Reporte_Numerologia.docx"
    
    ' Limpiar
    Set wordDoc = Nothing
    Set wordApp = Nothing
    
    MsgBox "Reporte creado exitosamente", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Set wordDoc = Nothing
    Set wordApp = Nothing
End Sub

'==================================================================================

Function ConvertirMarkdownAWord(rutaMD As String, rutaWordSalida As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim contenidoMD As String
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim lineas() As String
    Dim i As Long
    Dim linea As String
    
    ' 1. Leer archivo Markdown en UTF-8
    contenidoMD = LeerArchivoUTF8(rutaMD)
    
    If Len(contenidoMD) = 0 Then
        MsgBox "No se pudo leer el archivo Markdown", vbExclamation
        ConvertirMarkdownAWord = False
        Exit Function
    End If
    
    ' 2. Crear documento Word
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False
    Set wordDoc = wordApp.Documents.Add
    
    ' 3. Procesar línea por línea
    lineas = Split(contenidoMD, vbCrLf)
    
    For i = LBound(lineas) To UBound(lineas)
        linea = lineas(i)
        
        ' Procesar según tipo de línea Markdown
        If Left(linea, 2) = "# " Then
            ' Título H1
            AgregarTitulo wordDoc, Mid(linea, 3), 1
            
        ElseIf Left(linea, 3) = "## " Then
            ' Título H2
            AgregarTitulo wordDoc, Mid(linea, 4), 2
            
        ElseIf Left(linea, 4) = "### " Then
            ' Título H3
            AgregarTitulo wordDoc, Mid(linea, 5), 3
            
        ElseIf Left(linea, 2) = "**" And Right(linea, 2) = "**" Then
            ' Texto en negrita
            AgregarParrafoNegrita wordDoc, Mid(linea, 3, Len(linea) - 4)
            
        ElseIf Left(linea, 2) = "- " Then
            ' Lista con viñetas
            AgregarItemLista wordDoc, Mid(linea, 3)
            
        ElseIf Len(Trim(linea)) > 0 Then
            ' Párrafo normal
            AgregarParrafo wordDoc, linea
            
        Else
            ' Línea vacía - espacio
            AgregarEspacio wordDoc
        End If
    Next i
    
    ' 4. Guardar documento
    wordDoc.SaveAs rutaWordSalida
    wordDoc.Close
    wordApp.Quit
    
    Set wordDoc = Nothing
    Set wordApp = Nothing
    
    ConvertirMarkdownAWord = True
    Exit Function
    
ErrorHandler:
    ConvertirMarkdownAWord = False
    If Not wordDoc Is Nothing Then wordDoc.Close False
    If Not wordApp Is Nothing Then wordApp.Quit
    Set wordDoc = Nothing
    Set wordApp = Nothing
End Function

' Funciones auxiliares para formatear Word
Private Sub AgregarTitulo(doc As Object, texto As String, nivel As Integer)
    Dim rango As Object
    Set rango = doc.Content
    rango.Collapse Direction:=0 ' wdCollapseEnd
    
    rango.InsertAfter texto & vbCrLf
    rango.Font.Bold = True
    
    Select Case nivel
        Case 1: rango.Font.Size = 18
        Case 2: rango.Font.Size = 14
        Case 3: rango.Font.Size = 12
    End Select
    
    Set rango = Nothing
End Sub

Private Sub AgregarParrafo(doc As Object, texto As String)
    Dim rango As Object
    Set rango = doc.Content
    rango.Collapse Direction:=0
    rango.InsertAfter texto & vbCrLf
    rango.Font.Bold = False
    rango.Font.Size = 11
    Set rango = Nothing
End Sub

Private Sub AgregarParrafoNegrita(doc As Object, texto As String)
    Dim rango As Object
    Set rango = doc.Content
    rango.Collapse Direction:=0
    rango.InsertAfter texto & vbCrLf
    rango.Font.Bold = True
    rango.Font.Size = 11
    Set rango = Nothing
End Sub

Private Sub AgregarItemLista(doc As Object, texto As String)
    Dim rango As Object
    Set rango = doc.Content
    rango.Collapse Direction:=0
    rango.InsertAfter "• " & texto & vbCrLf
    rango.Font.Bold = False
    rango.Font.Size = 11
    Set rango = Nothing
End Sub

Private Sub AgregarEspacio(doc As Object)
    Dim rango As Object
    Set rango = doc.Content
    rango.Collapse Direction:=0
    rango.InsertAfter vbCrLf
    Set rango = Nothing
End Sub

'==================================================================================

Sub GenerarReporteCompleto(nombrePersona As String, fechaNacimiento As Date)
    On Error GoTo ErrorHandler
    
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim rutaSalida As String
    
    ' Cálculos (usar tus clases)
    Dim caminoVida As clsCalculoCaminoVida
    Set caminoVida = New clsCalculoCaminoVida
    caminoVida.FechaNacimiento = fechaNacimiento
    caminoVida.Calcular
    
    ' Crear documento
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = True
    Set wordDoc = wordApp.Documents.Add
    
    ' === PORTADA ===
    With wordDoc.Content
        .Font.Name = "Calibri"
        .Font.Size = 24
        .Font.Bold = True
        .Text = "ANÁLISIS NUMEROLÓGICO COMPLETO"
        .ParagraphFormat.Alignment = 1 ' Centrado
    End With
    
    wordDoc.Content.InsertAfter vbCrLf & vbCrLf
    wordDoc.Content.InsertAfter nombrePersona & vbCrLf
    wordDoc.Content.InsertAfter Format(fechaNacimiento, "dd/mm/yyyy") & vbCrLf
    wordDoc.Content.InsertAfter vbCrLf & vbCrLf
    
    ' Salto de página
    wordDoc.Content.InsertBreak 7 ' wdPageBreak
    
    ' === CAMINO DE VIDA ===
    Dim rangoActual As Object
    Set rangoActual = wordDoc.Content
    rangoActual.Collapse Direction:=0
    
    rangoActual.InsertAfter "CAMINO DE VIDA - NÚMERO " & caminoVida.Resultado & vbCrLf & vbCrLf
    rangoActual.Font.Size = 16
    rangoActual.Font.Bold = True
    
    ' Leer interpretación desde archivo Markdown
    Dim rutaMD As String
    Dim interpretacion As String
    
    rutaMD = CurrentProject.Path & "\Interpretaciones\CaminoVida\" & _
             Format(caminoVida.Resultado, "00") & "_CaminoVida.md"
    
    interpretacion = LeerArchivoUTF8(rutaMD)
    
    ' Agregar interpretación (aquí podrías procesar el Markdown)
    rangoActual.InsertAfter interpretacion & vbCrLf & vbCrLf
    rangoActual.Font.Size = 11
    rangoActual.Font.Bold = False
    
    ' === GUARDAR ===
    rutaSalida = CurrentProject.Path & "\Reportes\" & _
                 Replace(nombrePersona, " ", "_") & "_Numerologia.docx"
    
    wordDoc.SaveAs rutaSalida
    
    MsgBox "Reporte generado: " & vbCrLf & rutaSalida, vbInformation
    
    Set rangoActual = Nothing
    Set wordDoc = Nothing
    Set wordApp = Nothing
    Set caminoVida = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

'==============================================================================


Function ReemplazarPlaceholdersPorEmojis(texto As String) As String
    ' Reemplaza placeholders por códigos de emoji VBA
    Dim resultado As String
    resultado = texto
    
    ' CELEBRACIÓN Y LOGROS
    resultado = Replace(resultado, "[#CELEBRACION#]", ChrW(&HD83C) & ChrW(&HDF89))
    resultado = Replace(resultado, "[#BRILLO#]", ChrW(&H2728))
    resultado = Replace(resultado, "[#ESTRELLA#]", ChrW(&HD83C) & ChrW(&HDF1F))
    resultado = Replace(resultado, "[#TROFEO#]", ChrW(&HD83C) & ChrW(&HDFC6))
    
    ' ADVERTENCIAS
    resultado = Replace(resultado, "[#ADVERTENCIA#]", ChrW(&H26A0) & ChrW(&HFE0F))
    resultado = Replace(resultado, "[#ALERTA#]", ChrW(&HD83D) & ChrW(&HDEA8))
    
    ' APROBACIÓN Y NEGACIÓN
    resultado = Replace(resultado, "[#CHECK#]", ChrW(&H2705))
    resultado = Replace(resultado, "[#CRUZ#]", ChrW(&H274C))
    
    ' TRABAJO Y ACCIÓN
    resultado = Replace(resultado, "[#MALETIN#]", ChrW(&HD83D) & ChrW(&HDCBC))
    resultado = Replace(resultado, "[#COHETE#]", ChrW(&HD83D) & ChrW(&HDE80))
    resultado = Replace(resultado, "[#DIANA#]", ChrW(&HD83C) & ChrW(&HDFAF))
    
    ' CONOCIMIENTO
    resultado = Replace(resultado, "[#LIBROS#]", ChrW(&HD83D) & ChrW(&HDCDA))
    resultado = Replace(resultado, "[#BOMBILLA#]", ChrW(&HD83D) & ChrW(&HDCA1))
    
    ' AMOR Y RELACIONES
    resultado = Replace(resultado, "[#CORAZON#]", ChrW(&H2764) & ChrW(&HFE0F))
    resultado = Replace(resultado, "[#APRETÓN_MANOS#]", ChrW(&HD83E) & ChrW(&HDD1D))
    
    ' DINERO
    resultado = Replace(resultado, "[#BOLSA_DINERO#]", ChrW(&HD83D) & ChrW(&HDCB0))
    resultado = Replace(resultado, "[#DIAMANTE#]", ChrW(&HD83D) & ChrW(&HDC8E))
    
    ' NATURALEZA
    resultado = Replace(resultado, "[#SEMILLA#]", ChrW(&HD83C) & ChrW(&HDF31))
    resultado = Replace(resultado, "[#FUEGO#]", ChrW(&HD83D) & ChrW(&HDD25))
    
    ' SÍMBOLOS BÁSICOS (Siempre funcionan)
    resultado = Replace(resultado, "[#TICK#]", ChrW(&H2713))
    resultado = Replace(resultado, "[#X#]", ChrW(&H2717))
    resultado = Replace(resultado, "[#PUNTO#]", ChrW(&H2022))
    resultado = Replace(resultado, "[#ESTRELLA_NEGRA#]", ChrW(&H2605))
    resultado = Replace(resultado, "[#TRIANGULO_DERECHA#]", ChrW(&H25BA))
    resultado = Replace(resultado, "[#LINEA_DOBLE#]", ChrW(&H2550))
    resultado = Replace(resultado, "[#LINEA_SIMPLE#]", ChrW(&H2500))
    
    ReemplazarPlaceholdersPorEmojis = resultado
End Function

' Uso en generación de Word:
Sub AgregarTextoConEmojis(wordDoc As Object, textoMD As String)
    Dim textoConEmojis As String
    textoConEmojis = ReemplazarPlaceholdersPorEmojis(textoMD)
    
    Dim rango As Object
    Set rango = wordDoc.Content
    rango.Collapse Direction:=0
    rango.InsertAfter textoConEmojis
End Sub

'========================================================================================


Sub GenerarReporteCompleto()
    Dim cv As clsCalculoCaminoVida
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim contenidoMD As String
    Dim rutaMD As String
    
    ' Calcular
    Set cv = New clsCalculoCaminoVida
    cv.FechaNacimiento = #3/15/1985#
    cv.Calcular
    
    ' Crear Word
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = True
    Set wordDoc = wordApp.Documents.Add
    
    ' Leer interpretación
    rutaMD = CurrentProject.Path & "\Interpretaciones\CaminoVida\" & _
             Format(cv.Resultado, "00") & "_CaminoVida.md"
    contenidoMD = LeerArchivoUTF8(rutaMD)
    
    ' Convertir a Word con emojis
    Call ConvertirMarkdownAWord(contenidoMD, wordDoc, True)
    
    ' Guardar
    wordDoc.SaveAs CurrentProject.Path & "\Reportes\MiReporte.docx"
    
    ' Limpiar
    Set cv = Nothing
    Set wordDoc = Nothing
    Set wordApp = Nothing
End Sub
