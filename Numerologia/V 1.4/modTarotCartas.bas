Attribute VB_Name = "modTarotCartas"
Option Compare Database
Option Explicit

Public Sub CrearTablaTarotNueva()

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim idx As DAO.Index
    
    Set db = CurrentDb
    
    ' Eliminar si existe
    On Error Resume Next
    db.TableDefs.Delete "tbmTarotCartas"
    On Error GoTo 0
    
    ' Crear tabla
    Set tdf = db.CreateTableDef("tbmTarotCartas")
    
    ' Campos
    With tdf
        .Fields.Append .CreateField("NumeroGlobal", dbByte)
        .Fields.Append .CreateField("NombreCarta", dbText, 100)
        .Fields.Append .CreateField("Palo", dbText, 20)
        .Fields.Append .CreateField("Elemento", dbText, 20)
        .Fields.Append .CreateField("NumeroInterno", dbByte)
        .Fields.Append .CreateField("Figura", dbText, 20)
        .Fields.Append .CreateField("NombreFicheroFinal", dbText, 255)
        .Fields.Append .CreateField("RutaMarkdown", dbText, 255)
    End With
    
    ' Añadir la tabla a la BD
    db.TableDefs.Append tdf
    
    ' Crear índice de clave primaria sobre NumeroGlobal
    Set idx = tdf.CreateIndex("PK_tbmTarotCartas")
    With idx
        .Fields.Append .CreateField("NumeroGlobal")
        .Primary = True
        .Unique = True
    End With
    tdf.Indexes.Append idx
    
    MsgBox "Tabla tbmTarotCartas creada con clave primaria en NumeroGlobal.", vbInformation

End Sub

Public Sub PoblarOrdenJavane()

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim orden As Integer
    
    Set db = CurrentDb
    
    ' Asegurar que el campo existe
    On Error Resume Next
    db.Execute "ALTER TABLE tbmTarotCartas ADD COLUMN OrdenJavane INTEGER"
    On Error GoTo 0
    
    Set rs = db.OpenRecordset("SELECT * FROM tbmTarotCartas")
    
    Do While Not rs.EOF
        
        If rs!palo = "Mayor" Then
            rs.Edit
            rs!OrdenJavane = Null
            rs.Update
        
        Else
            Select Case rs!figura
                
                Case "Rey"
                    orden = 1
                Case "Reina"
                    orden = 2
                Case "Caballero"
                    orden = 3
                Case "Sota"
                    orden = 4
                
                Case Else
                    ' Números del As al 10
                    orden = 4 + rs!numeroInterno
            End Select
            
            rs.Edit
            rs!OrdenJavane = orden
            rs.Update
        End If
        
        rs.MoveNext
    Loop
    
    MsgBox "Campo OrdenJavane completado correctamente.", vbInformation

End Sub

Public Sub GenerarNumeroJavane()

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim contador As Integer
    Dim palo As Variant
    Dim figuraOrden As Variant
    
    Set db = CurrentDb
    
    ' 1. Asegurar que el campo existe
'    On Error Resume Next
'    db.Execute "ALTER TABLE tbmTarotCartas ADD COLUMN NumeroJavane INTEGER"
    On Error GoTo 0
    
    ' 2. Resetear valores
'    db.Execute "UPDATE tbmTarotCartas SET NumeroJavane = Null"
    
    contador = -1 '0
    
    ' 3. Primero: Arcanos Mayores en su orden natural
    sql = "SELECT * FROM tbmTarotCartas WHERE Palo='Mayor' ORDER BY NumeroGlobal"
    Set rs = db.OpenRecordset(sql)
    
    Do While Not rs.EOF
        contador = contador + 1
        rs.Edit
        rs!NumeroJavane = contador
        rs.Update
        rs.MoveNext
    Loop
    rs.Close
    
    ' 4. Orden de palos según Javane
    Dim palos As Variant
    palos = Array("Bastos", "Copas", "Espadas", "Oros")
    
    ' 5. Orden interno de Javane
    ' Rey, Reina, Caballero, Sota, As, 2..10
    Dim figuras As Variant
    figuras = Array("Rey", "Reina", "Caballero", "Sota")
    
    ' 6. Recorrer cada palo
    For Each palo In palos
        
        ' 6.1 Figuras primero
        For Each figuraOrden In figuras
            
            sql = "SELECT * FROM tbmTarotCartas " & _
                  "WHERE Palo='" & palo & "' AND Figura='" & figuraOrden & "' " & _
                  "ORDER BY NumeroInterno"
            
            Set rs = db.OpenRecordset(sql)
            
            Do While Not rs.EOF
                contador = contador + 1
                rs.Edit
                rs!NumeroJavane = contador
                rs.Update
                rs.MoveNext
            Loop
            rs.Close
        Next figuraOrden
        
        ' 6.2 Números del As al 10
        sql = "SELECT * FROM tbmTarotCartas " & _
              "WHERE Palo='" & palo & "' AND (Figura Is Null OR Figura='') " & _
              "ORDER BY NumeroInterno"
        
        Set rs = db.OpenRecordset(sql)
        
        Do While Not rs.EOF
            contador = contador + 1
            rs.Edit
            rs!NumeroJavane = contador
            rs.Update
            rs.MoveNext
        Loop
        rs.Close
        
    Next palo
    
    MsgBox "NumeroJavane generado correctamente.", vbInformation

End Sub

'------------------------------------------------------------------------------
Public Sub VerificarTablaTarotPitagorico()

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
'    Dim i As Long
    Dim esperado As Long
    Dim nombreEsperado As String
    Dim ficheroEsperado As String
'    Dim palo As String
'    Dim figura As String
    Dim numeroInterno As Long
    Dim nombreCarta As String
    Dim errores As Long
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM tbmTarotCartas ORDER BY NumeroGlobal")
    
    Debug.Print "=== VERIFICACIÓN TABLA PITAGÓRICA ==="
    
    esperado = 0
    errores = 0
    
    Do While Not rs.EOF
        
        ' 1. Verificar consecutividad
        If rs!numeroGlobal <> esperado Then
            Debug.Print "? Error: NumeroGlobal no consecutivo ? "; rs!numeroGlobal; " (esperado "; esperado; ")"
            errores = errores + 1
        End If
        
        ' 2. Verificar coherencia NombreCarta ? NumeroInterno
        nombreCarta = Trim(rs!nombreCarta)
        numeroInterno = rs!numeroInterno
        
        If rs!palo <> "Mayor" Then
            Select Case numeroInterno
                Case 1: nombreEsperado = "As"
                Case 2: nombreEsperado = "Dos"
                Case 3: nombreEsperado = "Tres"
                Case 4: nombreEsperado = "Cuatro"
                Case 5: nombreEsperado = "Cinco"
                Case 6: nombreEsperado = "Seis"
                Case 7: nombreEsperado = "Siete"
                Case 8: nombreEsperado = "Ocho"
                Case 9: nombreEsperado = "Nueve"
                Case 10: nombreEsperado = "Diez"
                Case 11: nombreEsperado = "Sota"
                Case 12: nombreEsperado = "Caballero"
                Case 13: nombreEsperado = "Reina"
                Case 14: nombreEsperado = "Rey"
            End Select
            
            If LCase(nombreCarta) <> LCase(nombreEsperado) Then
                Debug.Print "? Error: NombreCarta no coincide con NumeroInterno ? "; rs!numeroGlobal, nombreCarta, "?", nombreEsperado
                errores = errores + 1
            End If
        End If
        
        ' 3. Verificar coherencia Figura
        If rs!figura & "" <> "" Then
            If LCase(rs!figura) <> LCase(nombreCarta) Then
                Debug.Print "? Error: Figura incorrecta ? "; rs!numeroGlobal, rs!figura, "?", nombreCarta
                errores = errores + 1
            End If
        End If
        
        ' 4. Verificar NombreFicheroFinal
        ficheroEsperado = Format(rs!numeroGlobal, "00") & "_" & Replace(nombreCarta, " ", "_") & _
                          IIf(rs!palo = "Mayor", "", "_de_" & rs!palo) & ".md"
        
        If LCase(rs!NombreFicheroFinal) <> LCase(ficheroEsperado) Then
            Debug.Print "? Error: NombreFicheroFinal incorrecto ? "; rs!numeroGlobal, rs!NombreFicheroFinal, "?", ficheroEsperado
            errores = errores + 1
        End If
        
        ' 5. Verificar RutaMarkdown
        If Replace(LCase(rs!RutaMarkdown), """", "") <> LCase(rs!NombreFicheroFinal) Then
            Debug.Print "? Error: RutaMarkdown incorrecta ? "; rs!numeroGlobal, rs!RutaMarkdown
            errores = errores + 1
        End If
        
        esperado = esperado + 1
        rs.MoveNext
    Loop
    
    Debug.Print "=== FIN VERIFICACIÓN ==="
    Debug.Print "Errores detectados: "; errores

End Sub
'------------------------------------------------------------------------------

Public Sub PoblarTablaTarotCorregida()

    Dim db As DAO.Database
    Set db = CurrentDb
    
    ' Borrado previo de la tabla
    db.Execute "DELETE * FROM tbmTarotCartas"
    
    ' Arcanos Mayores
    Call InsertarMayoresCorregido(db)
    
    ' Arcanos Menores
    Call InsertarMenoresCorregido(db, 23, "Bastos", "Fuego")
    Call InsertarMenoresCorregido(db, 37, "Copas", "Agua")
    Call InsertarMenoresCorregido(db, 51, "Espadas", "Aire")
    Call InsertarMenoresCorregido(db, 65, "Oros", "Tierra")
    
    MsgBox "Tabla tbmTarotCartas corregida y poblada.", vbInformation

End Sub

Private Sub InsertarMayoresCorregido(db As DAO.Database)

    Dim nombres As Variant
    nombres = Array( _
        "El Loco (0)", "El Mago", "La Suma Sacerdotisa", "La Emperatriz", "El Emperador", _
        "El Hierofante", "Los Amantes", "El Carro", "La Fuerza", "El Ermitaño", _
        "La Rueda de la Fortuna", "La Justicia", "El Colgado", "La Muerte", _
        "La Templanza", "El Diablo", "La Torre", "La Estrella", "La Luna", _
        "El Sol", "El Juicio", "El Mundo", "El Loco" _
    )
    
    Dim i As Integer
    Dim nombreArchivo As String
    Dim nombreLimpio As String
    Dim rs As Recordset
    
    Set rs = db.OpenRecordset("tbmTarotCartas")
    
    
    For i = 0 To 22
        
        If i = 0 Then
            nombreArchivo = ""
        Else
            nombreLimpio = LimpiarNombre(nombres(i))
            nombreArchivo = Format(i, "00") & "_" & nombreLimpio & ".md"
        End If
        
        rs.AddNew
        
        rs("NumeroGlobal") = i
        rs("NombreCarta") = CStr(nombres(i))
        rs("Palo") = "Mayor"
        rs("Elemento") = "Espíritu"
        rs("NumeroInterno") = i
        rs("Figura") = ""
        rs("NombreFicheroFinal") = nombreArchivo
        rs("RutaMarkdown") = Chr(34) & nombreArchivo & Chr(34)
        
        rs.Update
        
        'db.Execute "INSERT INTO tbmTarotCartas " & _
                   "(NumeroGlobal, NombreCarta, Palo, Elemento, NumeroInterno, Figura, NombreFicheroFinal, RutaMarkdown) " & _
                   "VALUES (" & i & ", '" & CStr(nombres(i)) & "', 'Mayor', 'Espíritu', " & i & ", '', '" & nombreArchivo & "', '" & nombreArchivo & "');"
    Next i

End Sub

Private Sub InsertarMenoresCorregido(db As DAO.Database, inicio As Integer, palo As String, elemento As String)

    Dim nombres As Variant
    nombres = Array("As", "Dos", "Tres", "Cuatro", "Cinco", "Seis", "Siete", "Ocho", "Nueve", "Diez", _
                    "Sota", "Caballero", "Reina", "Rey")
    
    Dim i As Integer
    Dim numeroGlobal As Integer
    Dim numeroInterno As Integer
    Dim figura As String
    Dim nombreCarta As String
    Dim nombreArchivo As String
    
    For i = 0 To 13
        
        numeroGlobal = inicio + i
        numeroInterno = i + 1
        
        nombreCarta = nombres(i)
        figura = IIf(i >= 10, nombres(i), "")
        
        nombreArchivo = numeroGlobal & "_" & LimpiarNombre(nombreCarta) & "_de_" & palo & ".md"
        
        db.Execute "INSERT INTO tbmTarotCartas " & _
                   "(NumeroGlobal, NombreCarta, Palo, Elemento, NumeroInterno, Figura, NombreFicheroFinal, RutaMarkdown) " & _
                   "VALUES (" & numeroGlobal & ", '" & nombreCarta & "', '" & palo & "', '" & elemento & "', " & numeroInterno & ", '" & figura & "', '" & nombreArchivo & "', '" & nombreArchivo & "');"
    
    Next i

End Sub

Private Function LimpiarNombre(txt As Variant) As String
    Dim t As String
    t = CStr(txt)
    t = Replace(t, " ", "_")
    t = Replace(t, "á", "a")
    t = Replace(t, "é", "e")
    t = Replace(t, "í", "i")
    t = Replace(t, "ó", "o")
    t = Replace(t, "ú", "u")
    t = Replace(t, "ñ", "n")
    LimpiarNombre = t
End Function



Public Sub RenombrarYCopiarTarotMD()

    Dim fso As New FileSystemObject
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim carpetaBase As String
    Dim carpetasOrigen As Variant
    Dim carpetaDestino As String
    Dim carpetaBackup As String
    Dim i As Integer
    Dim archivoOriginal As String
'    Dim archivoNuevo As String
    Dim rutaOrigen As String
    Dim rutaDestino As String
    Dim rutaBackup As String
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM tbmTarotCartas WHERE NombreFicheroFinal IS NOT NULL")
    
    carpetaBase = "N:\Numerologia\Interpretaciones\Tarot\0_Completo\"  ' ? AJUSTA ESTA RUTA
    carpetasOrigen = Array("0_Mayores", "1_Bastos", "2_Copas", "3_Espadas", "4_Oros")
    carpetaDestino = carpetaBase & "TarotMD\"
'    carpetaBackup = carpetaBase & "BackupOriginales\"
    
    ' Crear carpetas si no existen
    If Not fso.FolderExists(carpetaDestino) Then fso.CreateFolder carpetaDestino
'    If Not fso.FolderExists(carpetaBackup) Then fso.CreateFolder carpetaBackup
    
    Do While Not rs.EOF
        
        archivoOriginal = ""
        
        ' Buscar archivo original en carpetas
        For i = 0 To UBound(carpetasOrigen)
            rutaOrigen = carpetaBase & carpetasOrigen(i) & "\"
            
            archivoOriginal = BuscarArchivoEnCarpeta(fso, rutaOrigen, rs!numeroInterno, rs!nombreCarta, rs!palo)
            
            If archivoOriginal <> "" Then Exit For
        Next i
        
        If archivoOriginal = "" Then
            Debug.Print "Archivo no encontrado para " & rs!numeroGlobal & " - " & rs!nombreCarta
        Else
            rutaBackup = carpetaBackup & fso.GetFileName(archivoOriginal)
            rutaDestino = carpetaDestino & rs!NombreFicheroFinal
            
            ' Copiar a backup
'            fso.CopyFile archivoOriginal, rutaBackup, True
            
            ' Renombrar y mover
            fso.CopyFile archivoOriginal, rutaDestino, True
            
            Debug.Print "Copiado: " & archivoOriginal & " ? " & rutaDestino
        End If
        
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    MsgBox "Renombrado y copia completados.", vbInformation

End Sub

Private Function BuscarArchivoEnCarpeta(fso As FileSystemObject, carpeta As String, numeroInterno As Integer, nombreCarta As String, palo As String) As String

    Dim nombreEsperado As String
    Dim archivo As File
    
    ' Construir nombre esperado según patrón real
    nombreEsperado = palo & "_" & Format(numeroInterno, "00") & "_" & LimpiarNombre(nombreCarta) & ".md"
    
    For Each archivo In fso.GetFolder(carpeta).Files
        If LCase(fso.GetFileName(archivo)) = LCase(nombreEsperado) Then
            BuscarArchivoEnCarpeta = archivo.Path
            Exit Function
        End If
    Next archivo
    
    BuscarArchivoEnCarpeta = ""

End Function

Public Sub RenombrarImagenesTarot()

    Dim fso As New FileSystemObject
    Dim carpetaImg As String
    Dim archivo As File
    Dim fName As String
    Dim numero As String
    Dim nombreLimpio As String
    Dim nombreConGuiones As String
    Dim nuevoNombre As String
    Dim rutaNueva As String
    
    carpetaImg = "C:\TuRutaFinal\img\"   ' ? AJUSTA ESTA RUTA
    
    For Each archivo In fso.GetFolder(carpetaImg).Files
        If LCase(fso.GetExtensionName(archivo.Name)) = "png" Then
            
            fName = archivo.Name
            
            ' Extraer número (dos primeros caracteres)
            numero = Left(fName, 2)
            
            ' Extraer el resto del nombre
            nombreLimpio = Trim(Mid(fName, 3))
            
            ' Sustituir espacios por _
            nombreConGuiones = Replace(nombreLimpio, " ", "_")
            
            ' Construir nuevo nombre
            nuevoNombre = numero & "_" & nombreConGuiones
            
            ' Ruta final
            rutaNueva = carpetaImg & nuevoNombre
            
            ' Renombrar si es diferente
            If archivo.Path <> rutaNueva Then
                archivo.Name = nuevoNombre
            End If
            
        End If
    Next archivo
    
    MsgBox "Imágenes renombradas correctamente.", vbInformation

End Sub

Public Sub DiagnosticoImagenesTarot()

    Dim fso As New FileSystemObject
    Dim carpetaImg As String
    Dim archivo As File
    Dim fName As String
    Dim numeroImg As Integer
    Dim nombreImg As String
    Dim nombreNormalizado As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim encontrado As Boolean
    Dim nombreTablaNorm As String
    
    carpetaImg = "N:\Numerologia\Interpretaciones\Tarot\0_Completo\TarotMD\img\"   ' ? AJUSTA ESTA RUTA
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT NumeroGlobal, NombreCarta FROM tbmTarotCartas")
    
    Debug.Print "=== DIAGNÓSTICO DE IMÁGENES TAROT ==="
    
    ' Recorrer todas las imágenes
    For Each archivo In fso.GetFolder(carpetaImg).Files
        If LCase(fso.GetExtensionName(archivo.Name)) = "png" Then
            
            fName = archivo.Name
            
            ' Extraer número (dos primeros caracteres)
            numeroImg = CInt(Left(fName, 2))
            
            ' Extraer nombre (resto)
            nombreImg = Trim(Mid(fName, 3))
            nombreImg = Replace(nombreImg, ".png", "")
            
            ' Normalizar nombre de imagen
            nombreNormalizado = Normalizar(nombreImg)
            
            encontrado = False
            
            rs.MoveFirst
            Do While Not rs.EOF
                
                nombreTablaNorm = Normalizar(rs!nombreCarta)
                
                ' Coincidencia por número y nombre
                If rs!numeroGlobal = numeroImg And nombreTablaNorm = nombreNormalizado Then
                    encontrado = True
                    Exit Do
                End If
                
                rs.MoveNext
            Loop
            
            If Not encontrado Then
                Debug.Print "? DESAJUSTE ? Imagen: "; fName; _
                            " | Número: "; numeroImg; _
                            " | Nombre: "; nombreImg
            End If
            
        End If
    Next archivo
    
    Debug.Print "=== FIN DEL DIAGNÓSTICO ==="

End Sub

Private Function Normalizar(txt As String) As String
    Dim t As String
    t = LCase(txt)
    
    t = Replace(t, "á", "a")
    t = Replace(t, "é", "e")
    t = Replace(t, "í", "i")
    t = Replace(t, "ó", "o")
    t = Replace(t, "ú", "u")
    t = Replace(t, "ñ", "n")
    
    t = Replace(t, "_", " ")
    
    ' Quitar dobles espacios
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    
    Normalizar = Trim(t)
End Function

Public Sub RenombrarImagenesSegunTabla()

    Dim fso As New FileSystemObject
    Dim carpetaImg As String
    Dim archivo As File
    Dim fName As String
    Dim nombreImg As String
    Dim nombreNormImg As String
    Dim numeroImg As Integer
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim nombreTabla As String
    Dim nombreNormTabla As String
    Dim nuevoNombre As String
    Dim rutaNueva As String
    Dim encontrado As Boolean
    
    carpetaImg = "N:\Numerologia\Interpretaciones\Tarot\0_Completo\TarotMD\img\"   ' AJUSTA SI ES NECESARIO
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT NumeroGlobal, NombreCarta, Palo, NombreFicheroFinal FROM tbmTarotCartas")
    
    Debug.Print "=== RENOMBRADO DE IMÁGENES SEGÚN TABLA PITAGÓRICA ==="
    
    For Each archivo In fso.GetFolder(carpetaImg).Files
        
        If LCase(fso.GetExtensionName(archivo.Name)) = "png" Then
            
            fName = archivo.Name
            
            ' Número original del archivo (para desambiguar El Loco)
            numeroImg = CInt(Left(fName, 2))
            
            ' Quitar número inicial y extensión
            nombreImg = Trim(Mid(fName, 3))
            nombreImg = Replace(nombreImg, ".png", "")
            
            ' Normalizar nombre de imagen
            nombreNormImg = Normalizar(nombreImg)
            
            encontrado = False
            
            rs.MoveFirst
            Do While Not rs.EOF
                
                nombreTabla = rs!nombreCarta & IIf(rs!palo = "Mayor", "", " de " & rs!palo)
                nombreNormTabla = Normalizar(nombreTabla)
                
                ' Caso general: emparejar por nombre
                If nombreNormTabla = nombreNormImg Then
                    
                    ' Caso especial: El Loco (0) y El Loco (22)
                    If nombreNormImg = "el loco" Then
                        ' Si el archivo empieza por 00 ? NumeroGlobal 0
                        ' Si el archivo empieza por 22 ? NumeroGlobal 22
                        If rs!numeroGlobal <> numeroImg Then
                            rs.MoveNext
                            GoTo SiguienteIteracion
                        End If
                    End If
                    
                    ' Construir nombre final correcto
                    nuevoNombre = Replace(rs!NombreFicheroFinal, ".md", ".png")
                    rutaNueva = carpetaImg & nuevoNombre
                    
                    ' Evitar error si ya existe un archivo con ese nombre
                    If LCase(archivo.Name) <> LCase(nuevoNombre) Then
                        If fso.FileExists(rutaNueva) Then
                            Debug.Print "? No renombrado (ya existe): "; fName; " ? "; nuevoNombre
                        Else
                            Debug.Print "? Emparejado: "; fName; " ? "; nuevoNombre
                            archivo.Name = nuevoNombre
                        End If
                    Else
                        Debug.Print "? Ya correcto: "; fName
                    End If
                    
                    encontrado = True
                    Exit Do
                End If
                
SiguienteIteracion:
                rs.MoveNext
            Loop
            
            If Not encontrado Then
                Debug.Print "? Sin coincidencia: "; fName
            End If
            
        End If
    Next archivo
    
    Debug.Print "=== FIN DEL RENOMBRADO ==="

End Sub

''--- Guardar temporal con BOM (ADODB SIEMPRE lo pone)
'stmTemp.Position = 0
'
'Set stmOut = CreateObject("ADODB.Stream")
'With stmOut
'    .Type = 1: .Open
'End With
'
'stmTemp.CopyTo stmOut
'stmTemp.Close
'
'' Guardar el temporal con BOM
'Dim rutaTemporal As String
'rutaTemporal = archivo.Path & ".tmp"
'
'stmOut.SaveToFile rutaTemporal, 2
'stmOut.Close
'
''--- Crear archivo sin BOM en binario
'Dim rutaSinBOM As String
'rutaSinBOM = archivo.Path & ".nobom"
'
'Dim fIn As Integer, fOut As Integer
'Dim size As Long
'Dim b() As Byte
'
'fIn = FreeFile
'Open rutaTemporal For Binary As #fIn
'    size = LOF(fIn)
'    If size > 3 Then
'        ReDim b(1 To size - 3)
'        Get #fIn, 4, b   ' saltar los 3 bytes del BOM
'    Else
'        ReDim b(0)
'    End If
'Close #fIn
'
'fOut = FreeFile
'Open rutaSinBOM For Binary As #fOut
'    If size > 3 Then Put #fOut, , b
'Close #fOut
'
''--- Reemplazar el original usando FSO
'If fso.FileExists(archivo.Path) Then
'    fso.DeleteFile archivo.Path, True
'End If
'
'fso.MoveFile rutaSinBOM, archivo.Path
'
''--- Limpiar temporal
'If fso.FileExists(rutaTemporal) Then
'    fso.DeleteFile rutaTemporal, True
'End If

Public Sub ActualizarMarkdownTarot_Clasico_UnicoProceso()

    Const adTypeBinary = 1
    Const adTypeText = 2
    Const adModeReadWrite = 3
    Const adSaveCreateOverWrite = 2

    '===========================
    ' Declaraciones limpias
    '===========================
    Dim fso As Object
    Dim carpetaMD As String, carpetaBackup As String
    Dim archivo As Object

    Dim stmIn As Object
    Dim stmTemp As Object
    Dim UTFStream As Object
    Dim BinaryStream As Object

    Dim linea As String
    Dim tituloOriginal As String, tituloLimpio As String
    Dim nombreImagen As String, lineaImagen As String
    Dim EsCabecera As Boolean
    Dim posGuion As Long, posibleNumero As String

    Dim contenido As String, lineas As Variant
    Dim idx As Long

    Dim rutaTemporal As String
    Dim rutaFinal As String

    '===========================
    ' Rutas
    '===========================
    carpetaMD = "N:\Numerologia\Interpretaciones\Tarot\0_Completo\TarotMD\"
    carpetaBackup = "N:\Numerologia\Interpretaciones\Tarot\0_Completo\TarotMD_Backup\"

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(carpetaBackup) Then fso.CreateFolder carpetaBackup

    '===========================
    ' Procesar cada archivo
    '===========================
    For Each archivo In fso.GetFolder(carpetaMD).Files
        If LCase(fso.GetExtensionName(archivo.Name)) = "md" Then

            '--- Copia de seguridad
            fso.CopyFile archivo.Path, carpetaBackup & archivo.Name, True

            Debug.Print "Procesando: "; archivo.Name;
            
            If archivo.Name = "23_As_de_Bastos.md" Then
                Stop
            End If
            
            '===========================
            ' Leer archivo UTF-8
            '===========================
            Set stmIn = CreateObject("ADODB.Stream")
            With stmIn
                .Type = adTypeText
                .Charset = "utf-8"
                .Open
                .LoadFromFile archivo.Path
                contenido = .ReadText
                .Close
            End With

            '===========================
            ' Crear stream temporal (UTF-8 con BOM)
            '===========================
            Set stmTemp = CreateObject("ADODB.Stream")
            With stmTemp
                .Type = adTypeText
                .Charset = "utf-8"
                .Open
            End With

            '===========================
            ' Procesar líneas
            '===========================
            lineas = Split(contenido, vbLf)
            EsCabecera = True

            For idx = LBound(lineas) To UBound(lineas)

                linea = Trim$(lineas(idx))

                '--- Título
                If Left$(linea, 2) = "# " And EsCabecera Then

                    tituloOriginal = linea
                    tituloLimpio = tituloOriginal

                    posGuion = InStr(tituloOriginal, "-")
                    If posGuion > 0 Then
                        posibleNumero = Trim(Mid$(tituloOriginal, 3, posGuion - 3))
                        If IsNumeric(posibleNumero) Then
                            tituloLimpio = "# " & Trim(Mid$(tituloOriginal, posGuion + 1))
                        End If
                    End If

                    stmTemp.WriteText tituloLimpio & vbLf

                '--- Imagen
                ElseIf Left$(linea, 2) = "![" And EsCabecera Then

                    nombreImagen = Replace(archivo.Name, ".md", ".png")
                    lineaImagen = "![" & Mid$(tituloLimpio, 3) & "](img/" & nombreImagen & ")"
                    stmTemp.WriteText lineaImagen & vbLf

                '--- Fin de cabecera
                ElseIf Left$(linea, 3) = "## " Then

                    EsCabecera = False
                    stmTemp.WriteText linea & vbLf

                Else
                    stmTemp.WriteText linea & vbLf
                End If

            Next idx

            '===========================
            ' Guardar temporal con BOM
            '===========================
            stmTemp.Position = 0
            rutaTemporal = archivo.Path & ".tmp"
            stmTemp.SaveToFile rutaTemporal, 2
            stmTemp.Close

            '===========================
            ' Quitar BOM usando tu método ADODB
            '===========================
            rutaFinal = archivo.Path

            Set UTFStream = CreateObject("ADODB.Stream")
            Set BinaryStream = CreateObject("ADODB.Stream")

            ' Abrir temporal con BOM
            UTFStream.Type = adTypeText
            UTFStream.Mode = adModeReadWrite
            UTFStream.Charset = "UTF-8"
            UTFStream.Open
            UTFStream.LoadFromFile rutaTemporal

            ' Saltar BOM
            UTFStream.Position = 3

            ' Copiar a binario
            BinaryStream.Type = adTypeBinary
            BinaryStream.Mode = adModeReadWrite
            BinaryStream.Open
            UTFStream.CopyTo BinaryStream

            ' Guardar sin BOM directamente en el archivo original
            BinaryStream.SaveToFile rutaFinal, adSaveCreateOverWrite

            UTFStream.Close
            BinaryStream.Close

            ' Borrar temporal
            fso.DeleteFile rutaTemporal, True

            Debug.Print " --> COMPLETADO"
        End If
        
    Next archivo

    MsgBox "Markdown actualizado correctamente.", vbInformation

End Sub


'Public Sub ActualizarMarkdownTarot_Clasico_UnicoProceso()
'
'    Dim fso As Object
'    Dim carpetaMD As String, carpetaBackup As String
'    Dim archivo As Object
'    Dim stmIn As Object, stmTemp As Object
'    Dim linea As String
'    Dim tituloOriginal As String, tituloLimpio As String
'    Dim nombreImagen As String, lineaImagen As String
'    Dim EsCabecera As Boolean
'    Dim posGuion As Long, posibleNumero As String
'
'    Dim contenido As String, lineas As Variant
'    Dim idx As Long
'
'    Dim rutaTemporal As String
'    Dim rutaFinal As String
'
'    carpetaMD = "N:\Numerologia\Interpretaciones\Tarot\0_Completo\TarotMD\"
'    carpetaBackup = "N:\Numerologia\Interpretaciones\Tarot\0_Completo\TarotMD_Backup\"
'
'    Set fso = CreateObject("Scripting.FileSystemObject")
'    If Not fso.FolderExists(carpetaBackup) Then fso.CreateFolder carpetaBackup
'
'    For Each archivo In fso.GetFolder(carpetaMD).Files
'        If LCase(fso.GetExtensionName(archivo.Name)) = "md" Then
'
'            '--- Copia de seguridad
'            fso.CopyFile archivo.Path, carpetaBackup & archivo.Name, True
'
'            '--- Leer archivo original UTF-8
'            Set stmIn = CreateObject("ADODB.Stream")
'            With stmIn
'                .Type = 2: .Charset = "utf-8": .Open
'                .LoadFromFile archivo.Path
'                contenido = .ReadText
'                .Close
'            End With
'
'            '--- Crear stream temporal (UTF-8 con BOM)
'            Set stmTemp = CreateObject("ADODB.Stream")
'            With stmTemp
'                .Type = 2: .Charset = "utf-8": .Open
'            End With
'
'            '--- Dividir por LF (UNIX)
'            lineas = Split(contenido, vbLf)
'
'            EsCabecera = True
'
'            For idx = LBound(lineas) To UBound(lineas)
'
'                linea = Trim$(lineas(idx))
'
'                '--- Detectar título
'                If Left$(linea, 2) = "# " And EsCabecera Then
'
'                    tituloOriginal = linea
'                    tituloLimpio = tituloOriginal
'
'                    posGuion = InStr(tituloOriginal, "-")
'                    If posGuion > 0 Then
'                        posibleNumero = Trim(Mid$(tituloOriginal, 3, posGuion - 3))
'                        If IsNumeric(posibleNumero) Then
'                            tituloLimpio = "# " & Trim(Mid$(tituloOriginal, posGuion + 1))
'                        End If
'                    End If
'
'                    stmTemp.WriteText tituloLimpio & vbLf
'
'                '--- Detectar imagen antigua
'                ElseIf Left$(linea, 2) = "![" And EsCabecera Then
'
'                    nombreImagen = Replace(archivo.Name, ".md", ".png")
'                    lineaImagen = "![" & Mid$(tituloLimpio, 3) & "](img/" & nombreImagen & ")"
'
'                    stmTemp.WriteText lineaImagen & vbLf
'
'                '--- Detectar inicio de contenido real
'                ElseIf Left$(linea, 3) = "## " Then
'
'                    EsCabecera = False
'                    stmTemp.WriteText linea & vbLf
'
'                Else
'                    stmTemp.WriteText linea & vbLf
'                End If
'
'            Next idx
'
'            '--- Guardar temporal con BOM
'            stmTemp.Position = 0
'            rutaTemporal = archivo.Path & ".tmp"
'            stmTemp.SaveToFile rutaTemporal, 2
'            stmTemp.Close
'
'            '===========================================================
'            '   Quitar BOM usando tu método ADODB (adaptado a UTF-8)
'            '===========================================================
'
'            rutaFinal = archivo.Path & ".final"
'
'            Dim UTFStream As Object
'            Dim BinaryStream As Object
'
'            Set UTFStream = CreateObject("ADODB.Stream")
'            Set BinaryStream = CreateObject("ADODB.Stream")
'
'            ' Abrir el archivo temporal (UTF-8 con BOM)
'            UTFStream.Type = 2
'            UTFStream.Mode = 3
'            UTFStream.Charset = "UTF-8"
'            UTFStream.Open
'            UTFStream.LoadFromFile rutaTemporal
'
'            ' Saltar los 3 bytes del BOM
'            UTFStream.Position = 3
'
'            ' Copiar el resto a un stream binario
'            BinaryStream.Type = 1
'            BinaryStream.Mode = 3
'            BinaryStream.Open
'
'            UTFStream.CopyTo BinaryStream
'
'            ' Guardar sin BOM
'            BinaryStream.SaveToFile rutaFinal, 2
'
'            UTFStream.Close
'            BinaryStream.Close
'
'            '===========================================================
'
'            '--- Reemplazar el original
'            If fso.FileExists(archivo.Path) Then
'                fso.DeleteFile archivo.Path, True
'            End If
'
'            fso.MoveFile rutaFinal, archivo.Path
'
'            '--- Limpiar temporal
'            If fso.FileExists(rutaTemporal) Then
'                fso.DeleteFile rutaTemporal, True
'            End If
'
'        End If
'    Next archivo
'
'    MsgBox "Markdown actualizado correctamente.", vbInformation
'
'End Sub

Public Function UTF8_RemoveBOM(ByVal utf8WithBOM As String, ByVal utf8NoBOM As String)

    Const adTypeBinary = 1
    Const adTypeText = 2
    Const adModeReadWrite = 3
    Const adSaveCreateOverWrite = 2

    Dim UTFStream As Object
    Dim BinaryStream As Object

    Set UTFStream = CreateObject("ADODB.Stream")
    Set BinaryStream = CreateObject("ADODB.Stream")

    '--- Abrir el archivo UTF-8 con BOM
    UTFStream.Type = adTypeText
    UTFStream.Mode = adModeReadWrite
    UTFStream.Charset = "UTF-8"
    UTFStream.Open
    UTFStream.LoadFromFile utf8WithBOM

    '--- Saltar los 3 bytes del BOM
    UTFStream.Position = 3

    '--- Copiar el resto a un stream binario
    BinaryStream.Type = adTypeBinary
    BinaryStream.Mode = adModeReadWrite
    BinaryStream.Open

    UTFStream.CopyTo BinaryStream

    '--- Guardar sin BOM
    BinaryStream.SaveToFile utf8NoBOM, adSaveCreateOverWrite

    '--- Cerrar
    UTFStream.Close
    BinaryStream.Close

    Set UTFStream = Nothing
    Set BinaryStream = Nothing

End Function

'Public Sub ActualizarMarkdownTarot_Clasico()
'
'    Dim fso As Object
'    Dim carpetaMD As String, carpetaBackup As String
'    Dim archivo As Object
'    Dim stmIn As Object, stmTemp As Object
'    Dim linea As String
'    Dim tituloOriginal As String, tituloLimpio As String
'    Dim nombreImagen As String, lineaImagen As String
'    Dim EsCabecera As Boolean
'    Dim posGuion As Long, posibleNumero As String
'
'    Dim contenido As String, lineas As Variant
'    Dim idx As Long
'
'    Dim rutaTemporal As String
'    Dim rutaSinBOM As String
'    Dim fIn As Integer, fOut As Integer
'    Dim size As Long
'    Dim b() As Byte
'
'    carpetaMD = "N:\Numerologia\Interpretaciones\Tarot\0_Completo\TarotMD\"
'    carpetaBackup = "N:\Numerologia\Interpretaciones\Tarot\0_Completo\TarotMD_Backup\"
'
'    Set fso = CreateObject("Scripting.FileSystemObject")
'    If Not fso.FolderExists(carpetaBackup) Then fso.CreateFolder carpetaBackup
'
'    For Each archivo In fso.GetFolder(carpetaMD).Files
'        If LCase(fso.GetExtensionName(archivo.Name)) = "md" Then
'
'            '--- Copia de seguridad
'            fso.CopyFile archivo.Path, carpetaBackup & archivo.Name, True
'
'            '--- Leer archivo original (UTF-8, pero ADODB no toca los LF)
'            Set stmIn = CreateObject("ADODB.Stream")
'            With stmIn
'                .Type = 2: .Charset = "utf-8": .Open
'                .LoadFromFile archivo.Path
'                contenido = .ReadText
'                .Close
'            End With
'
'            '--- Crear stream temporal (UTF-8 con BOM, luego lo quitamos)
'            Set stmTemp = CreateObject("ADODB.Stream")
'            With stmTemp
'                .Type = 2: .Charset = "utf-8": .Open
'            End With
'
'            '--- Dividir por LF (UNIX)
'            lineas = Split(contenido, vbLf)
'
'            EsCabecera = True
'
'            For idx = LBound(lineas) To UBound(lineas)
'
'                linea = Trim$(lineas(idx))
'
'                '--- Detectar título
'                If Left$(linea, 2) = "# " And EsCabecera Then
'
'                    tituloOriginal = linea
'                    tituloLimpio = tituloOriginal
'
'                    ' Limpiar número si es Mayor
'                    posGuion = InStr(tituloOriginal, "-")
'                    If posGuion > 0 Then
'                        posibleNumero = Trim(Mid$(tituloOriginal, 3, posGuion - 3))
'                        If IsNumeric(posibleNumero) Then
'                            tituloLimpio = "# " & Trim(Mid$(tituloOriginal, posGuion + 1))
'                        End If
'                    End If
'
'                    stmTemp.WriteText tituloLimpio & vbLf
'
'                '--- Detectar imagen antigua
'                ElseIf Left$(linea, 2) = "![" And EsCabecera Then
'
'                    nombreImagen = Replace(archivo.Name, ".md", ".png")
'                    lineaImagen = "![" & Mid$(tituloLimpio, 3) & "](img/" & nombreImagen & ")"
'
'                    stmTemp.WriteText lineaImagen & vbLf
'
'                '--- Detectar inicio de contenido real
'                ElseIf Left$(linea, 3) = "## " Then
'
'                    EsCabecera = False
'                    stmTemp.WriteText linea & vbLf
'
'                Else
'                    '--- Copiar línea tal cual
'                    stmTemp.WriteText linea & vbLf
'                End If
'
'            Next idx
'
'            '--- Guardar temporal con BOM
'            stmTemp.Position = 0
'
'            rutaTemporal = archivo.Path & ".tmp"
'            stmTemp.SaveToFile rutaTemporal, 2
'            stmTemp.Close
'
'            '--- Crear archivo sin BOM en binario
'            rutaSinBOM = archivo.Path & ".nobom"
'
'            fIn = FreeFile
'            Open rutaTemporal For Binary As #fIn
'                size = LOF(fIn)
'                If size > 3 Then
'                    ReDim b(1 To size - 3)
'                    Get #fIn, 4, b   ' saltar BOM
'                Else
'                    ReDim b(0)
'                End If
'            Close #fIn
'
'            fOut = FreeFile
'            Open rutaSinBOM For Binary As #fOut
'                If size > 3 Then Put #fOut, , b
'            Close #fOut
'
'            '--- Reemplazar el original
'            If fso.FileExists(archivo.Path) Then
'                fso.DeleteFile archivo.Path, True
'            End If
'
'            fso.MoveFile rutaSinBOM, archivo.Path
'
'            '--- Limpiar temporal
'            If fso.FileExists(rutaTemporal) Then
'                fso.DeleteFile rutaTemporal, True
'            End If
'
'        End If
'    Next archivo
'
'    MsgBox "Markdown actualizado correctamente.", vbInformation
'
'End Sub



'Public Sub ActualizarMarkdownTarot_Clasico()
'
'    Dim fso As Object
'    Dim carpetaMD As String, carpetaBackup As String
'    Dim archivo As Object
'    Dim stmIn As Object, stmTemp As Object, stmOut As Object
'    Dim linea As String
'    Dim tituloOriginal As String, tituloLimpio As String
'    Dim nombreImagen As String, lineaImagen As String
'    Dim leyendoCabecera As Boolean
'    Dim posGuion As Long, posibleNumero As String
'    Dim EsCabecera As Boolean
'
'    Dim idx As Long
'    Dim contenido As String, lineas As Variant
'
'
'    carpetaMD = "N:\Numerologia\Interpretaciones\Tarot\0_Completo\TarotMD\"
'    carpetaBackup = "N:\Numerologia\Interpretaciones\Tarot\0_Completo\TarotMD_Backup\"
'
'    Set fso = CreateObject("Scripting.FileSystemObject")
'    If Not fso.FolderExists(carpetaBackup) Then fso.CreateFolder carpetaBackup
'
'    For Each archivo In fso.GetFolder(carpetaMD).Files
'        If LCase(fso.GetExtensionName(archivo.Name)) = "md" Then
'
'            If archivo.Name = "23_As_de_Bastos.md" Then
'                Stop
'            End If
'
'            fso.CopyFile archivo.Path, carpetaBackup & archivo.Name, True
'
'            '--- Abrir original
'            Set stmIn = CreateObject("ADODB.Stream")
'            With stmIn
'                .Type = 2: .Charset = "utf-8": .Open
'                .LoadFromFile archivo.Path
'                contenido = .ReadText: .Close
'            End With
'
'            '--- Crear temporal
'            Set stmTemp = CreateObject("ADODB.Stream")
'            With stmTemp
'                .Type = 2: .Charset = "utf-8": .Open
'            End With
'
'            lineas = Split(contenido, vbLf)
'            'leyendoCabecera = True
'
'            EsCabecera = True
'
'            'Do Until stmIn.EOS
'            For idx = LBound(lineas) To UBound(lineas)
'
'                'linea = stmIn.ReadText(-2) ' Leer línea completa
'                linea = Trim$(lineas(idx))
'
'                'If leyendoCabecera Then
'
'                '--- Detectar título
'                If Left$(Trim$(linea), 2) = "# " And EsCabecera = True Then
'
'                    tituloOriginal = Trim$(linea)
'                    'EsCabecera = True
'
'                    ' Limpiar número si es Mayor
'                    posGuion = InStr(tituloOriginal, "-")
'                    tituloLimpio = tituloOriginal
'
'                    If posGuion > 0 Then
'                        posibleNumero = Trim(Mid$(tituloOriginal, 3, posGuion - 3))
'                        If IsNumeric(posibleNumero) Then
'                            tituloLimpio = "# " & Trim(Mid$(tituloOriginal, posGuion + 1))
'                        End If
'                    Else
'                        tituloLimpio = Trim$(linea)
'                    End If
'
'                    ' Escribir título limpio
'                    stmTemp.WriteText tituloLimpio & vbLf '& vbLf
'
'                ElseIf Left$(Trim$(linea), 2) = "![" And EsCabecera Then
'                    ' Nueva imagen
'                    nombreImagen = Replace(archivo.Name, ".md", ".png")
'                    lineaImagen = "![" & Mid$(tituloLimpio, 3) & "](img/" & nombreImagen & ")"
'
'                    stmTemp.WriteText lineaImagen & vbLf '& vbLf
'
'                    ' No escribir la línea original
'                    'GoTo SiguienteLinea
'
'                    '--- Detectar fin de cabecera
'                ElseIf Left$(Trim$(linea), 3) = "## " Then
'                    'leyendoCabecera = False
'                    EsCabecera = False
'                    stmTemp.WriteText linea & vbLf
'
'                    'GoTo SiguienteLinea
'
'                Else 'If EsCabecera Then
'
'                    'stmTemp.WriteText linea
'                    stmTemp.WriteText linea & vbLf
'                End If
'
'                ' Mientras estemos en cabecera, ignorar todo
'                'GoTo SiguienteLinea
'
'                'Else
'                '--- Copiar contenido real
'                'stmTemp.WriteText linea & vbCrLf
'                'End If
'            Next idx
'            'SiguienteLinea:
'
'            'Loop
'
'            'stmIn.Close
'
'            '            '--- Guardar temporal con BOM
'            '            stmTemp.Position = 0
'
'            '            ' Convertir a binario para quitar BOM
'            '            Set stmOut = CreateObject("ADODB.Stream")
'            '            With stmOut
'            '                .Type = 1: .Open
'            '            End With
'            '
'            '            stmTemp.CopyTo stmOut
'            '            stmTemp.Close
'            '
'            '            ' Saltar BOM
'            '            If stmOut.size >= 3 Then stmOut.Position = 3
'            '
'            '            ' Guardar sin BOM
'            '            stmOut.SaveToFile archivo.Path, 2
'            '            stmOut.Close
'
'            '--- Guardar temporal con BOM (ADODB SIEMPRE lo pone)
'            stmTemp.Position = 0
'
'            Set stmOut = CreateObject("ADODB.Stream")
'            With stmOut
'                .Type = 1: .Open
'            End With
'
'            stmTemp.CopyTo stmOut
'            stmTemp.Close
'
'            ' Guardar el temporal con BOM
'            Dim rutaTemporal As String
'            rutaTemporal = archivo.Path & ".tmp"
'
'            stmOut.SaveToFile rutaTemporal, 2
'            stmOut.Close
'
'            '--- Crear archivo sin BOM en binario
'            Dim rutaSinBOM As String
'            rutaSinBOM = archivo.Path & ".nobom"
'
'            Dim fIn As Integer, fOut As Integer
'            Dim size As Long
'            Dim b() As Byte
'
'            fIn = FreeFile
'            Open rutaTemporal For Binary As #fIn
'            size = LOF(fIn)
'            If size > 3 Then
'                ReDim b(1 To size - 3)
'                Get #fIn, 4, b   ' saltar los 3 bytes del BOM
'            Else
'                ReDim b(0)
'            End If
'            Close #fIn
'
'            fOut = FreeFile
'            Open rutaSinBOM For Binary As #fOut
'            If size > 3 Then Put #fOut, , b
'            Close #fOut
'
'            '--- Reemplazar el original usando FSO
'            If fso.FileExists(archivo.Path) Then
'                fso.DeleteFile archivo.Path, True
'            End If
'
'            fso.MoveFile rutaSinBOM, archivo.Path
'
'            '--- Limpiar temporal
'            If fso.FileExists(rutaTemporal) Then
'                fso.DeleteFile rutaTemporal, True
'            End If
'
'
'        End If
'    Next archivo
'
'    MsgBox "Markdown actualizado correctamente.", vbInformation
'
'End Sub



Public Function ConvertToUTF8_NoBOM(ByVal ansiFile As String, ByVal utf8File As String)
    
    Const adTypeBinary = 1
    Const adTypeText = 2
    
'    Const adModeRead = 1
    Const adModeReadWrite = 3
    
'    Const adSaveCreateNotExist = 1
    Const adSaveCreateOverWrite = 2
    
    
    Dim UTFStream As Object 'New ADODB.stream
    Dim ANSIStream As Object 'New ADODB.stream
    Dim BinaryStream As Object 'New ADODB.stream

    Set UTFStream = CreateObject("ADODB.Stream")
    Set ANSIStream = CreateObject("ADODB.Stream")
    Set BinaryStream = CreateObject("ADODB.Stream")


    ANSIStream.Type = 2
    ANSIStream.Mode = adModeReadWrite
    ANSIStream.Charset = "iso-8859-1"
    ANSIStream.Open
    ANSIStream.LoadFromFile ansiFile '"C:\Users\eavj6\Desktop\mibat.txt"  'ANSI File
    
    UTFStream.Type = adTypeText
    UTFStream.Mode = adModeReadWrite
    UTFStream.Charset = "UTF-8"
    UTFStream.Open
    ANSIStream.CopyTo UTFStream
    

    UTFStream.Position = 3 'skip BOM
    BinaryStream.Type = adTypeBinary
    BinaryStream.Mode = adModeReadWrite
    BinaryStream.Open

    UTFStream.CopyTo BinaryStream

    BinaryStream.SaveToFile utf8File, adSaveCreateOverWrite '"C:\Users\eavj6\Desktop\mibat.txt"
    BinaryStream.Flush
    BinaryStream.Close

    Set UTFStream = Nothing
    Set ANSIStream = Nothing
    Set BinaryStream = Nothing

End Function

'Public Sub ActualizarMarkdownTarot_SuperSimple()
'
'    Dim fso As Object
'    Dim carpetaMD As String, carpetaBackup As String
'    Dim archivo As Object
'    Dim stmIn As Object, stmTemp As Object, stmOut As Object
'    Dim contenido As String, nuevoContenido As String
'    Dim lineas() As String
'    Dim i As Long, idxTitulo As Long, idxDescripcion As Long
'    Dim tituloOriginal As String, tituloLimpio As String
'    Dim posGuion As Long, posibleNumero As String
'    Dim nombreImagen As String, lineaImagen As String
'
'    carpetaMD = "N:\Numerologia\Interpretaciones\Tarot\0_Completo\TarotMD\"
'    carpetaBackup = "N:\Numerologia\Interpretaciones\Tarot\0_Completo\TarotMD_Backup\"
'
'    Set fso = CreateObject("Scripting.FileSystemObject")
'    If Not fso.FolderExists(carpetaBackup) Then fso.CreateFolder carpetaBackup
'
'    For Each archivo In fso.GetFolder(carpetaMD).Files
'        If LCase(fso.GetExtensionName(archivo.Name)) = "md" Then
'
'            ' Copia de seguridad
'            fso.CopyFile archivo.Path, carpetaBackup & archivo.Name, True
'
'            ' Leer archivo completo
'            Set stmIn = CreateObject("ADODB.Stream")
'            With stmIn
'                .Type = 2: .Charset = "utf-8": .Open
'                .LoadFromFile archivo.Path
'                contenido = .ReadText
'                .Close
'            End With
'
'            If Len(contenido) = 0 Then GoTo Siguiente
'
'            lineas = Split(contenido, vbCrLf)
'
'            '-----------------------------------------
'            ' 1) BUSCAR TÍTULO (# ...)
'            '-----------------------------------------
'            idxTitulo = -1
'            For i = 0 To UBound(lineas)
'                If Left$(Trim$(lineas(i)), 2) = "# " Then
'                    idxTitulo = i
'                    Exit For
'                End If
'            Next i
'            If idxTitulo = -1 Then GoTo Siguiente
'
'            tituloOriginal = Trim$(lineas(idxTitulo))
'            If Left$(tituloOriginal, 1) = ChrW(&HFEFF) Then tituloOriginal = Mid$(tituloOriginal, 2)
'
'            ' Limpiar número si es Mayor
'            tituloLimpio = tituloOriginal
'            posGuion = InStr(tituloOriginal, "-")
'            If posGuion > 0 Then
'                posibleNumero = Trim(Mid$(tituloOriginal, 3, posGuion - 3))
'                If IsNumeric(posibleNumero) Then
'                    tituloLimpio = "# " & Trim(Mid$(tituloOriginal, posGuion + 1))
'                End If
'            End If
'
'            '-----------------------------------------
'           ' 2) BUSCAR ## (inicio del contenido real)
'            '-----------------------------------------
'            idxDescripcion = -1
'            For i = idxTitulo + 1 To UBound(lineas)
'                If Left$(Trim$(lineas(i)), 3) = "## " Then
'                    idxDescripcion = i
'                    Exit For
'                End If
'            Next i
'            If idxDescripcion = -1 Then GoTo Siguiente
'
'            '-----------------------------------------
'            ' 3) CONSTRUIR CABECERA NUEVA
'            '-----------------------------------------
'            nombreImagen = Replace(archivo.Name, ".md", ".png")
'            lineaImagen = "![" & Mid$(tituloLimpio, 3) & "](img/" & nombreImagen & ")"
'
'            nuevoContenido = tituloLimpio & vbCrLf & vbCrLf & lineaImagen & vbCrLf & vbCrLf
'
'            '-----------------------------------------
'            ' 4) COPIAR DESDE ## HASTA EL FINAL
'            '-----------------------------------------
'            For i = idxDescripcion To UBound(lineas)
'                nuevoContenido = nuevoContenido & lineas(i)
'                If i < UBound(lineas) Then nuevoContenido = nuevoContenido & vbCrLf
'            Next i
'
'            '-----------------------------------------
'            ' 5) GUARDAR SIN BOM
'            '-----------------------------------------
'            Set stmTemp = CreateObject("ADODB.Stream")
'            With stmTemp
'                .Type = 2: .Charset = "utf-8": .Open
'                .WriteText nuevoContenido
'                .Position = 0
'            End With
'
'            Set stmOut = CreateObject("ADODB.Stream")
'            With stmOut
'                .Type = 1: .Open
'            End With
'
'            stmTemp.CopyTo stmOut
'            stmTemp.Close
'
'            If stmOut.Size >= 3 Then stmOut.Position = 3
'            stmOut.SaveToFile archivo.Path, 2
'            stmOut.Close
'
'Siguiente:
'        End If
'    Next archivo
'
'    MsgBox "Markdown actualizado correctamente.", vbInformation
'
'End Sub

