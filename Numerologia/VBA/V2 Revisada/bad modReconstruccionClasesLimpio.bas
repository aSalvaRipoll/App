Attribute VB_Name = "bad modReconstruccionClasesLimpio"
'Option Compare Database
'Option Explicit
'
'' ============================================================
'' RECONSTRUCTOR AUTOMÁTICO DE CLASES (VERSIÓN LIMPIA)
'' ============================================================
'' Este script:
''   1. Exporta todas las clases a archivos .cls
''   2. Elimina cabeceras antiguas, Option Compare/Explicit duplicados
''   3. Inserta cabecera correcta y limpia
''   4. Elimina clases corruptas del proyecto
''   5. Importa las clases reconstruidas
'' ============================================================
'
'Public Sub ReconstruirClasesLimpias()
'    Dim comp As Object
'    Dim rutaTemp As String
'    Dim rutaOriginal As String
'    Dim rutaNueva As String
'    Dim nombre As String
'
'    rutaTemp = CurrentProject.Path & "\_clases_reconstruidas_limpias\"
'    MkDirIfNotExists rutaTemp
'
'    Debug.Print "Reconstruyendo clases en: " & rutaTemp
'
'    ' 1. Exportar todas las clases
'    For Each comp In Application.VBE.ActiveVBProject.VBComponents
'        If comp.Type = vbext_ct_ClassModule Then
'
'            nombre = comp.Name
'            rutaOriginal = rutaTemp & nombre & "_ORIGINAL.cls"
'
'            Debug.Print "Exportando: "; nombre
'
'            comp.Export rutaOriginal
'
'            ' 2. Crear archivo reconstruido limpio
'            rutaNueva = rutaTemp & nombre & ".cls"
'            ReconstruirArchivoClaseLimpio rutaOriginal, rutaNueva, nombre
'
'        End If
'    Next comp
'
'    ' 3. Eliminar clases antiguas del proyecto
'    Debug.Print
'    Debug.Print "Eliminando clases antiguas..."
'    For Each comp In Application.VBE.ActiveVBProject.VBComponents
'        If comp.Type = vbext_ct_ClassModule Then
'            Application.VBE.ActiveVBProject.VBComponents.Remove comp
'        End If
'    Next comp
'
'    ' 4. Importar clases reconstruidas
'    Debug.Print
'    Debug.Print "Importando clases reconstruidas..."
'    Dim archivo As String
''    archivo = Dir(rutaTemp & "*.cls")
''
''    Do While archivo <> ""
''        Application.VBE.ActiveVBProject.VBComponents.Import rutaTemp & archivo
''        archivo = Dir
''    Loop
'    archivo = Dir(rutaTemp & "*.cls")
'
'    Do While archivo <> ""
'
'        ' Saltar archivos originales
'        If InStr(1, archivo, "_ORIGINAL", vbTextCompare) = 0 Then
'            Debug.Print "Importando: "; archivo
'            Application.VBE.ActiveVBProject.VBComponents.Import rutaTemp & archivo
'        Else
'            Debug.Print "Saltando archivo original: "; archivo
'        End If
'
'        archivo = Dir
'    Loop
'
'    MsgBox "Reconstrucción completada. Todas las clases han sido regeneradas correctamente y limpiamente.", vbInformation
'End Sub
'
'' ============================================================
'' RECONSTRUYE UN ARCHIVO .CLS LIMPIO
'' ============================================================
'Private Sub ReconstruirArchivoClaseLimpio(rutaOriginal As String, rutaNueva As String, nombreClase As String)
'    Dim fIn As Integer, fOut As Integer
'    Dim contenido As String
'    Dim lineas() As String
'    Dim i As Long
'    Dim linea As String
'
'    ' Leer contenido original
'    fIn = FreeFile
'    Open rutaOriginal For Input As #fIn
'    contenido = Input$(LOF(fIn), fIn)
'    Close #fIn
'
'    lineas = Split(contenido, vbCrLf)
'
'    ' Crear archivo limpio
'    fOut = FreeFile
'    Open rutaNueva For Output As #fOut
'
'    ' CABECERA PERFECTA
'    Print #fOut, "VERSION 1.0 CLASS"
'    Print #fOut, "BEGIN"
'    Print #fOut, "  MultiUse = -1  'True"
'    Print #fOut, "END"
'    Print #fOut, "Attribute VB_Name = """ & nombreClase & """"
'    Print #fOut, "Attribute VB_GlobalNameSpace = False"
'    Print #fOut, "Attribute VB_Creatable = False"
'    Print #fOut, "Attribute VB_PredeclaredId = False"
'    Print #fOut, "Attribute VB_Exposed = False"
'    Print #fOut, "Option Compare Database"
'    Print #fOut, "Option Explicit"
'    Print #fOut, ""
'
'    ' Insertar contenido original sin cabeceras ni Option Compare/Explicit duplicados
'    For i = 0 To UBound(lineas)
'        linea = Trim(lineas(i))
'
'        If EsLineaCabecera(linea) Then GoTo Siguiente
'        If linea Like "Option Compare*" Then GoTo Siguiente
'        If linea Like "Option Explicit*" Then GoTo Siguiente
'
'        Print #fOut, lineas(i)
'
'Siguiente:
'    Next i
'
'    Close #fOut
'End Sub
'
'' ============================================================
'' DETECTA LÍNEAS DE CABECERA A ELIMINAR
'' ============================================================
'Private Function EsLineaCabecera(linea As String) As Boolean
'    linea = Trim(linea)
'
'    If linea Like "VERSION *" Then EsLineaCabecera = True: Exit Function
'    If linea Like "BEGIN*" Then EsLineaCabecera = True: Exit Function
'    If linea Like "END*" Then EsLineaCabecera = True: Exit Function
'    If InStr(1, linea, "Attribute VB_Name", vbTextCompare) > 0 Then EsLineaCabecera = True: Exit Function
'    If InStr(1, linea, "MultiUse", vbTextCompare) > 0 Then EsLineaCabecera = True: Exit Function
'    If InStr(1, linea, "Attribute VB_", vbTextCompare) > 0 Then EsLineaCabecera = True: Exit Function
'
'    EsLineaCabecera = False
'End Function
'
'' ============================================================
'' CREA CARPETA SI NO EXISTE
'' ============================================================
'Private Sub MkDirIfNotExists(ruta As String)
'    On Error Resume Next
'    MkDir ruta
'    On Error GoTo 0
'End Sub
'
'
