Attribute VB_Name = "bad modReconstruccionClases"
'Option Compare Database
'Option Explicit
'
'' ============================================================
'' RECONSTRUCTOR AUTOMÁTICO DE CLASES PARA ACCESS 2019
'' ============================================================
'' Este script:
''   1. Exporta todas las clases a archivos .cls
''   2. Crea archivos .cls nuevos con cabecera correcta
''   3. Copia el contenido original sin modificarlo
''   4. Elimina los módulos corruptos del proyecto
''   5. Importa las clases reconstruidas
'' ============================================================
'
'Public Sub ReconstruirTodasLasClases()
'    Dim comp As Object
'    Dim rutaTemp As String
'    Dim rutaOriginal As String
'    Dim rutaNueva As String
'    Dim nombre As String
'
'    rutaTemp = CurrentProject.Path & "\_clases_reconstruidas\"
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
'            Application.VBE.ActiveVBProject.VBComponents(nombre).Export rutaOriginal
'
'            ' 2. Crear archivo reconstruido
'            rutaNueva = rutaTemp & nombre & ".cls"
'            ReconstruirArchivoClase rutaOriginal, rutaNueva, nombre
'
'        End If
'    Next comp
'
'    ' 3. Eliminar clases corruptas del proyecto
'    Debug.Print "Eliminando clases antiguas..."
'    For Each comp In Application.VBE.ActiveVBProject.VBComponents
'        If comp.Type = vbext_ct_ClassModule Then
'            Application.VBE.ActiveVBProject.VBComponents.Remove comp
'        End If
'    Next comp
'
'    ' 4. Importar clases reconstruidas
'    Debug.Print "Importando clases reconstruidas..."
'    Dim archivo As String
'    archivo = Dir(rutaTemp & "*.cls")
'
'    Do While archivo <> ""
'        Application.VBE.ActiveVBProject.VBComponents.Import rutaTemp & archivo
'        archivo = Dir
'    Loop
'
'    MsgBox "Reconstrucción completada. Todas las clases han sido regeneradas correctamente.", vbInformation
'End Sub
'
'' ============================================================
'' RECONSTRUYE UN ARCHIVO .CLS CON CABECERA CORRECTA
'' ============================================================
'Private Sub ReconstruirArchivoClase(rutaOriginal As String, rutaNueva As String, nombreClase As String)
'    Dim fIn As Integer, fOut As Integer
'    Dim linea As String
'    Dim contenido As String
'
'    fIn = FreeFile
'    Open rutaOriginal For Input As #fIn
'    contenido = Input$(LOF(fIn), fIn)
'    Close #fIn
'
'    fOut = FreeFile
'    Open rutaNueva For Output As #fOut
'
'    ' CABECERA CORRECTA DE CLASE
'    Print #fOut, "VERSION 1.0 CLASS"
'    Print #fOut, "BEGIN"
'    Print #fOut, "  MultiUse = -1  'True"
'    Print #fOut, "END"
'    Print #fOut, "Attribute VB_Name = """ & nombreClase & """"
'    Print #fOut, "Option Compare Database"
'    Print #fOut, "Option Explicit"
'    Print #fOut, ""
'
'    ' Insertar contenido original SIN cabecera previa
'    Dim lineas() As String
'    Dim i As Long
'
'    lineas = Split(contenido, vbCrLf)
'
'    For i = 0 To UBound(lineas)
'        If Not EsLineaCabecera(lineas(i)) Then
'            Print #fOut, lineas(i)
'        End If
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
