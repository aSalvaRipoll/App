Attribute VB_Name = "Módulo1"
Option Compare Database
Option Explicit

Dim fso As Object


'Public Sub CrearTablaIndice()
'
'    Dim db As DAO.Database
'    Dim t As DAO.TableDef
'
'    Set db = CurrentDb
'
'    On Error Resume Next
'    db.TableDefs.Delete "tbmIndiceInterpretaciones"
'    On Error GoTo 0
'
'    Set t = db.CreateTableDef("tbmIndiceInterpretaciones")
'
'    With t
'        .Fields.Append .CreateField("IDIndice", dbLong)
'        '.Fields("IDIndice").Attributes = dbAutoIncrField
'
'        .Fields.Append .CreateField("Categoria", dbText, 50)
'        .Fields.Append .CreateField("Valor", dbText, 20)
'        .Fields.Append .CreateField("Fichero", dbText, 255)
'        .Fields.Append .CreateField("Ruta", dbText, 255)
'        .Fields.Append .CreateField("EsKarmico", dbBoolean)
'        .Fields.Append .CreateField("EsMaestro", dbBoolean)
'    End With
'
'    db.TableDefs.Append t
'
'    MsgBox "Tabla Índice creada correctamente.", vbInformation
'
'End Sub

Public Sub GenerarIndiceInterpretaciones()

    
    Dim carpetaBase As Object
    Dim carpeta As Object
    Dim db As DAO.Database
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set db = CurrentDb
    
    Dim RutaBase As String
    'rutaBase = CurrentProject.Path & "\interpretaciones"
    RutaBase = "N:\Numerologia\Interpretaciones"
    
    ' Limpiar tabla
    db.Execute "DELETE FROM tbmIndiceInterpretaciones"
    
    Set carpetaBase = fso.GetFolder(RutaBase)
    
    ' Recorrer carpetas raíz
    For Each carpeta In carpetaBase.SubFolders
        Call RecorrerCarpeta(carpeta, carpeta.Name)
    Next carpeta
    
    MsgBox "Índice generado correctamente.", vbInformation

End Sub

Public Sub RecorrerCarpeta(ByVal carpeta As Object, ByVal categoria As String, Optional ByVal subcategoria As String = "")

    Dim archivo As Object
    Dim subcarp As Object
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim Valor As String
    Dim esKarmico As Boolean
    Dim esMaestro As Boolean
    
    'Dim categoria As String
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("tbmIndiceInterpretaciones", dbOpenDynaset)
    
    'categoria = carpeta.Name
    
    Debug.Print categoria; " --> "; subcategoria
    ' Procesar archivos .md
    For Each archivo In carpeta.Files
        DoEvents
        'If LCase(Right(archivo.Name, 3)) = ".md" Then
        If LCase(fso.GetExtensionName(archivo.Name)) = "md" Then
            
            Valor = ExtraerValorDesdeNombre(archivo.Name)
            esKarmico = (InStr(Valor, "-") > 0)
            esMaestro = (Valor = "11" Or Valor = "22" Or Valor = "33" Or Valor = "44")
            
            rs.AddNew
            rs!IDIndice = sMax("IDIndice", "tbmIndiceInterpretaciones") + 1 ' AutoNext("IDIndice", "tbmIndiceInterpretaciones")
            rs!categoria = categoria
            rs!subcategoria = subcategoria
            rs!Valor = Valor
            rs!Fichero = archivo.Name
            rs!Ruta = archivo.Path
            rs!esKarmico = esKarmico
            rs!esMaestro = esMaestro
            rs.Update
            
        End If
    Next archivo
    
    ' Recorrer subcarpetas
    For Each subcarp In carpeta.SubFolders
        Call RecorrerCarpeta(subcarp, categoria, subcarp.Name)
    Next subcarp

End Sub


'Public Sub GenerarIndiceInterpretaciones()
'
'    Dim db As DAO.Database
'    Dim rs As DAO.Recordset
'    Dim fso As Object
'    Dim carpetaBase As String
'    Dim carpeta As Object
'    Dim archivo As Object
'    Dim categoria As String
'    Dim valor As String
'    Dim esKarmico As Boolean
'    Dim esMaestro As Boolean
'
'    'carpetaBase = CurrentProject.Path & "\interpretaciones"
'    carpetaBase = "N:\Numerologia\Interpretaciones"
'
'    Set db = CurrentDb
'    Set rs = db.OpenRecordset("tbmIndiceInterpretaciones", dbOpenDynaset)
'    Set fso = CreateObject("Scripting.FileSystemObject")
'
'    ' Limpiar tabla
'    db.Execute "DELETE FROM tbmIndiceInterpretaciones"
'
'    ' Recorrer carpetas
'    For Each carpeta In fso.GetFolder(carpetaBase).SubFolders
'
'        categoria = carpeta.Name
'
'        For Each archivo In carpeta.Files
'            DoEvents
'
'            If LCase(fso.GetExtensionName(archivo.Name)) = "md" Then
'
'                valor = ExtraerValorDesdeNombre(archivo.Name)
'                esKarmico = (InStr(valor, "-") > 0)
'                esMaestro = (valor = "11" Or valor = "22" Or valor = "33" Or valor = "44")
'
'                rs.AddNew
'                rs!IDIndice = AutoNext("IDIndice", "tbmIndiceInterpretaciones")
'                rs!categoria = categoria
'                rs!valor = valor
'                rs!Fichero = archivo.Name
'                rs!Ruta = carpeta.Path
'                rs!esKarmico = esKarmico
'                rs!esMaestro = esMaestro
'                rs.Update
'
'            End If
'
'        Next archivo
'
'    Next carpeta
'
'    MsgBox "Índice generado correctamente.", vbInformation
'
'End Sub

Public Function ExtraerValorDesdeNombre(ByVal Nombre As String) As String

    Dim base As String
    base = Replace(Nombre, ".md", "")
    
    Dim partes() As String
    partes = Split(base, "_")
    
    ' El valor siempre está en la última parte
    ExtraerValorDesdeNombre = partes(UBound(partes))

End Function

