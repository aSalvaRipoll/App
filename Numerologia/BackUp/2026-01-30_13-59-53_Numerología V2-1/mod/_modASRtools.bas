Attribute VB_Name = "_modASRtools"

Option Compare Database
Option Explicit

Public Enum ShellWinMode
    wmHidden '0 Hide the window and activate another window.
    wmActive '1 Activate and display the window. (restore size and position) Specify this flag when displaying a window for the first time.
    wmMinimized '2 Activate & minimize.
    wmMaximized '3 Activate & maximize.
    wmRestore '4 Restore. The active window remains active.
    wmRestoreActive '5 Activate & Restore.
    wmActiveNext '6 Minimize & activate the next top-level window in the Z order.
    wmMinimizeActive '7 Minimize. The active window remains active.
    wmCurrent '8 Display the window in its current state. The active window remains active.
    wmRestoreMinimized '9 Restore & Activate. Specify this flag when restoring a minimized window.
    wmShowState '10 Sets the show-state based on the state of the program that started the application.
End Enum

''Declare PtrSafe Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As LongPtr
'Private Declare PtrSafe Function FindExecutableA Lib "shell32.dll" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long

#If False Then
    Dim wmHidden, wmActive, wmMinimized, wmMaximized, wmRestore, wmRestoreActive, wmActiveNext, wmMinimizeActive, wmCurrent, wmRestoreMinimized, wmShowState
#End If


Sub Exporta()
    
    Dim ruta As String
    Dim RutaSch As String
    Dim RutaDfn As String
    Dim RutaBase As String
    Dim RutaZip As String
    Dim strFName As String
    Dim fso As Object
    Dim boExport As Boolean
    Dim fDate As String
    Dim objOrderInfo As AdditionalData
'    Dim objOrderDetailsInfo As AdditionalData
    Dim tdf As Variant
    Dim qdf As Variant
    Dim obj As Variant
    Dim boGetTbl As Boolean
    Dim boGetDfn As Boolean
    Dim boGetSch As Boolean
    
    boGetTbl = True
    boGetDfn = True
    boGetSch = True
    
    Set fso = CreateObject("Scripting.FileSystemObject")
        
    RutaZip = CurrentProject.Path & "\BackUp"
    PrepRuta RutaZip
    'RutaBase = RutaZip & "\" & fso.GetBaseName(CurrentDb.Name)
    RutaBase = RutaZip & "\" & Format(Now, "yyyy-mm-dd_hh-nn-ss") & "_" & fso.GetBaseName(CurrentDb.Name)
    PrepRuta RutaBase
    
    Call AddCustomProperty("Build", dbInteger, GetProperty("Build", 0) + 1)
    
    fDate = "" '"_" & Format(Date, "yyyymmdd")
    
    ' Inicializamos los diccionarios
    Call InicializarDependencias
    
    If boGetTbl Then
        ruta = RutaBase & "\tdf"
        PrepRuta ruta
        'RutaSch = ruta & "\sch"
        'RutaDfn = ruta & "\dfn"
            
        'If boGetDfn Then _
            PrepRuta RutaDfn
        'If boGetSch Then _
            PrepRuta RutaSch
        
        For Each tdf In CurrentData.AllTables
            DoEvents
            Debug.Print tdf.Name;
            If tdf.Attributes And dbSystemObject Then
                Debug.Print " Sistema --> No se copia"
            ElseIf tdf.Attributes And dbAttachedTable Or tdf.Attributes = 2097152 Then
                Debug.Print " Linked"
                Open ruta & "\Linked.txt" For Append As 1
                Print #1, tdf.Name; " | "; CurrentDb.TableDefs(tdf.Name).Connect
                Close (1)
            ElseIf tdf.Attributes = 0 Or tdf.Attributes = 8 Then
                Debug.Print " Local"
                
                ' Export the contents of the Customers table. The Orders and Order
                ' Details tables will be included in the XML file.
                'Application.ExportXML ObjectType:=acExportTable, DataSource:=tdf.Name, _
                                      DataTarget:=Ruta & "\" & tdf.Name & ".dfn", _
                                      SchemaTarget:=Ruta & "\" & tdf.Name & ".sch" ', _
                                      AdditionalData:=objOrderInfo
                If boGetDfn Then
                    'Application.ExportXML ObjectType:=acExportTable, DataSource:=tdf.Name, _
                                      DataTarget:=RutaDfn & "\" & tdf.Name & ".xml"
                    Application.ExportXML ObjectType:=acExportTable, DataSource:=tdf.Name, _
                                      DataTarget:=ruta & "\" & tdf.Name & ".xml"
                End If

                If boGetSch Then
                    'Application.ExportXML ObjectType:=acExportTable, DataSource:=tdf.Name, _
                                      DataTarget:=RutaDfn & "\" & tdf.Name & ".xml"
                    Call Application.ExportXML(ObjectType:=acExportTable, DataSource:=tdf.Name, _
                                      SchemaTarget:=ruta & "\" & tdf.Name & ".xsd")
                End If
                
                Call RegistrarObjeto(tdf.Name, "Tabla")
                
            Else
                Debug.Print " ???? "; tdf.Attributes
            End If
            'SaveAsText acTable, tdf.Name, Ruta & "\" & tdf.Name & ".txt"
            'SaveAsText acTableDataMacro, tdf.Name, Ruta & "\" & tdf.Name & "_DataMacros.txt"
        Next
    End If
    
'    ruta = RutaBase & "\qdf"
'    PrepRuta ruta
'    For Each qdf In CurrentDb.QueryDefs
'        DoEvents
'        If Left(qdf.Name, 1) <> "~" Then
'            Debug.Print qdf.Name
'            SaveAsText acQuery, qdf.Name, fso.BuildPath(ruta, qdf.Name & fDate & ".txt")
'        End If
'    Next
    
    ruta = RutaBase & "\sql"
    PrepRuta ruta
    For Each qdf In CurrentDb.QueryDefs
        DoEvents
        If Left(qdf.Name, 1) <> "~" Then
            Debug.Print qdf.Name
            
            Open ruta & "\" & qdf.Name & fDate & ".sql" For Output As #1
            Print #1, qdf.sql
            Close
            Call RegistrarObjeto(qdf.Name & fDate & ".sql", "Consulta")
        End If
    Next
    
    ruta = RutaBase & "\frm"
    PrepRuta ruta
    For Each obj In CurrentProject.AllForms
        DoEvents
        Debug.Print obj.Name
        
        SaveAsText acForm, obj.Name, fso.BuildPath(ruta, obj.Name & fDate & ".txt")
        
        ' --- NUEVO: detectar dependencias ---
        Call RegistrarDependencias(obj.Name & fDate & ".txt", ruta)
        Call RegistrarObjeto(obj.Name & fDate & ".txt", "Formulario")
        
    Next
    
    ruta = RutaBase & "\rpt"
    PrepRuta ruta
    For Each obj In CurrentProject.AllReports
        DoEvents
        Debug.Print obj.Name
        
        SaveAsText acReport, obj.Name, fso.BuildPath(ruta, obj.Name & fDate & ".txt")
        
        ' --- NUEVO: detectar dependencias ---
        Call RegistrarDependencias(obj.Name & fDate & ".txt", ruta)
        Call RegistrarObjeto(obj.Name & fDate & ".txt", "Informe")
        
    Next
    
    ruta = RutaBase & "\scr"
    PrepRuta ruta
    For Each obj In CurrentProject.AllMacros
        DoEvents
        Debug.Print obj.Name
        SaveAsText acMacro, obj.Name, fso.BuildPath(ruta, obj.Name & fDate & ".txt")
        
        ' --- NUEVO: detectar dependencias ---
        Call RegistrarDependencias(obj.Name & fDate & ".txt", ruta)
        Call RegistrarObjeto(obj.Name & fDate & ".txt", "Macro")
    Next
    
    ruta = RutaBase & "\mod"
    PrepRuta ruta
    For Each obj In VBE.ActiveVBProject.VBComponents
        boExport = True
        strFName = obj.Name & fDate

        ''' Concatenate the correct filename for export.
        Select Case obj.Type
            Case 1 'vbext_ct_StdModule
                strFName = strFName & ".bas"
                Call RegistrarObjeto(strFName, "Módulo")
            Case 2 'vbext_ct_ClassModule
                strFName = strFName & ".cls"
                Call RegistrarObjeto(strFName, "Clase")
            Case 3 'vbext_ct_MSForm
                strFName = strFName & ".frm"
                Call RegistrarObjeto(strFName, "User form")
            Case 11 'vbext_ct_ActiveXDesigner
                strFName = strFName & ".dsg"
                Call RegistrarObjeto(strFName, "Designer")
            Case 100 'vbext_ct_Document
                'Ya se exporta con el form
                'strFName = strFName & ".cls"
                boExport = False
        End Select
        Debug.Print strFName
        If boExport Then
            obj.Export fso.BuildPath(ruta, strFName)
            
            ' --- NUEVO: detectar dependencias ---
            Call RegistrarDependencias(strFName, ruta)
            'Call RegistrarObjeto(strFName, "Módulo")
        End If
    Next

    ' Agregamos referencias, dependencias y metadatos
    ' --- NUEVO: generar el .ref con todo ---
    ExportarReferencias RutaBase
    
    MsgBox "Fin"
    
    Shell "explorer " & RutaBase
    
    Set fso = Nothing
End Sub

'========================================================================================================
' CreaPropiedades

Sub AgregaProps()
    Dim x
    
    BorrarPropiedad "vVersion"
    
    BorrarPropiedad "vMajor"
    BorrarPropiedad "vMinor"
    BorrarPropiedad "vBuild"
    
    
    
    x = AddCustomProperty("Major", dbInteger, 2)
    x = AddCustomProperty("Minor", dbInteger, 1)
    x = AddCustomProperty("Revision", dbInteger, 0)
    x = AddCustomProperty("Build", dbInteger, 5)

End Sub



Function AddCustomProperty(strName As String, _
                           varType As Variant, _
                           varValue As Variant) As Boolean
    ' The following generic object variables are required
    ' when there is no reference to the DAO 3.6 object library.
    Dim objDatabase As Object
    Dim objProperty As Object

    Const PROP_NOT_FOUND_ERROR = 3270

    Set objDatabase = CurrentDb
    On Error GoTo AddProp_Err
    objDatabase.Properties(strName) = varValue

    AddCustomProperty = True

AddProp_End:
    Exit Function

AddProp_Err:
    If Err = PROP_NOT_FOUND_ERROR Then
        Set objProperty = objDatabase.CreateProperty(strName, varType, varValue)
        objDatabase.Properties.Append objProperty
        Resume
    Else
        AddCustomProperty = False
        Resume AddProp_End
    End If
End Function


Sub BorrarPropiedad(strNombre)

    On Error Resume Next
    CurrentDb.Properties.Delete strNombre

End Sub


'========================================================================================================


Sub ImportaDesdeDialogo()
    Dim RutaBase As String
    
    RutaBase = SeleccionaCarpeta("Selecciona la carpeta base del backup")
    
    If RutaBase = "" Then
        MsgBox "Operación cancelada"
        Exit Sub
    End If
    
    Importa RutaBase
End Sub


Sub Importa(RutaBase As String)

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim Carpeta As Object, Archivo As Object
    Dim ruta As String, Nombre As String, Ext As String
    
    '----------------------------------------------------------
    ' 1. TABLAS (XML)
    '----------------------------------------------------------
    ruta = fso.BuildPath(RutaBase, "tdf")
    If fso.FolderExists(ruta) Then
        
        'Esquemas
        If fso.FolderExists(fso.BuildPath(ruta, "sch")) Then
            For Each Archivo In fso.GetFolder(fso.BuildPath(ruta, "sch")).Files
                If LCase(fso.GetExtensionName(Archivo.Name)) = "xml" Then
                    Nombre = fso.GetBaseName(Archivo.Name)
                    Application.ImportXML Archivo.Path, acStructureOnly
                End If
            Next
        End If
        
        'Datos
        If fso.FolderExists(fso.BuildPath(ruta, "dfn")) Then
            For Each Archivo In fso.GetFolder(fso.BuildPath(ruta, "dfn")).Files
                If LCase(fso.GetExtensionName(Archivo.Name)) = "xml" Then
                    Nombre = fso.GetBaseName(Archivo.Name)
                    Application.ImportXML Archivo.Path, acAppendData
                End If
            Next
        End If
        
    End If
    
    '----------------------------------------------------------
    ' 2. CONSULTAS
    '----------------------------------------------------------
    ruta = fso.BuildPath(RutaBase, "qdf")
    If fso.FolderExists(ruta) Then
        For Each Archivo In fso.GetFolder(ruta).Files
            If LCase(fso.GetExtensionName(Archivo.Name)) = "txt" Then
                Nombre = fso.GetBaseName(Archivo.Name)
                LoadFromText acQuery, Nombre, Archivo.Path
            End If
        Next
    End If
    
    ruta = fso.BuildPath(RutaBase, "sql")
    If fso.FolderExists(ruta) Then
        Dim qdf As DAO.QueryDef
        For Each Archivo In fso.GetFolder(ruta).Files
            If LCase(fso.GetExtensionName(Archivo.Name)) = "sql" Then
                Nombre = fso.GetBaseName(Archivo.Name)
                Set qdf = CurrentDb.CreateQueryDef(Nombre, Archivo.OpenAsTextStream.ReadAll)
            End If
        Next
    End If
    
    '----------------------------------------------------------
    ' 3. FORMULARIOS
    '----------------------------------------------------------
    ruta = fso.BuildPath(RutaBase, "frm")
    If fso.FolderExists(ruta) Then
        For Each Archivo In fso.GetFolder(ruta).Files
            If LCase(fso.GetExtensionName(Archivo.Name)) = "txt" Then
                Nombre = fso.GetBaseName(Archivo.Name)
                LoadFromText acForm, Nombre, Archivo.Path
            End If
        Next
    End If
    
    '----------------------------------------------------------
    ' 4. INFORMES
    '----------------------------------------------------------
    ruta = fso.BuildPath(RutaBase, "rpt")
    If fso.FolderExists(ruta) Then
        For Each Archivo In fso.GetFolder(ruta).Files
            If LCase(fso.GetExtensionName(Archivo.Name)) = "txt" Then
                Nombre = fso.GetBaseName(Archivo.Name)
                LoadFromText acReport, Nombre, Archivo.Path
            End If
        Next
    End If
    
    '----------------------------------------------------------
    ' 5. MACROS
    '----------------------------------------------------------
    ruta = fso.BuildPath(RutaBase, "scr")
    If fso.FolderExists(ruta) Then
        For Each Archivo In fso.GetFolder(ruta).Files
            If LCase(fso.GetExtensionName(Archivo.Name)) = "txt" Then
                Nombre = fso.GetBaseName(Archivo.Name)
                LoadFromText acMacro, Nombre, Archivo.Path
            End If
        Next
    End If
    
    '----------------------------------------------------------
    ' 6. MÓDULOS VBA
    '----------------------------------------------------------
    ruta = fso.BuildPath(RutaBase, "mod")
    If fso.FolderExists(ruta) Then
        For Each Archivo In fso.GetFolder(ruta).Files
            Ext = LCase(fso.GetExtensionName(Archivo.Name))
            If Ext = "bas" Or Ext = "cls" Or Ext = "frm" Or Ext = "dsg" Then
                VBE.ActiveVBProject.VBComponents.Import Archivo.Path
            End If
        Next
    End If
    
    MsgBox "Importación completada"
    
End Sub

Function SeleccionaCarpeta(Optional Titulo As String = "Selecciona carpeta") As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    
    With fd
        .Title = Titulo
        .AllowMultiSelect = False
        If .Show = -1 Then
            SeleccionaCarpeta = .SelectedItems(1)
        Else
            SeleccionaCarpeta = ""
        End If
    End With
End Function


'========================================================================================================
Sub C_Zip2(ZipPath As String, ZipName As String, ByVal ruta As String, Optional Modo As ShellWinMode = wmCurrent)
    
    Const PATH_TO_7Z = "D:\USR\LOCAL\7Zip\7z.exe"
    Dim cmdLine As String
    Dim DestPath As String, Nombre As String
    
    'DestPath = fso.GetParentFolderName(CurrentDb.Name) & "\BackUp"
 
    DestPath = ZipPath
    Nombre = ZipName & "_" & Format(Date, "yyyymmdd") & ".zip"
    
    cmdLine = PATH_TO_7Z & " a -tzip -x!*.zip """ & DestPath & "\" & Nombre & """ """ & ruta & """"
    Debug.Print cmdLine
    
    Call RunShell(cmdLine, Modo)

End Sub

Sub PrepRuta(strPath As String)

    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(strPath) Then fso.CreateFolder strPath
    Set fso = Nothing
    
End Sub

Sub C_ZipMe(ByVal Nombre As String, ByVal Carpeta As String, Optional Modo As ShellWinMode = wmCurrent)
 
    Const PATH_TO_7Z = "D:\USR\LOCAL\7Zip\7z.exe"
 
    Call RunShell(PATH_TO_7Z & " a -tzip -x!" & CurrentProject.Path & Carpeta & "\*.zip """ & Nombre & """ """ & Carpeta & """", Modo)
    
End Sub

Sub RunShell(ByRef applPath As String, Optional Modo As ShellWinMode = wmActive)

    Dim WshShell As Object
    Dim ErrorCode As Integer
    Dim ShellCmd As String
    
    ShellCmd = applPath

    Set WshShell = CreateObject("WScript.Shell")
    ErrorCode = WshShell.Run(ShellCmd, Modo, True)

    Set WshShell = Nothing

End Sub

'Function GetExecutable(strFile As String) As String
'
'    Dim strPath As String
'
'    Dim intLen As Integer
'
'    strPath = Space(255)
'
'    intLen = FindExecutableA(strFile, "\", strPath)
'
'    GetExecutable = Trim(strPath)
'
'End Function


'Function Obtener_Path_Access(UnaRutaBd As String) As String
'
'  Dim I     As LongPtr
'  Dim S2    As String
'  Dim Path  As String
'
'    Const SYS_OUT_OF_MEM        As Long = &H0
'    Const ERROR_FILE_NOT_FOUND  As Long = &H2
'    Const ERROR_PATH_NOT_FOUND  As Long = &H3
'    Const ERROR_BAD_FORMAT      As Long = &HB
'    Const NO_ASSOC_FILE         As Long = &H1F
'    Const MIN_SUCCESS_LNG       As Long = &H20
'    Const MAX_PATH              As Long = &H104
'
'    Const USR_NULL              As String = "NULL"
'    Const S_DIR                 As String = "C:\" '// Change as required (drive that .exe will be on)
'
'  S2 = String(MAX_PATH, Chr(32)) & Chr$(0)
'
'  I = FindExecutable(UnaRutaBd & Chr$(0), vbNullString, S2)
'
'  If I > MIN_SUCCESS_LNG Then
'    Path = Left$(S2, InStr(S2, Chr$(0)) - 1)
'    'If Mid(Path, InStrRev(Path, "\") + 1) = "MSACCESS.EXE" Then
'        Obtener_Path_Access = Path
'    'Else
'    '    Obtener_Path_Access = ""
'    'End If
'  Else
'        Obtener_Path_Access = ""
'  End If
'End Function


