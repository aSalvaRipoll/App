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
    
    Dim Ruta As String
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
    Dim tdf As Variant, qdf As Variant, obj As Variant
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
    Shell "explorer " & RutaBase
    
    fDate = "" '"_" & Format(Date, "yyyymmdd")
    
    If boGetTbl Then
        Ruta = RutaBase & "\tdf"
        PrepRuta Ruta
        RutaSch = Ruta & "\sch"
        RutaDfn = Ruta & "\dfn"
            
        If boGetDfn Then _
            PrepRuta RutaDfn
        If boGetSch Then _
            PrepRuta RutaSch
        
        For Each tdf In CurrentData.AllTables
            DoEvents
            Debug.Print tdf.Name;
            If tdf.Attributes And dbSystemObject Then
                Debug.Print " Sistema --> No se copia"
            ElseIf tdf.Attributes And dbAttachedTable Or tdf.Attributes = 2097152 Then
                Debug.Print " Linked"
                Open Ruta & "\Linked.txt" For Append As 1
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
                    Application.ExportXML ObjectType:=acExportTable, DataSource:=tdf.Name, _
                                      DataTarget:=RutaDfn & "\" & tdf.Name & ".xml"
                End If

                If boGetSch Then
                    Application.ExportXML ObjectType:=acExportTable, DataSource:=tdf.Name, _
                                      SchemaTarget:=RutaSch & "\" & tdf.Name & ".xml" ', _

                End If
            
            Else
                Debug.Print " ???? "; tdf.Attributes
            End If
            'SaveAsText acTable, tdf.Name, Ruta & "\" & tdf.Name & ".txt"
            'SaveAsText acTableDataMacro, tdf.Name, Ruta & "\" & tdf.Name & "_DataMacros.txt"
        Next
    End If
    
    Ruta = RutaBase & "\qdf"
    PrepRuta Ruta
    For Each qdf In CurrentDb.QueryDefs
        DoEvents
        If Left(qdf.Name, 1) <> "~" Then
            Debug.Print qdf.Name
            SaveAsText acQuery, qdf.Name, fso.BuildPath(Ruta, qdf.Name & fDate & ".txt")
        End If
    Next
    Ruta = RutaBase & "\sql"
    PrepRuta Ruta
    For Each qdf In CurrentDb.QueryDefs
        DoEvents
        If Left(qdf.Name, 1) <> "~" Then
            Debug.Print qdf.Name
            Open Ruta & "\" & qdf.Name & fDate & ".sql" For Output As #1
            Print #1, qdf.sql
            Close
        End If
    Next
    
    Ruta = RutaBase & "\frm"
    PrepRuta Ruta
    For Each obj In CurrentProject.AllForms
        DoEvents
        Debug.Print obj.Name
        SaveAsText acForm, obj.Name, fso.BuildPath(Ruta, obj.Name & fDate & ".txt")
    Next
    
    Ruta = RutaBase & "\rpt"
    PrepRuta Ruta
    For Each obj In CurrentProject.AllReports
        DoEvents
        Debug.Print obj.Name
        SaveAsText acReport, obj.Name, fso.BuildPath(Ruta, obj.Name & fDate & ".txt")
    Next
    
    Ruta = RutaBase & "\scr"
    PrepRuta Ruta
    For Each obj In CurrentProject.AllMacros
        DoEvents
        Debug.Print obj.Name
        SaveAsText acMacro, obj.Name, fso.BuildPath(Ruta, obj.Name & fDate & ".txt")
    Next
    
    Ruta = RutaBase & "\mod"
    PrepRuta Ruta
    For Each obj In VBE.ActiveVBProject.VBComponents
        boExport = True
        strFName = obj.Name & fDate

        ''' Concatenate the correct filename for export.
        Select Case obj.Type
            Case 1 'vbext_ct_StdModule
                strFName = strFName & ".bas"
            Case 2 'vbext_ct_ClassModule
                strFName = strFName & ".cls"
            Case 3 'vbext_ct_MSForm
                strFName = strFName & ".frm"
            Case 11 'vbext_ct_ActiveXDesigner
                strFName = strFName & ".dsg"
            Case 100 'vbext_ct_Document
                'Ya se exporta con el form
                'strFName = strFName & ".cls"
                boExport = False
        End Select
        Debug.Print strFName
        If boExport Then
'        obj.Export szExportPath & szFileName
            obj.Export fso.BuildPath(Ruta, strFName)
        End If
    Next

'    For Each obj In CurrentProject.AllModules
'        DoEvents
'        Debug.Print obj.Name
'        SaveAsText acModule, obj.Name, fso.BuildPath(Ruta, obj.Name & fDate & ".txt")
'    Next
    
'    'C_ZipMe RutaZip & "\" & fso.GetBaseName(CurrentDb.Name) & "_" & Format(Date, "yyyymmdd") & ".zip", RutaBase & "\"
'    C_Zip2 RutaZip, fso.GetBaseName(CurrentDb.Name), RutaBase
'
'    'For Each carpeta In fso.GetFolder(RutaBase).SubFolders
'        fso.DeleteFolder RutaBase, True
'    'Next
    
    MsgBox "Fin"
    
    Set fso = Nothing
End Sub
Sub C_Zip2(ZipPath As String, ZipName As String, ByVal Ruta As String, Optional Modo As ShellWinMode = wmCurrent)
    
    Const PATH_TO_7Z = "D:\USR\LOCAL\7Zip\7z.exe"
    Dim cmdLine As String
    Dim DestPath As String, Nombre As String
    
    'DestPath = fso.GetParentFolderName(CurrentDb.Name) & "\BackUp"
 
    DestPath = ZipPath
    Nombre = ZipName & "_" & Format(Date, "yyyymmdd") & ".zip"
    
    cmdLine = PATH_TO_7Z & " a -tzip -x!*.zip """ & DestPath & "\" & Nombre & """ """ & Ruta & """"
    Debug.Print cmdLine
    
    Call RunShell(cmdLine, Modo)

End Sub

Sub PrepRuta(strPath As String)

    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(strPath) Then fso.CreateFolder strPath
    Set fso = Nothing
    
End Sub

Sub C_ZipMe(ByVal Nombre As String, ByVal carpeta As String, Optional Modo As ShellWinMode = wmCurrent)
 
    Const PATH_TO_7Z = "D:\USR\LOCAL\7Zip\7z.exe"
 
    Call RunShell(PATH_TO_7Z & " a -tzip -x!" & CurrentProject.Path & carpeta & "\*.zip """ & Nombre & """ """ & carpeta & """", Modo)
    
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


