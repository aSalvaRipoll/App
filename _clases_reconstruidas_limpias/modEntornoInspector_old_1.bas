Attribute VB_Name = "modEntornoInspector_old_1"

Option Compare Database
Option Explicit

'=====================================================
' Módulo: modEntornoInspector
' Detección de Add-Ins ACCDA instalados en Access
'=====================================================

Private Const intForReading As Long = 1
Private Const intUnicode As Long = -1

Private objFSO As Object
Public colAddins As Collection   ' Colección de clsAddin

'-----------------------------------------------------
' Función pública: ¿Está instalado el Inspector?
'-----------------------------------------------------
Public Function InspectorInstalado() As Boolean
    Dim ai As clsAddin
    Dim nombre As String

    nombre = "InspectorVBA.accda"   ' Ajusta si usas otro nombre

    ' Asegurar que la colección está cargada
    If colAddins Is Nothing Then
        ListaComplementosAccess
    End If

    ' Buscar el AddIn en la colección
    For Each ai In colAddins
        If LCase$(ai.addin_Name) = LCase$(nombre) Then
            InspectorInstalado = True
            Exit Function
        End If
    Next ai
End Function

'-----------------------------------------------------
' Entrada principal: carga colAddins desde el registro
'-----------------------------------------------------

'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : ListaComplementosAccess
' Autor original    : Alba Salvá
' Creado            : 21/02/2023
' Adaptado por      : Alba Salvá
' Propósito         : listar todos los complementos de Access en un listbox de un formulario
'-----------------------------------------------------------------------------------------------------------------------------------------------
Public Sub ListaComplementosAccess()
    ' Inicializar colección
    Set colAddins = Nothing

    ' Exportar registro y analizarlo
    ExportaRegistro
    
End Sub

'-----------------------------------------------------
' Exporta la rama Office del registro a un archivo temporal
'-----------------------------------------------------
Private Sub ExportaRegistro()
    
    Dim objShell As Object
    Dim strRegPath As String
    Dim strTempFile As String
    Dim strRawFile As String
    Dim strCommand As String
    Dim objRegFile As Object
    Dim objInputFile As Object

    Set objShell = CreateObject("WScript.Shell")
    Set objFSO = CreateObject("Scripting.FileSystemObject")

    strRegPath = "HKEY_LOCAL_MACHINE\Software\Microsoft\Office"
    strRawFile = Replace(strRegPath, "\", "_") & ".reg"
    strTempFile = Environ$("TEMP") & "\OfficeExport.reg"

    ' Exportar registro (64 bits)
    strCommand = "cmd /c REG EXPORT """ & strRegPath & """ """ & strRawFile & """ /reg:64"
    objShell.Run strCommand, 0, True

    ' DoEvents es necesario para permitir que Windows termine de escribir el archivo
    DoEvents

    ' Crear archivo final limpio
    Set objRegFile = objFSO.CreateTextFile(strTempFile, True, True)
    objRegFile.WriteLine "Windows Registry Editor Version 5.00"

    If objFSO.FileExists(strRawFile) Then
        Set objInputFile = objFSO.OpenTextFile(strRawFile, intForReading, False, intUnicode)

        ' Saltar la primera línea del archivo exportado
        If Not objInputFile.AtEndOfStream Then
            objInputFile.SkipLine
            objRegFile.Write objInputFile.ReadAll
        End If

        objInputFile.Close
        objFSO.DeleteFile strRawFile, True
    End If

    objRegFile.Close

    ' Analizar el archivo resultante
    AnalizaRegistro strTempFile

    ' Limpiar archivo temporal
    objFSO.DeleteFile strTempFile, True

    Set objShell = Nothing
    Set objFSO = Nothing
End Sub

'-----------------------------------------------------
' Analiza el archivo exportado y construye colAddins
'-----------------------------------------------------
Private Sub AnalizaRegistro(strFileName As String)
    
    Dim objAddin As clsAddin
    Dim objInputFile As Object
    Dim strLine As String
    Dim salir As Boolean

    Set colAddins = New Collection

    Set objInputFile = objFSO.OpenTextFile(strFileName, intForReading, False, intUnicode)

    While Not objInputFile.AtEndOfStream
        strLine = objInputFile.ReadLine

        ' Buscar claves de AddIns de Access
        If InStr(strLine, "Access\Menu Add-Ins\") Then

            ' Extraer nombre del AddIn
            strLine = Mid(strLine, InStrRev(strLine, "\") + 1)
            strLine = Left(strLine, Len(strLine) - 1)

            Set objAddin = New clsAddin
            objAddin.addin_Name = strLine

            salir = False

            ' Buscar la línea Library
            While Not salir And Not objInputFile.AtEndOfStream
                strLine = objInputFile.ReadLine

                If InStr(strLine, "Library") Then
                    strLine = Mid(strLine, InStr(strLine, "=") + 1)
                    strLine = Replace(strLine, Chr(34), "")
                    strLine = Replace(strLine, "\\", "\")
                    objAddin.library = strLine
                    salir = True
                End If
            Wend

            colAddins.Add objAddin, objAddin.addin_Name
        End If
    Wend

    objInputFile.Close
    
End Sub

