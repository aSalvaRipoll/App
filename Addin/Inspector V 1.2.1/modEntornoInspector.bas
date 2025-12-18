Attribute VB_Name = "modEntornoInspector"
Option Compare Database
Option Explicit

'=====================================================
' Módulo: modEntornoInspector
' Detección robusta de Add-Ins ACCDA instalados en Access
'=====================================================

Private Const intForReading As Long = 1
Private Const intUnicode As Long = -1

Private Const NOMBRE_ADDIN As String = "InspectorVBA.accda"

Private objFSO As Object
Public colAddins As Collection   ' Colección de clsAddin

'-----------------------------------------------------
' Función pública: ¿Está instalado el Inspector?
'-----------------------------------------------------
Public Function InspectorInstalado() As Boolean
    Dim ai As clsAddin

    ' Asegurar que la colección está cargada
    If colAddins Is Nothing Then
        ListaComplementosAccess
    End If

    ' Buscar el Add-In en la colección
    For Each ai In colAddins
        If LCase$(ai.addin_Name) = LCase$(NOMBRE_ADDIN) Then
            InspectorInstalado = True
            Exit Function
        End If
    Next ai

    InspectorInstalado = False
End Function

'-----------------------------------------------------
' Entrada principal: carga colAddins desde el registro
'-----------------------------------------------------
Public Sub ListaComplementosAccess()
    ' Inicializar colección y FSO
    Set colAddins = New Collection
    Set objFSO = CreateObject("Scripting.FileSystemObject")

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

    strRegPath = "HKEY_LOCAL_MACHINE\Software\Microsoft\Office"
    strRawFile = Environ$("TEMP") & "\OfficeRaw.reg"
    strTempFile = Environ$("TEMP") & "\OfficeExport.reg"

    ' Exportar registro (64 bits)
    strCommand = "cmd /c REG EXPORT """ & strRegPath & """ """ & strRawFile & """ /reg:64"
    objShell.Run strCommand, 0, True

    DoEvents

    ' Crear archivo final limpio
    Set objRegFile = objFSO.CreateTextFile(strTempFile, True, True)
    objRegFile.WriteLine "Windows Registry Editor Version 5.00"

    If objFSO.FileExists(strRawFile) Then
        Set objInputFile = objFSO.OpenTextFile(strRawFile, intForReading, False, intUnicode)

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
End Sub

'-----------------------------------------------------
' Analiza el archivo exportado y construye colAddins
'-----------------------------------------------------
Private Sub AnalizaRegistro(strFileName As String)
    Dim objAddin As clsAddin
    Dim objInputFile As Object
    Dim strLine As String
    Dim salir As Boolean

    Set objInputFile = objFSO.OpenTextFile(strFileName, intForReading, False, intUnicode)

    While Not objInputFile.AtEndOfStream
        strLine = objInputFile.ReadLine

        ' Buscar claves de Add-Ins de Access
        If InStr(strLine, "Access\Menu Add-Ins\") Then

            ' Extraer nombre del Add-In
            strLine = Mid$(strLine, InStrRev(strLine, "\") + 1)
            strLine = Left$(strLine, Len(strLine) - 1)

            Set objAddin = New clsAddin
            objAddin.addin_Name = strLine

            salir = False

            ' Buscar la propiedad Library
            While Not salir And Not objInputFile.AtEndOfStream
                strLine = objInputFile.ReadLine

                If InStr(strLine, "Library") Then
                    strLine = Mid$(strLine, InStr(strLine, "=") + 1)
                    strLine = Replace$(strLine, Chr$(34), "")
                    strLine = Replace$(strLine, "\\", "\")
                    objAddin.library = strLine
                    salir = True
                End If
            Wend

            ' Solo añadir si el archivo existe
            If objFSO.FileExists(objAddin.library) Then

                ' ? NUEVO: inicializar propiedades derivadas
                objAddin.Initialize

                ' Añadir a la colección
                colAddins.Add objAddin, objAddin.addin_Name
            End If
        End If
    Wend

    objInputFile.Close
End Sub

'-----------------------------------------------------
' Devuelve la ruta del Add-In Inspector
'-----------------------------------------------------
Public Function RutaAddinInspector() As String
    Dim ai As clsAddin

    If colAddins Is Nothing Then
        ListaComplementosAccess
    End If

    For Each ai In colAddins
        If LCase$(ai.addin_Name) = LCase$(NOMBRE_ADDIN) Then
            RutaAddinInspector = ai.library
            Exit Function
        End If
    Next ai
End Function

