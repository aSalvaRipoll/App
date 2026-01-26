Attribute VB_Name = "modIconos_Blindado"
' ------------------------------------------------------
' Nombre:    modIconos_Blindado
' Tipo:      Módulo
' Propósito:
' Autor:     asalv
' Fecha:     15/01/2026
' ------------------------------------------------------

Option Compare Database
Option Explicit


'********************************************************************************
'<< File          : Numerología
'<< Version       : 1.0
'<< Creation date : 20/05/2020
'<< Revision date : 20/05/2020
'<< Author        : Alba Salvá (Isis)
'<< Projects      : ImgMSO
'<< Description   : Poner imágenes MSO en Access.
'********************************************************************************

'********************************************************************************
'<< HISTORIAL DE CAMBIOS
'********************************************************************************
'<< 1.0         Alba Salvá (Isis)
'<< 20/05/2020  Versión inicial
'********************************************************************************

'********************************************************************************
'<< CONSTANTES
'********************************************************************************
'<< Nombre del módulo
Private Const MSV_NOMBRE_MODULO As String = "modGraficos"

' ============================================================
'  modIconos_Blindado
'  Conversión segura de ImageMso ? archivo .ico
'  Compatible con Access 2010–365
' ============================================================

' ------------------------------------------------------------
' 1. Resolver icono (archivo o ImageMso)
' ------------------------------------------------------------
Public Function ResolveIconFile_Blindado(Icon As String, TempName As String) As String
    Dim BmpPath As String
    Dim IcoPath As String

    If Len(Icon) = 0 Then
        ResolveIconFile_Blindado = vbNullString
        Exit Function
    End If

    ' Caso 1: ImageMso
    If Left$(Icon, 9) = "ImageMso:" Then
'        BmpPath = Environ$("TEMP") & "\" & TempName & ".bmp"
'        IcoPath = Environ$("TEMP") & "\" & TempName & ".ico"

        BmpPath = CurrentProject.Path & "\" & TempName & ".bmp"
        IcoPath = CurrentProject.Path & "\" & TempName & ".ico"

        If SaveImageMsoAsBmp_Blindado(Mid$(Icon, 10), BmpPath) Then
            If ConvertBmpToIco_Blindado(BmpPath, IcoPath) Then
                ResolveIconFile_Blindado = IcoPath
                Exit Function
            End If
        End If

        ResolveIconFile_Blindado = vbNullString
        Exit Function
    End If

    ' Caso 2: archivo .ico existente
    If Len(Dir$(Icon)) > 0 Then
        ResolveIconFile_Blindado = Icon
        Exit Function
    End If

    ResolveIconFile_Blindado = vbNullString
End Function



' ------------------------------------------------------------
' 2. Guardar ImageMso como BMP usando tu rutina
' ------------------------------------------------------------
Public Function SaveImageMsoAsBmp_Blindado(ImageMsoName As String, OutputBmp As String) As Boolean
    On Error GoTo ErrHandler

    Dim img As Object

    If GetPictureFromMso(ImageMsoName, img, 16) Then
        If Not img Is Nothing Then
            SavePicture img, OutputBmp
            SaveImageMsoAsBmp_Blindado = True
        End If
    End If

    Exit Function

ErrHandler:
    MsgBox "Error en SaveImageMsoAsBmp_Blindado:" & vbCrLf & _
           Err.Number & " - " & Err.Description, vbCritical
    SaveImageMsoAsBmp_Blindado = False
End Function

'-------------------------------------------------------------------------
' Imagen de un Mso (formato de IPictureDisp)
'-------------------------------------------------------------------------
Public Function GetPictureFromMso(pMso As String, PImage As Object, Optional pSize As Long = 32) As Boolean
'---------------------------------------------------------------------------------------
' Procedimiento  : GetPictureFromMso
' Tipo           : Function
' Fecha / Hora   : 29/05/2020 11:38
' Autor          : Alba Salvá (Isis)
' Retorno        : Boolean
' Propósito      :
'---------------------------------------------------------------------------------------
'

    ' Creación del menú para los elementos
    Dim c As Object 'Office.CommandBars
    Dim o As Object
    On Error GoTo ErrorTrap
    #If Access2000 = False Then
        Set c = CurrentProject.Application.CommandBars
        Set o = c.GetImageMso(pMso, pSize, pSize)
    #End If
    If Not o Is Nothing Then
        Set PImage = o
        GetPictureFromMso = True
    Else
        GetPictureFromMso = False
        Set PImage = Nothing
    End If
    Exit Function
ErrorTrap:
    GetPictureFromMso = False
    Set PImage = Nothing
End Function



' ------------------------------------------------------------
' 3. Convertir BMP ? ICO
'    (versión simple y estable)
' ------------------------------------------------------------
Public Function ConvertBmpToIco_Blindado(BmpPath As String, IcoPath As String) As Boolean
    On Error GoTo ErrHandler

    Dim pic As StdPicture
    Set pic = LoadPicture(BmpPath)

    SavePicture pic, IcoPath

    ConvertBmpToIco_Blindado = True
    Exit Function

ErrHandler:
    MsgBox "Error en ConvertBmpToIco_Blindado:" & vbCrLf & _
           Err.Number & " - " & Err.Description, vbCritical
    ConvertBmpToIco_Blindado = False
End Function


