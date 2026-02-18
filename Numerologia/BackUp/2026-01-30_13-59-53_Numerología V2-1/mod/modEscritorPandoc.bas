Attribute VB_Name = "modEscritorPandoc"
' ------------------------------------------------------
' Nombre:    modEscritorPandoc
' Tipo:      Módulo
' Propósito:
' Autor:     asalv
' Fecha:     15/01/2026
' ------------------------------------------------------

Option Compare Database
Option Explicit

' ============================================================
' CONFIGURACIÓN
' ============================================================

' Usamos solo "pandoc" porque ya está en el PATH
Private Const PANDOC_PATH As String = "pandoc"


' ============================================================
' Ejecutar un comando y esperar a que termine
' ============================================================

Private Sub RunCommand(cmd As String)
    CreateObject("WScript.Shell").Run cmd, 0, True
End Sub


' ============================================================
' Comprobar si Pandoc está disponible en el sistema
' ============================================================

Public Function PandocDisponible() As Boolean
    On Error GoTo ErrHandler

    ' Ejecutamos "pandoc -v" y esperamos
    RunCommand PANDOC_PATH & " -v"

    PandocDisponible = True
    Exit Function

ErrHandler:
    PandocDisponible = False
End Function


' ============================================================
' Convertir Markdown ? HTML usando Pandoc
' ============================================================

Public Function PandocMdToHtml(mdPath As String, htmlPath As String) As Boolean
    If Not PandocDisponible() Then
        MsgBox "Pandoc no está instalado o no está en el PATH.", vbCritical
        Exit Function
    End If

    On Error GoTo ErrHandler
    Dim cmd As String

    cmd = PANDOC_PATH & _
          " """ & mdPath & """" & _
          " -f markdown -t html5 -s -o """ & htmlPath & """"

    RunCommand cmd
    PandocMdToHtml = True
    Exit Function

ErrHandler:
    PandocMdToHtml = False
End Function


' ============================================================
' Convertir Markdown ? DOCX usando Pandoc
' ============================================================

Public Function PandocMdToDocx(mdPath As String, docxPath As String) As Boolean
    If Not PandocDisponible() Then
        MsgBox "Pandoc no está instalado o no está en el PATH.", vbCritical
        Exit Function
    End If

    On Error GoTo ErrHandler
    Dim cmd As String

    cmd = PANDOC_PATH & _
          " """ & mdPath & """" & _
          " -f markdown -t docx -s -o """ & docxPath & """"

    RunCommand cmd
    PandocMdToDocx = True
    Exit Function

ErrHandler:
    PandocMdToDocx = False
End Function


' ============================================================
' Abrir DOCX en Word y exportar a PDF
' ============================================================

Public Function WordDocxToPdf(docxPath As String, pdfPath As String) As Boolean
    On Error GoTo ErrHandler
    Dim wd As Object, doc As Object

    Set wd = CreateObject("Word.Application")
    wd.Visible = False

    Set doc = wd.Documents.Open(docxPath)
    doc.ExportAsFixedFormat pdfPath, 17   ' 17 = PDF

    doc.Close False
    wd.Quit

    WordDocxToPdf = True
    Exit Function

ErrHandler:
    WordDocxToPdf = False
End Function


' ============================================================
' Flujo completo: Markdown ? DOCX ? PDF
' ============================================================

Public Function ConvertMarkdownToPdf(mdPath As String, pdfPath As String) As Boolean
    Dim tempDocx As String

    If Not PandocDisponible() Then
        MsgBox "Pandoc no está instalado o no está en el PATH.", vbCritical
        Exit Function
    End If

    tempDocx = Environ$("TEMP") & "\pandoc_temp.docx"

    If Not PandocMdToDocx(mdPath, tempDocx) Then Exit Function
    If Not WordDocxToPdf(tempDocx, pdfPath) Then Exit Function

    ConvertMarkdownToPdf = True
End Function


