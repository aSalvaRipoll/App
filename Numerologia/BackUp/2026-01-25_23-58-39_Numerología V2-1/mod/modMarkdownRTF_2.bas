Attribute VB_Name = "modMarkdownRTF_2"

Option Compare Database
Option Explicit

' ============================================================
'   Convertir Markdown ? RTF (incluye tablas)
' ============================================================

Public Function ConvertirMarkdownRTF(md As String) As String
    Dim rtf As String
    Dim lineas() As String
    Dim linea As Variant
    Dim tmp As String
    
    Dim enTabla As Boolean
    Dim bloqueTabla As String
    
    ' Encabezado RTF básico
    rtf = "{\rtf1\ansi\deff0" & vbCrLf
    
    lineas = Split(md, vbCrLf)
    
    For Each linea In lineas
        
        ' --- Detectar tabla Markdown ---
        If Left(Trim(linea), 1) = "|" Then
            bloqueTabla = bloqueTabla & linea & vbCrLf
            enTabla = True
            GoTo Siguiente
        End If
        
        ' Fin de tabla
        If enTabla And Trim(linea) = "" Then
            rtf = rtf & ConvertirTablaMarkdownRTF(bloqueTabla)
            bloqueTabla = ""
            enTabla = False
            GoTo Siguiente
        End If
        
        ' Si estamos dentro de una tabla, seguir acumulando
        If enTabla Then
            bloqueTabla = bloqueTabla & linea & vbCrLf
            GoTo Siguiente
        End If
        
        ' --- Procesar línea normal ---
        tmp = ProcesarLineaMarkdown(linea)
        rtf = rtf & tmp & "\par" & vbCrLf
        
Siguiente:
    Next linea
    
    ' Si el archivo termina con una tabla
    If enTabla Then
        rtf = rtf & ConvertirTablaMarkdownRTF(bloqueTabla)
    End If
    
    rtf = rtf & "}"
    
    ConvertirMarkdownRTF = rtf
End Function

' ============================================================
'   Procesar línea Markdown normal
' ============================================================

Private Function ProcesarLineaMarkdown(linea As String) As String
    Dim txt As String
    txt = linea
    
    ' --- Títulos ---
    If Left(txt, 2) = "##" Then
        txt = "\b " & Mid(txt, 3) & " \b0"
        ProcesarLineaMarkdown = txt
        Exit Function
    End If
    
    If Left(txt, 1) = "#" Then
        txt = "\b\fs32 " & Mid(txt, 2) & " \fs20\b0"
        ProcesarLineaMarkdown = txt
        Exit Function
    End If
    
    ' --- Listas ---
    If Left(txt, 1) = "-" Then
        txt = "\bullet " & Mid(txt, 2)
    End If
    
    ' --- Negrita ---
    txt = ReemplazarNegrita(txt)
    
    ' --- Cursiva ---
    txt = ReemplazarCursiva(txt)
    
    ProcesarLineaMarkdown = txt
End Function

' ============================================================
'   Negrita y cursiva
' ============================================================

Private Function ReemplazarNegrita(txt As String) As String
    Do While InStr(txt, "**") > 0
        txt = Replace(txt, "**", "\b ", 1, 1)
        txt = Replace(txt, "**", " \b0", 1, 1)
    Loop
    ReemplazarNegrita = txt
End Function

Private Function ReemplazarCursiva(txt As String) As String
    Do While InStr(txt, "*") > 0
        txt = Replace(txt, "*", "\i ", 1, 1)
        txt = Replace(txt, "*", " \i0", 1, 1)
    Loop
    ReemplazarCursiva = txt
End Function

' ============================================================
'   Conversión de tablas Markdown ? RTF
' ============================================================

Private Function ConvertirTablaMarkdownRTF(mdTabla As String) As String
    Dim lineas() As String
    Dim linea As Variant
    Dim columnas() As String
    Dim rtf As String
    Dim i As Long
    Dim pos As Long
    Dim ancho As Long
    
    lineas = Split(mdTabla, vbCrLf)
    
    ' Ancho fijo por columna (2000 twips ˜ 3.5 cm)
    ancho = 2000
    
    rtf = ""
    
    For Each linea In lineas
        
        If Trim(linea) = "" Then GoTo Siguiente
        
        If InStr(linea, "|") = 0 Then GoTo Siguiente
        
        columnas = Split(linea, "|")
        
        ' Quitar bordes vacíos
        If columnas(0) = "" Then columnas = Split(Mid(linea, 2), "|")
        If columnas(UBound(columnas)) = "" Then ReDim Preserve columnas(0 To UBound(columnas) - 1)
        
        ' --- Nueva fila ---
        rtf = rtf & "\trowd" & vbCrLf
        
        ' Definir celdas
        pos = 0
        For i = 0 To UBound(columnas)
            pos = pos + ancho
            rtf = rtf & "\cellx" & pos
        Next i
        
        ' Contenido
        rtf = rtf & "\intbl "
        For i = 0 To UBound(columnas)
            rtf = rtf & Trim(columnas(i)) & "\cell "
        Next i
        
        rtf = rtf & "\row" & vbCrLf
        
Siguiente:
    Next linea
    
    ConvertirTablaMarkdownRTF = rtf
End Function

