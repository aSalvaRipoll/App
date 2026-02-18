Attribute VB_Name = "___versiones_viejas"
Option Compare Database
Option Explicit

'Private Sub Silabear 0()
'
'    Dim texto As String
'    Dim i As Long
'    Dim c As String, prev As String, nextC As String
'    Dim silaba As String
'    Dim resultado As Collection
'    Set resultado = New Collection
'
'    texto = ObjDTO.TextoNormalizado
'
'    If Len(texto) = 0 Then
'        ObjDTO.SilabasAuto = ""
'        Exit Sub
'    End If
'
'    silaba = ""
'
'    For i = 1 To Len(texto)
'
'        c = Mid$(texto, i, 1)
'
'        prev = ""
'        If i > 1 Then
'            prev = Mid$(texto, i - 1, 1)
'        End If
'
'        nextC = ""
'        If i < Len(texto) Then
'            nextC = Mid$(texto, i + 1)
'        End If
'
'        ' --- Espacio ---
'        If c = " " Then
'            If silaba <> "" Then resultado.Add silaba
'            resultado.Add " "
'            silaba = ""
'            GoTo siguiente
'        End If
'
'        ' --- Decidir si cortar ANTES de añadir c ---
'        If silaba <> "" Then
'
'            Dim ult As String
'            ult = Right$(silaba, 1)
'
'            ' 1) Hiato
'            If EsVocal(ult) And EsVocal(c) Then
'                If EsHiato(ult, c) Then
'                    resultado.Add silaba
'                    silaba = ""
'                End If
'
'            ' 2) VCV
'            ElseIf EsVocal(ult) And EsConsonante(c) And EsVocal(nextC) Then
'                resultado.Add silaba
'                silaba = ""
'
'            ' 3) CCV
'            ElseIf EsConsonante(ult) And EsConsonante(c) And EsVocal(nextC) Then
'                If Not EsGrupoInseparable(ult & c) Then
'                    resultado.Add silaba
'                    silaba = ""
'                End If
'
'            ' 4) CCC
'            ElseIf EsConsonante(ult) And EsConsonante(c) And EsConsonante(nextC) Then
'                resultado.Add silaba
'                silaba = ""
'            End If
'        End If
'
'        silaba = silaba & c
'
'siguiente:
'    Next i
'
'    If silaba <> "" Then resultado.Add silaba
'
'    ' --- Construir salida final ---
'    Dim out As String
'    out = ""
'
'    For i = 1 To resultado.Count
'        out = out & resultado(i)
'        If i < resultado.Count Then out = out & " | "
'    Next i
'
'    ObjDTO.SilabasAuto = out
'
'End Sub


