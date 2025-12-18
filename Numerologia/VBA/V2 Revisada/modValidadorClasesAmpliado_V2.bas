Attribute VB_Name = "modValidadorClasesAmpliado_V2"

Option Compare Database
Option Explicit

Public Sub VerificarClases()
    Dim comp As Object   ' VBIDE.VBComponent
    Dim dictNombres As Object
    Dim nombre As String
    Dim errores As String
    
    Set dictNombres = CreateObject("Scripting.Dictionary")
    errores = ""
    
    Debug.Print "=== VERIFICACIÓN DE CLASES (Access 2019) ==="
    
    For Each comp In Application.VBE.ActiveVBProject.VBComponents
        
        Select Case comp.Type
            
            Case vbext_ct_ClassModule
                nombre = ObtenerVBName(comp)
                
                If Len(nombre) = 0 Then
                    errores = errores & "ERROR: Clase sin VB_Name ? " & comp.Name & vbCrLf
                End If
                
                If dictNombres.Exists(nombre) Then
                    errores = errores & "ERROR: Nombre de clase duplicado ? " & nombre & vbCrLf
                Else
                    dictNombres.Add nombre, True
                End If
                
                If Not EsPublicNotCreatable(comp) Then
                    errores = errores & "ADVERTENCIA: Clase no es PublicNotCreatable ? " & nombre & vbCrLf
                End If
                
                Debug.Print "Clase detectada ? "; nombre
            
            Case vbext_ct_StdModule
                Debug.Print "Módulo estándar ? "; comp.Name
        End Select
        
    Next comp
    
    Debug.Print "=== FIN DE VERIFICACIÓN ==="
    
    If Len(errores) > 0 Then
        Debug.Print vbCrLf & "=== ERRORES DETECTADOS ==="
        Debug.Print errores
        MsgBox "Se detectaron problemas. Revisa la ventana Inmediato.", vbExclamation
    Else
        MsgBox "Todas las clases están correctas a nivel de estructura básica.", vbInformation
    End If
End Sub

Private Function ObtenerVBName(comp As Object) As String
    Dim linea As String
    Dim i As Long
    
    For i = 1 To comp.CodeModule.CountOfLines
        linea = comp.CodeModule.Lines(i, 1)
        If InStr(1, linea, "Attribute VB_Name", vbTextCompare) > 0 Then
            ObtenerVBName = ExtraerNombreVB(linea)
            Exit Function
        End If
    Next i
End Function

Private Function ExtraerNombreVB(linea As String) As String
    Dim p As Long
    p = InStr(linea, "=")
    If p > 0 Then
        ExtraerNombreVB = Trim(Replace(Mid(linea, p + 1), """", ""))
    End If
End Function

Private Function EsPublicNotCreatable(comp As Object) As Boolean
    Dim linea As String
    Dim i As Long
    
    For i = 1 To comp.CodeModule.CountOfLines
        linea = comp.CodeModule.Lines(i, 1)
        
        If InStr(1, linea, "MultiUse", vbTextCompare) > 0 Then
            If InStr(1, linea, "-1", vbTextCompare) > 0 Then
                EsPublicNotCreatable = True
            Else
                EsPublicNotCreatable = False
            End If
            Exit Function
        End If
    Next i
    
    EsPublicNotCreatable = False
End Function


