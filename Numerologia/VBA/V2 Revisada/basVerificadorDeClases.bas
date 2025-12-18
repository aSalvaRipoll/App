Attribute VB_Name = "basVerificadorDeClases"
Option Compare Database
Option Explicit

' ============================================================
' VERIFICADOR FINAL DE CLASES
' ============================================================
' Este verificador:
'   - Recorre todas las clases del proyecto
'   - Comprueba que son vbext_ct_ClassModule
'   - Comprueba que tienen VB_Name válido
'   - Comprueba que no hay duplicados
'   - Comprueba que no hay módulos estándar disfrazados
'   - Muestra un informe claro en la ventana Inmediato
' ============================================================

Public Sub VerificarClasesFinal()
    Dim comp As Object
    Dim dictNombres As Object
    Dim nombre As String
    Dim tipo As String
    Dim errores As Long
    
    Set dictNombres = CreateObject("Scripting.Dictionary")
    errores = 0
    
    Debug.Print
    Debug.Print "==============================================="
    Debug.Print "   VERIFICACIÓN FINAL DE CLASES DEL PROYECTO"
    Debug.Print "==============================================="
    Debug.Print
    
    For Each comp In Application.VBE.ActiveVBProject.VBComponents
        
        ' Solo analizamos clases
        If comp.Type = vbext_ct_ClassModule Then
            
            nombre = comp.Name
            
            Debug.Print "Clase encontrada: "; nombre
            
            ' 1. Verificar duplicados
            If dictNombres.Exists(nombre) Then
                Debug.Print "  >> ERROR: Clase duplicada: "; nombre
                errores = errores + 1
            Else
                dictNombres.Add nombre, True
            End If
            
            ' 2. Verificar VB_Name
            If Len(nombre) = 0 Then
                Debug.Print "  >> ERROR: Clase sin VB_Name"
                errores = errores + 1
            End If
            
            ' 3. Verificar que no sea módulo estándar disfrazado
            If Not TieneCabeceraDeClase(comp) Then
                Debug.Print "  >> ERROR: No tiene cabecera interna de clase"
                errores = errores + 1
            End If
            
            Debug.Print "  OK"
            Debug.Print
            
        End If
    Next comp
    
    Debug.Print "==============================================="
    Debug.Print "   RESULTADO FINAL"
    Debug.Print "==============================================="
    
    If errores = 0 Then
        Debug.Print "? Todas las clases están correctas."
    Else
        Debug.Print "? Se han encontrado "; errores; " problemas."
        Debug.Print "Revisa los mensajes anteriores."
    End If
    
    Debug.Print "==============================================="
End Sub

' ============================================================
' VERIFICA SI UNA CLASE TIENE CABECERA INTERNA REAL
' ============================================================
Private Function TieneCabeceraDeClase(comp As Object) As Boolean
    Dim linea As String
    Dim i As Long
    
    On Error Resume Next
    
    With comp.CodeModule
        For i = 1 To .CountOfLines
            linea = Trim(.Lines(i, 1))
            
            If linea Like "VERSION *" Then TieneCabeceraDeClase = True: Exit Function
            If linea = "BEGIN" Then TieneCabeceraDeClase = True: Exit Function
            If InStr(1, linea, "Attribute VB_Name", vbTextCompare) > 0 Then TieneCabeceraDeClase = True: Exit Function
        Next i
    End With
    
    TieneCabeceraDeClase = False
End Function

Sub PruebaInstancia()

Dim x As New clsCalculoAlma
Debug.Print TypeName(x)

End Sub
