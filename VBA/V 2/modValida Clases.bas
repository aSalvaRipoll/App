Attribute VB_Name = "modValida Clases"
Option Compare Database
Option Explicit

Public Sub ValidarClasesDerivadas()
    Dim cls As AccessObject
    Dim errores As String
    Dim nombre As String
    
    errores = ""
    
    For Each cls In CurrentProject.AllModules
        nombre = cls.Name
        
        If Left(nombre, 3) = "cls" And nombre <> "clsCalculoBase" Then
            errores = errores & ValidarClase(nombre)
        End If
    Next cls
    
    If errores = "" Then
        MsgBox "? Todas las clases derivadas están correctas.", vbInformation
    Else
        MsgBox "? Se encontraron problemas:" & vbCrLf & errores, vbCritical
    End If
End Sub

Private Function ValidarClase(nombreClase As String) As String
    Dim m As Module
    Dim txt As String
    Dim err As String
    
    Set m = Modules(nombreClase)
    txt = m.Lines(1, m.CountOfLines)
    
    ' Validar TIPO_CALCULO
    If InStr(txt, "TIPO_CALCULO") = 0 Then
        err = err & "• Falta TIPO_CALCULO en " & nombreClase & vbCrLf
    End If
    
    ' Validar Calcular
    If InStr(txt, "Function Calcular") = 0 Then
        err = err & "• Falta Calcular() en " & nombreClase & vbCrLf
    End If
    
    ' Validar ObtenerInterpretacion
    If InStr(txt, "Function ObtenerInterpretacion") = 0 Then
        err = err & "• Falta ObtenerInterpretacion() en " & nombreClase & vbCrLf
    End If
    
    ' Detectar Implements
    If InStr(txt, "Implements") > 0 Then
        err = err & "• No debe usar Implements en " & nombreClase & vbCrLf
    End If
    
    ' Detectar instancias internas de la base
    If InStr(txt, "As clsCalculoBase") > 0 Then
        err = err & "• No debe instanciar clsCalculoBase dentro de " & nombreClase & vbCrLf
    End If
    
    ValidarClase = err
End Function

