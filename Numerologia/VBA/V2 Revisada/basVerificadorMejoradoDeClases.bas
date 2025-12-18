Attribute VB_Name = "basVerificadorMejoradoDeClases"

Option Compare Database
Option Explicit

' ============================================================
' VERIFICADOR MEJORADO DE CLASES
' ============================================================
' Este verificador:
'   - Comprueba si cada componente es realmente una clase
'   - Comprueba si es instanciable
'   - Comprueba si tiene VB_Name válido
'   - Comprueba si tiene PredeclaredId
'   - NO depende del texto del módulo
' ============================================================

Public Sub VerificarClasesMejorado()
    Dim comp As Object
    Dim errores As Long
    Dim nombre As String
    
    errores = 0
    
    Debug.Print
    Debug.Print "==============================================="
    Debug.Print "   VERIFICACIÓN MEJORADA DE CLASES DEL PROYECTO"
    Debug.Print "==============================================="
    Debug.Print
    
    For Each comp In Application.VBE.ActiveVBProject.VBComponents
        
        ' Solo analizamos módulos de clase
        If comp.Type = vbext_ct_ClassModule Then
            
            nombre = comp.Name
            Debug.Print "Clase encontrada: "; nombre
            
            ' 1. Verificar nombre válido
            If Len(nombre) = 0 Then
                Debug.Print "  >> ERROR: VB_Name vacío"
                errores = errores + 1
            End If
            
            ' 2. Verificar instanciación
            If Not EsInstanciable(nombre) Then
                Debug.Print "  >> ERROR: No se puede instanciar (posible clase corrupta)"
                errores = errores + 1
            End If
            
            ' 3. Verificar PredeclaredId (si existe)
            If TienePredeclaredId(comp) Then
                Debug.Print "  PredeclaredId: True"
            Else
                Debug.Print "  PredeclaredId: False"
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
' COMPRUEBA SI UNA CLASE ES INSTANCIABLE
' ============================================================
Private Function EsInstanciable(nombreClase As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim obj As Object
    Set obj = CreateObject(nombreClase)
    
    ' Si llega aquí, es instanciable
    EsInstanciable = True
    Exit Function
    
ErrHandler:
    ' Intento alternativo: New
    On Error GoTo ErrHandler2
    Dim x As Object
    Set x = VBA.CreateObject("", nombreClase)
    EsInstanciable = True
    Exit Function
    
ErrHandler2:
    EsInstanciable = False
End Function

' ============================================================
' DETECTA SI UNA CLASE TIENE PredeclaredId
' ============================================================
Private Function TienePredeclaredId(comp As Object) As Boolean
    On Error Resume Next
    TienePredeclaredId = comp.Properties("PredeclaredId")
End Function

