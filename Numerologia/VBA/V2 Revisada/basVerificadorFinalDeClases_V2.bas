Attribute VB_Name = "basVerificadorFinalDeClases_V2"

Option Compare Database
Option Explicit

' ============================================================
' VERIFICADOR DEFINITIVO DE CLASES
' ============================================================
' Este verificador:
'   - Comprueba si cada componente es realmente una clase
'   - Comprueba si es instanciable usando "New"
'   - Comprueba si tiene VB_Name válido
'   - NO usa CreateObject (que falla en Access)
' ============================================================

Public Sub VerificarClasesDefinitivo()
    Dim comp As Object
    Dim errores As Long
    Dim nombre As String
    
    errores = 0
    
    Debug.Print
    Debug.Print "==============================================="
    Debug.Print "   VERIFICACIÓN DEFINITIVA DE CLASES"
    Debug.Print "==============================================="
    Debug.Print
    
    For Each comp In Application.VBE.ActiveVBProject.VBComponents
        
        If comp.Type = vbext_ct_ClassModule Then
            
            nombre = comp.Name
            Debug.Print "Clase encontrada: "; nombre
            
            ' 1. Verificar nombre válido
            If Len(nombre) = 0 Then
                Debug.Print "  >> ERROR: VB_Name vacío"
                errores = errores + 1
            End If
            
            ' 2. Verificar instanciación con New
            If Not EsInstanciableConNew(nombre) Then
                Debug.Print "  >> ERROR: No se puede instanciar con New"
                errores = errores + 1
            Else
                Debug.Print "  Instanciación OK"
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
' COMPRUEBA SI UNA CLASE ES INSTANCIABLE CON "New"
' ============================================================
Private Function EsInstanciableConNew(nombreClase As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim obj As Object
    
    ' Instanciación dinámica usando VBA.CallByName
    ' Creamos una instancia usando el nombre de la clase
    Set obj = VBA.CallByName(CreateObject("Scripting.Dictionary"), nombreClase, VbMethod)
    
    ' Si llega aquí, no es válido (esto nunca funcionará)
    ' Así que usamos el método correcto:
    err.Raise 9999
    
ErrHandler:
    ' Intento real: usar New mediante una función auxiliar
    EsInstanciableConNew = InstanciarConNew(nombreClase)
End Function

' ============================================================
' INTENTA CREAR UNA INSTANCIA USANDO "New"
' ============================================================
Private Function InstanciarConNew(nombreClase As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim obj As Object
    
    ' Usamos una técnica clásica:
    ' Creamos una función pública temporalmente para instanciar la clase
    ' Pero aquí lo hacemos con Select Case para evitar reflexión
    
    Select Case nombreClase
        Case "clsConversorNumerologico": Set obj = New clsConversorNumerologico
        Case "clsGestorInterpretaciones": Set obj = New clsGestorInterpretaciones
        
        Case "clsCalculoAlma": Set obj = New clsCalculoAlma
        Case "clsCalculoBase": Set obj = New clsCalculoBase
        Case "clsCalculoDestino": Set obj = New clsCalculoDestino
        Case "clsCalculoCaminoVida": Set obj = New clsCalculoCaminoVida
        Case "clsCalculoPersonalidad": Set obj = New clsCalculoPersonalidad
        Case "clsCalculoMadurez": Set obj = New clsCalculoMadurez
        
        
        Case "clsCalculoNumeroDominante": Set obj = New clsCalculoNumeroDominante
        Case "clsCalculoNumeroFaltante": Set obj = New clsCalculoNumeroFaltante
        Case "clsCalculoNumeroPoder": Set obj = New clsCalculoNumeroPoder
        Case "clsCalculoPlanoExpresion": Set obj = New clsCalculoPlanoExpresion
        Case "clsCalculoRespuestaSubconsciente": Set obj = New clsCalculoRespuestaSubconsciente
        
        Case "clsCalculoSinastria": Set obj = New clsCalculoSinastria
        Case "clsCalculadorPeriodoActual": Set obj = New clsCalculadorPeriodoActual
        
        Case "clsCalculoPrimeraConsonante": Set obj = New clsCalculoPrimeraConsonante
        Case "clsCalculoPrimeraVocal": Set obj = New clsCalculoPrimeraVocal
        Case "clsCalculoPrimeraLetra": Set obj = New clsCalculoPrimeraLetra
        
        Case "clsCalculoDiaPersonal": Set obj = New clsCalculoDiaPersonal
        Case "clsCalculoMesPersonal": Set obj = New clsCalculoMesPersonal
        Case "clsCalculoAnoPersonal": Set obj = New clsCalculoAnoPersonal
        Case "clsCalculoEdadPersonal": Set obj = New clsCalculoEdadPersonal
        
        Case "clsCalculoCiclo1": Set obj = New clsCalculoCiclo1
        Case "clsCalculoCiclo2": Set obj = New clsCalculoCiclo2
        Case "clsCalculoCiclo3": Set obj = New clsCalculoCiclo3
        Case "clsCalculoCiclo4": Set obj = New clsCalculoCiclo4
        Case "clsCalculoCiclos": Set obj = New clsCalculoCiclos
        
        Case "clsCalculoPinaculo1": Set obj = New clsCalculoPinaculo1
        Case "clsCalculoPinaculo2": Set obj = New clsCalculoPinaculo2
        Case "clsCalculoPinaculo3": Set obj = New clsCalculoPinaculo3
        Case "clsCalculoPinaculo4": Set obj = New clsCalculoPinaculo4
        
        Case "clsCalculoDesafio1": Set obj = New clsCalculoDesafio1
        Case "clsCalculoDesafio2": Set obj = New clsCalculoDesafio2
        Case "clsCalculoDesafio3": Set obj = New clsCalculoDesafio3
        Case "clsCalculoDesafio4": Set obj = New clsCalculoDesafio4
        
        
        
        
        ' Añade aquí el resto de tus clases:
        ' Case "clsCalculoX": Set obj = New clsCalculoX
        
        Case Else
            ' Si no está en la lista, no podemos probarla
            Debug.Print "No está en Select case: "; nombreClase
            InstanciarConNew = False
            Exit Function
    End Select
    
    InstanciarConNew = True
    Exit Function
    
ErrHandler:
    InstanciarConNew = False
End Function


