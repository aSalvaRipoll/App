Attribute VB_Name = "modInspectorMain"

Option Compare Database
Option Explicit

Public Sub EjecutarInspectorProyecto()
    Dim insp As clsAnalizadorProyecto
    Dim res As clsResultadoAnalisis
    Dim hayErrores As Boolean
    
    If Not AsegurarReferenciaVBIDE() Then
        Debug.Print "? No se puede ejecutar el Inspector sin VBIDE."
        Exit Sub
    End If
    
    Set insp = New clsAnalizadorProyecto
    
    Debug.Print
    Debug.Print "==============================================="
    Debug.Print "   INSPECTOR DE PROYECTO VBA - INICIO"
    Debug.Print "==============================================="
    Debug.Print
    
    insp.AnalizarProyectoActual
    
    For Each res In insp.Resultados
        Debug.Print res.Formatear
        If res.Severidad = sevError Then hayErrores = True
    Next res
    
    Debug.Print
    Debug.Print "==============================================="
    Debug.Print "   RESUMEN"
    Debug.Print "==============================================="
    
    If hayErrores Then
        Debug.Print "Se han detectado ERRORES. Revisa los resultados."
    Else
        Debug.Print "No se han detectado errores graves."
    End If
    
    Debug.Print "==============================================="
End Sub



'Option Compare Database
'Option Explicit
'
'Public Sub EjecutarInspectorProyecto()
'    Dim insp As clsAnalizadorProyecto
'    Dim res As clsResultadoAnalisis
'    Dim hayErrores As Boolean
'
''    If Not ReferenciaExtensibilidadActiva() Then
''        Debug.Print "??  La referencia 'Microsoft Visual Basic for Applications Extensibility 5.3' no está activa."
''        Debug.Print "    El inspector no podrá analizar clases ni módulos."
''        Debug.Print "    Ejecuta ActivarReferenciaExtensibilidad para activarla."
''        Debug.Print
''        If MsgBox("¿Desea intentar la agregación automática?", vbYesNo + vbQuestion, "Agregar referencia") = True Then
''            ActivarReferenciaExtensibilidad
''            If Not ReferenciaExtensibilidadActiva() Then
''                Debug.Print "??  La referencia 'Microsoft Visual Basic for Applications Extensibility 5.3' no está activa."
''                Debug.Print "    El inspector no podrá analizar clases ni módulos."
''                Debug.Print "    Agregue la referencia 'Microsoft Visual Basic for Applications Extensibility 5.3' manualmente."
''                Debug.Print
''            End If
''        End If
''    End If
'
'    If Not AsegurarReferenciaVBIDE() Then
'        Debug.Print "? No se puede ejecutar el Inspector sin VBIDE."
'        Exit Sub
'    End If
'
'    Set insp = New clsAnalizadorProyecto
'
'    Debug.Print
'    Debug.Print "==============================================="
'    Debug.Print "   INSPECTOR DE PROYECTO VBA - INICIO"
'    Debug.Print "==============================================="
'    Debug.Print
'
'    insp.AnalizarProyectoActual
'
'    For Each res In insp.Resultados
'        Debug.Print res.Formatear
'        If res.Severidad = sevError Then hayErrores = True
'    Next res
'
'    Debug.Print
'    Debug.Print "==============================================="
'    Debug.Print "   RESUMEN"
'    Debug.Print "==============================================="
'
'    If hayErrores Then
'        Debug.Print "Se han detectado ERRORES. Revisa los resultados."
'    Else
'        Debug.Print "No se han detectado errores graves."
'    End If
'
'    Debug.Print "==============================================="
'End Sub
'
