Attribute VB_Name = "modInspectorRepair"

Option Compare Database
Option Explicit

Public Sub RepararProblemasProyecto()
    Dim insp As clsAnalizadorProyecto
    Dim res As clsResultadoAnalisis
    
    If Not AsegurarReferenciaVBIDE() Then
        Debug.Print "? No se puede ejecutar la reparación sin VBIDE."
        Exit Sub
    End If
    
    Set insp = New clsAnalizadorProyecto
    insp.AnalizarProyectoActual
    
    Debug.Print
    Debug.Print "==============================================="
    Debug.Print "   INICIO DE REPARACIÓN"
    Debug.Print "==============================================="
    
    For Each res In insp.Resultados
        If res.esReparable Then
            RepararResultado res
        End If
    Next res
    
    Debug.Print "==============================================="
    Debug.Print "   FIN DE REPARACIÓN"
    Debug.Print "==============================================="
End Sub

Private Sub RepararResultado(res As clsResultadoAnalisis)
    Dim codigo As String
    codigo = res.codigoReparacion
    
    Debug.Print "Reparando: "; res.Formatear
    
    Select Case True
        
        Case codigo = "ADD_OPTION_EXPLICIT"
            RepararOptionExplicit res.TipoElemento, res.NombreElemento
        
        Case Left$(codigo, 14) = "FIX_REFERENCE:"
            RepararReferenciaEspecifica Mid$(codigo, 15)
        
        Case Else
            Debug.Print "  >> No hay rutina de reparación definida para: "; codigo
    End Select
End Sub

Private Sub RepararOptionExplicit(Tipo As TipoElementoInspector, nombre As String)
    Dim comp As Object
    Dim vbProj As Object
    Dim cm As Object
    Dim insertLinea As Long
    Dim i As Long, linea As String
    
    Set vbProj = Application.VBE.ActiveVBProject
    
    For Each comp In vbProj.VBComponents
        If comp.Name = nombre Then
            Set cm = comp.CodeModule
            
            insertLinea = 1
            For i = 1 To cm.CountOfLines
                linea = Trim(cm.Lines(i, 1))
                If Len(linea) = 0 Or Left$(linea, 1) = "'" Then
                    insertLinea = i + 1
                Else
                    Exit For
                End If
            Next i
            
            cm.InsertLines insertLinea, "Option Explicit"
            Debug.Print "  >> Option Explicit añadido en "; nombre
            Exit Sub
        End If
    Next comp
End Sub

Private Sub RepararReferenciaEspecifica(nombreRef As String)
    Dim vbProj As VBIDE.VBProject
    Dim ref As Reference
    On Error Resume Next
    
    Set vbProj = Application.VBE.ActiveVBProject
    
    If nombreRef = "VBIDE" Then
        If AsegurarReferenciaVBIDE() Then
            Debug.Print "  >> Referencia VBIDE reparada."
        Else
            Debug.Print "  >> No se pudo reparar VBIDE automáticamente."
        End If
    Else
        Debug.Print "  >> Reparación automatizada no implementada para referencia: "; nombreRef
    End If
End Sub



'Option Compare Database
'Option Explicit
'
'' ============================================================
''   MOTOR DE REPARACIÓN AVANZADA
'' ============================================================
'
'Public Sub RepararProblemasProyecto()
'    Dim insp As clsAnalizadorProyecto
'    Dim res As clsResultadoAnalisis
'
'    Set insp = New clsAnalizadorProyecto
'    insp.AnalizarProyectoActual
'
'    Debug.Print
'    Debug.Print "==============================================="
'    Debug.Print "   INICIO DE REPARACIÓN"
'    Debug.Print "==============================================="
'
'    For Each res In insp.Resultados
'        If res.esReparable Then
'            RepararResultado res
'        End If
'    Next res
'
'    Debug.Print "==============================================="
'    Debug.Print "   FIN DE REPARACIÓN"
'    Debug.Print "==============================================="
'End Sub
'
'Private Sub RepararResultado(res As clsResultadoAnalisis)
'    Dim codigo As String
'    codigo = res.codigoReparacion
'
'    Debug.Print "Reparando: "; res.Formatear
'
'    Select Case True
'
'        Case codigo = "ADD_OPTION_EXPLICIT"
'            RepararOptionExplicit res.TipoElemento, res.NombreElemento
'
'        Case Left$(codigo, 14) = "FIX_REFERENCE:"
'            RepararReferenciaEspecifica Mid$(codigo, 15)
'
'        Case codigo = "REMOVE_BOM"
'            RepararBOMEnModulo res.TipoElemento, res.NombreElemento
'
'        ' Aquí irás añadiendo nuevos códigos de reparación:
'        ' Case codigo = "REBUILD_CLASS_HEADER": ...
'
'        Case Else
'            Debug.Print "  >> No hay rutina de reparación definida para: "; codigo
'    End Select
'End Sub
'
'Private Sub RepararOptionExplicit(tipo As TipoElementoInspector, nombre As String)
'    Dim comp As Object
'    Dim vbProj As Object
'    Dim cm As Object
'    Dim insertLinea As Long
'    Dim i As Long, linea As String
'
'    Set vbProj = Application.VBE.ActiveVBProject
'
'    For Each comp In vbProj.VBComponents
'        If comp.Name = nombre Then
'            Set cm = comp.CodeModule
'            ' Buscar la primera línea donde insertar Option Explicit
'            insertLinea = 1
'            For i = 1 To cm.CountOfLines
'                linea = Trim(cm.Lines(i, 1))
'                If Len(linea) = 0 Or Left$(linea, 1) = "'" Then
'                    insertLinea = i + 1
'                Else
'                    Exit For
'                End If
'            Next i
'
'            cm.InsertLines insertLinea, "Option Explicit"
'
'            Debug.Print "  >> Option Explicit añadido en "; nombre
'            Exit Sub
'        End If
'    Next comp
'End Sub
'
'Private Sub RepararReferenciaEspecifica(nombreRef As String)
'    Dim vbProj As VBIDE.VBProject
'    Dim ref As Reference
'    On Error Resume Next
'
'    Set vbProj = Application.VBE.ActiveVBProject
'
'    ' De momento, solo podemos tratar VBIDE de forma automática
'    If nombreRef = "VBIDE" Then
'        If AsegurarReferenciaVBIDE() Then
'            Debug.Print "  >> Referencia VBIDE reparada."
'        Else
'            Debug.Print "  >> No se pudo reparar VBIDE automáticamente."
'        End If
'    Else
'        Debug.Print "  >> Reparación automatizada no implementada para referencia: "; nombreRef
'    End If
'End Sub
'
'
