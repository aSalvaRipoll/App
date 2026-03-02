Attribute VB_Name = "___basInfoModulos"


Option Compare Database
Option Explicit

Public Sub ExportarInfoModulos()
    Dim vbProj As VBIDE.VBProject
    Dim vbComp As VBIDE.VBComponent
    Dim vbMod As VBIDE.CodeModule
    
    Dim ruta As String
    Dim f As Integer
    Dim linea As Long
    Dim nombreProc As String
    Dim tipoProc As VBIDE.vbext_ProcKind
    Dim visib As String
    
    Set vbProj = Application.VBE.ActiveVBProject
    
    ruta = CurrentProject.Path & "\ListadoModulos_" & vbProj.Name & ".txt"
    f = FreeFile
    
    Open ruta For Output As #f
    
    ' --- Nombre del proyecto ---
    Print #f, "PROYECTO: " & vbProj.Name
    Print #f, String(40, "-")
    Print #f, ""
    
    For Each vbComp In vbProj.VBComponents
        Set vbMod = vbComp.CodeModule
        
        Print #f, "Módulo: " & vbComp.Name
        Print #f, "Tipo:   " & TipoModulo(vbComp.Type)
        Print #f, "Procedimientos:"
        
        linea = 1
        Do While linea < vbMod.CountOfLines
            nombreProc = vbMod.ProcOfLine(linea, tipoProc)
            If nombreProc <> "" Then
                
                ' Obtener visibilidad del procedimiento
                visib = VisibilidadProcedimiento(vbMod, nombreProc, tipoProc)
                
                Print #f, "   - " & nombreProc & "   (" & visib & ")"
                
                linea = vbMod.ProcStartLine(nombreProc, tipoProc) + _
                        vbMod.ProcCountLines(nombreProc, tipoProc)
            Else
                linea = linea + 1
            End If
        Loop
        
        Print #f, ""
    Next vbComp
    
    Close #f
    
    MsgBox "Fichero generado en: " & ruta, vbInformation
End Sub

Private Function TipoModulo(tipo As VBIDE.vbext_ComponentType) As String
    Select Case tipo
        Case vbext_ct_StdModule: TipoModulo = "Módulo estándar"
        Case vbext_ct_ClassModule: TipoModulo = "Módulo de clase"
        Case vbext_ct_MSForm: TipoModulo = "Formulario"
        Case vbext_ct_Document: TipoModulo = "Módulo de formulario/informe"
        Case Else: TipoModulo = "Desconocido"
    End Select
End Function

Private Function VisibilidadProcedimiento(cm As VBIDE.CodeModule, _
                                          nombre As String, _
                                          tipo As VBIDE.vbext_ProcKind) As String
    Dim lineaInicio As Long
    Dim texto As String
    
    lineaInicio = cm.ProcStartLine(nombre, tipo)
    
    texto = Trim$(cm.Lines(lineaInicio, 1))
    While InStr(texto, nombre) = 0
        lineaInicio = lineaInicio + 1
        texto = Trim$(cm.Lines(lineaInicio, 1))
        If Left(texto, 1) = "'" Then texto = "'"
    Wend
    
    'texto = Trim$(cm.Lines(lineaInicio + 1, 1))
    
    If LCase$(Left$(texto, 6)) = "public" Then
        VisibilidadProcedimiento = "Public"
    ElseIf LCase$(Left$(texto, 7)) = "private" Then
        VisibilidadProcedimiento = "Private"
    Else
        ' Si no especifica, en módulos estándar es público por defecto
        VisibilidadProcedimiento = "Public (implícito)"
    End If
End Function

'Public Sub ExportarInfoModulos()
'    Dim vbProj As VBIDE.VBProject
'    Dim vbComp As VBIDE.VBComponent
'    Dim vbMod As VBIDE.CodeModule
'
'    Dim ruta As String
'    Dim f As Integer
'    Dim linea As Long
'    Dim nombreProc As String
'    Dim tipoProc As VBIDE.vbext_ProcKind
'
'    Set vbProj = Application.VBE.ActiveVBProject
'
'    ruta = CurrentProject.Path & "\ListadoModulos_" & vbProj.Name & ".txt"
'    f = FreeFile
'
'    Open ruta For Output As #f
'
'    ' --- Aquí añadimos el nombre del proyecto ---
'    Print #f, "PROYECTO: " & vbProj.Name
'    Print #f, String(40, "-")
'    Print #f, ""
'
'    For Each vbComp In vbProj.VBComponents
'        Set vbMod = vbComp.CodeModule
'
'        Print #f, "Módulo: " & vbComp.Name
'        Print #f, "Tipo:   " & TipoModulo(vbComp.Type)
'        Print #f, "Procedimientos:"
'
'        linea = 1
'        Do While linea < vbMod.CountOfLines
'            nombreProc = vbMod.ProcOfLine(linea, tipoProc)
'            If nombreProc <> "" Then
'                Print #f, "   - " & nombreProc
'                linea = vbMod.ProcStartLine(nombreProc, tipoProc) + _
'                        vbMod.ProcCountLines(nombreProc, tipoProc)
'            Else
'                linea = linea + 1
'            End If
'        Loop
'
'        Print #f, ""
'    Next vbComp
'
'    Close #f
'
'    MsgBox "Fichero generado en: " & ruta, vbInformation
'End Sub
'
'Private Function TipoModulo(tipo As VBIDE.vbext_ComponentType) As String
'    Select Case tipo
'        Case vbext_ct_StdModule: TipoModulo = "Módulo estándar"
'        Case vbext_ct_ClassModule: TipoModulo = "Módulo de clase"
'        Case vbext_ct_MSForm: TipoModulo = "Formulario"
'        Case vbext_ct_Document: TipoModulo = "Módulo de formulario/informe"
'        Case Else: TipoModulo = "Desconocido"
'    End Select
'End Function

'Public Sub ExportarInfoModulos()
'    Dim vbProj As VBIDE.VBProject
'    Dim vbComp As VBIDE.VBComponent
'    Dim vbMod As VBIDE.CodeModule
'
'    Dim ruta As String
'    Dim f As Integer
'    Dim linea As Long
'    Dim nombreProc As String
'    Dim tipoProc As VBIDE.vbext_ProcKind
'
'    Set vbProj = Application.VBE.ActiveVBProject
'
'    ruta = CurrentProject.Path & "\ListadoModulos.txt"
'    f = FreeFile
'
'    Open ruta For Output As #f
'
'    Print #f, "LISTADO DE MÓDULOS DEL PROYECTO"
'    Print #f, "--------------------------------"
'    Print #f, ""
'
'    For Each vbComp In vbProj.VBComponents
'        Set vbMod = vbComp.CodeModule
'
'        Print #f, "Módulo: " & vbComp.Name
'        Print #f, "Tipo:   " & TipoModulo(vbComp.Type)
'        Print #f, "Procedimientos:"
'
'        linea = 1
'        Do While linea < vbMod.CountOfLines
'            nombreProc = vbMod.ProcOfLine(linea, tipoProc)
'            If nombreProc <> "" Then
'                Print #f, "   - " & nombreProc
'                linea = vbMod.ProcStartLine(nombreProc, tipoProc) + _
'                        vbMod.ProcCountLines(nombreProc, tipoProc)
'            Else
'                linea = linea + 1
'            End If
'        Loop
'
'        Print #f, ""
'    Next vbComp
'
'    Close #f
'
'    MsgBox "Fichero generado en: " & ruta, vbInformation
'End Sub
'
'Private Function TipoModulo(tipo As VBIDE.vbext_ComponentType) As String
'    Select Case tipo
'        Case vbext_ct_StdModule: TipoModulo = "Módulo estándar"
'        Case vbext_ct_ClassModule: TipoModulo = "Módulo de clase"
'        Case vbext_ct_MSForm: TipoModulo = "Formulario"
'        Case vbext_ct_Document: TipoModulo = "Módulo de formulario/informe"
'        Case Else: TipoModulo = "Desconocido"
'    End Select
'End Function


