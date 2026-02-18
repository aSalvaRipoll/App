' ============================================================
' Módulo: modExportarReferencias
' Exporta el manifiesto técnico del proyecto en formato .ref
' ============================================================

Option Compare Database
Option Explicit

Public DicDependencias As Object   ' Scripting.Dictionary


Public Sub ExportarReferencias(RutaBase As String)
    ' 1. Construir ruta del archivo .ref
    ' 2. Abrir archivo para escritura
    ' 3. Escribir sección [Proyecto]
    ' 4. Escribir sección [Referencia] por cada referencia
    ' 5. Escribir sección [DependenciasInternas]
    ' 6. Escribir sección [Entorno]
    ' 7. Cerrar archivo
End Sub


Public Sub InicializarDependencias()
    Set DicDependencias = CreateObject("Scripting.Dictionary")
End Sub

' --- Secciones ---
Private Sub EscribirProyecto(f As Integer)
End Sub

Private Sub EscribirReferencias(f As Integer)
End Sub

Private Sub EscribirDependenciasInternas(f As Integer)
End Sub

Private Sub EscribirEntorno(f As Integer)
End Sub

' --- Auxiliares ---
Private Function LimpiarTexto(s As String) As String
End Function

Private Function DetectarDependencias(objName As String, texto As String) As String
End Function


Private Sub EscribirProyecto(f As Integer)

    Print #f, "[Proyecto]"
    Print #f, "Nombre = """ & VBE.ActiveVBProject.Name & """"
    Print #f, "VersionAccess = """ & Application.Version & """"
    Print #f, "OptionCompare = ""Database"""
    Print #f, "OptionExplicit = ""True"""
    Print #f, "FechaExportacion = """ & Format(Now, "yyyy-mm-dd hh:nn:ss") & """"
    Print #f, ""

End Sub

Private Sub EscribirReferencias(f As Integer)

    Dim ref As Reference
    Dim estado As String

    For Each ref In Application.References

        If ref.IsBroken Then
            estado = "MISSING"
        Else
            estado = "OK"
        End If

        Print #f, "[Referencia]"
        Print #f, "Nombre = """ & ref.Name & """"
        Print #f, "GUID = """ & ref.GUID & """"
        Print #f, "Version = """ & ref.Major & "." & ref.Minor & """"

        On Error Resume Next
        Print #f, "Ruta = """ & ref.FullPath & """"
        On Error GoTo 0

        Print #f, "Estado = """ & estado & """"
        Print #f, ""

    Next ref

End Sub


Private Sub EscribirDependenciasInternas(f As Integer)

    Dim obj As AccessObject
    Dim Texto As String
    Dim deps As String

    Print #f, "[DependenciasInternas]"

    ' --- Formularios ---
    For Each obj In CurrentProject.AllForms
        Texto = ObtenerTextoObjeto(acForm, obj.Name)
        deps = DetectarDependencias(Texto)
        If deps <> "" Then Print #f, "Form_" & obj.Name & " = """ & deps & """"
    Next obj

    ' --- Informes ---
    For Each obj In CurrentProject.AllReports
        Texto = ObtenerTextoObjeto(acReport, obj.Name)
        deps = DetectarDependencias(Texto)
        If deps <> "" Then Print #f, "Report_" & obj.Name & " = """ & deps & """"
    Next obj

    ' --- Módulos ---
    Dim comp As VBIDE.VBComponent
    For Each comp In VBE.ActiveVBProject.VBComponents
        If comp.Type = vbext_ct_StdModule Or comp.Type = vbext_ct_ClassModule Then
            Texto = comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines)
            deps = DetectarDependencias(Texto)
            If deps <> "" Then Print #f, "Modulo_" & comp.Name & " = """ & deps & """"
        End If
    Next comp

    Print #f, ""

End Sub

Private Function ObtenerTextoObjeto(Tipo As AcObjectType, Nombre As String) As String

    Dim RutaTemp As String
    Dim f As Integer
    Dim s As String

    RutaTemp = Environ("TEMP") & "\__temp_export.txt"
    SaveAsText Tipo, Nombre, RutaTemp

    f = FreeFile
    Open RutaTemp For Input As #f
    s = Input$(LOF(f), f)
    Close #f

    Kill RutaTemp

    ObtenerTextoObjeto = s

End Function


Private Function DetectarDependencias(Texto As String) As String

    Dim comp As VBIDE.VBComponent
    Dim lista As String

    For Each comp In VBE.ActiveVBProject.VBComponents
        If InStr(1, Texto, comp.Name, vbTextCompare) > 0 Then
            If lista <> "" Then lista = lista & ", "
            lista = lista & comp.Name
        End If
    Next comp

    DetectarDependencias = lista

End Function


Private Sub EscribirEntorno(f As Integer)

    Print #f, "[Entorno]"
    Print #f, "SO = """ & Environ$("OS") & """"
    Print #f, "Arquitectura = """ & Environ$("PROCESSOR_ARCHITECTURE") & """"
    Print #f, "Usuario = """ & Environ$("USERNAME") & """"
    Print #f, "Localizacion = """ & Application.LanguageSettings.LanguageID(msoLanguageIDUI) & """"
    Print #f, ""

End Sub


'================================================================================================================
'Parte a incluir
' --- NUEVO: detectar dependencias ---
    Call RegistrarDependencias("Form_" & obj.Name, RutaObj)
	
	
Public Sub RegistrarDependencias(Clave As String, RutaArchivo As String)

    Dim f As Integer
    Dim Texto As String
    Dim deps As String

    f = FreeFile
    Open RutaArchivo For Input As #f
    Texto = Input$(LOF(f), f)
    Close #f

    deps = DetectarDependencias(Texto)

    If deps <> "" Then
        DicDependencias(Clave) = deps
    End If

End Sub

Private Function DetectarDependencias(Texto As String) As String

    Dim comp As VBIDE.VBComponent
    Dim lista As String

    For Each comp In VBE.ActiveVBProject.VBComponents
        If InStr(1, Texto, comp.Name, vbTextCompare) > 0 Then
            If lista <> "" Then lista = lista & ", "
            lista = lista & comp.Name
        End If
    Next comp

    DetectarDependencias = lista

End Function
	
Private Sub EscribirDependenciasInternas(f As Integer)

    Dim clave As Variant

    Print #f, "[DependenciasInternas]"

    For Each clave In DicDependencias.Keys
        Print #f, clave & " = """ & DicDependencias(clave) & """"
    Next clave

    Print #f, ""

End Sub

