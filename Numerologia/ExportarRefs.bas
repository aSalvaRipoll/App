Option Compare Database
Option Explicit

Public DicDependencias As Object   ' Scripting.Dictionary

Public Sub InicializarDependencias()
    Set DicDependencias = CreateObject("Scripting.Dictionary")
End Sub

Public Sub RegistrarDependencias(NombreLogico As String, RutaCarpeta As String)

    Dim f As Integer
    Dim Texto As String
    Dim RutaArchivo As String
    Dim deps As String

    RutaArchivo = RutaCarpeta & "\" & NombreLogico

    f = FreeFile
    Open RutaArchivo For Input As #f
    Texto = Input$(LOF(f), f)
    Close #f

    deps = DetectarDependencias(Texto)

    If deps <> "" Then
        DicDependencias(NombreLogico) = deps
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


Public Sub ExportarReferencias(RutaBase As String)

    Dim RutaRef As String
    Dim f As Integer
    Dim ref As Reference
    Dim estado As String
    Dim clave As Variant

    RutaRef = RutaBase & "\proyecto.ref"
    f = FreeFile
    Open RutaRef For Output As #f

    ' --- [Proyecto] ---
    Print #f, "[Proyecto]"
    Print #f, "Nombre = """ & VBE.ActiveVBProject.Name & """"
    Print #f, "VersionAccess = """ & Application.Version & """"
    Print #f, "OptionCompare = ""Database"""
    Print #f, "OptionExplicit = ""True"""
    Print #f, "FechaExportacion = """ & Format(Now, "yyyy-mm-dd hh:nn:ss") & """"
    Print #f, ""

    ' --- [Referencia] ---
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

    ' --- [DependenciasInternas] ---
    Print #f, "[DependenciasInternas]"
    For Each clave In DicDependencias.Keys
        Print #f, clave & " = """ & DicDependencias(clave) & """"
    Next clave
    Print #f, ""

    ' --- [Entorno] ---
    Print #f, "[Entorno]"
    Print #f, "SO = """ & Environ$("OS") & """"
    Print #f, "Arquitectura = """ & Environ$("PROCESSOR_ARCHITECTURE") & """"
    Print #f, "Usuario = """ & Environ$("USERNAME") & """"
    Print #f, "Localizacion = """ & Application.LanguageSettings.LanguageID(msoLanguageIDUI) & """"
    Print #f, ""

    Close #f

End Sub
