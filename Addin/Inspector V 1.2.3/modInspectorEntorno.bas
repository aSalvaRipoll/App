Attribute VB_Name = "modInspectorEntorno"

Option Compare Database
Option Explicit

' ============================================================
'  MÓDULO DE ENTORNO DEL INSPECTOR
'  Detecta idioma, entorno, modo runtime, referencias rotas,
'  y registra errores automáticamente.
' ============================================================


' ------------------------------------------------------------
' DETECCIÓN DE ENTORNO
' ------------------------------------------------------------

' Devuelve el idioma del sistema (ej: "es-ES", "en-US")
Public Function IdiomaSistema() As String
    On Error Resume Next
    IdiomaSistema = Environ$("LANG")
    If IdiomaSistema = "" Then IdiomaSistema = Application.LanguageSettings.LanguageID(msoLanguageIDUI)
End Function

' Devuelve True si Access está en modo Runtime
Public Function EsModoRuntime() As Boolean
    On Error Resume Next
    EsModoRuntime = (SysCmd(acSysCmdRuntime) <> 0)
End Function

' Devuelve la ruta de documentos del usuario
Public Function RutaDocumentos() As String
    On Error Resume Next
    RutaDocumentos = Environ$("USERPROFILE") & "\Documents"
End Function

' Devuelve la ruta AppData\Roaming
Public Function RutaAppData() As String
    On Error Resume Next
    RutaAppData = Environ$("APPDATA")
End Function


' ------------------------------------------------------------
' VALIDACIÓN DE REFERENCIAS
' ------------------------------------------------------------

' Devuelve True si hay referencias rotas
Public Function HayReferenciasRotas() As Boolean
    Dim ref As Reference
    For Each ref In Application.References
        If ref.IsBroken Then
            HayReferenciasRotas = True
            Exit Function
        End If
    Next
End Function

' Devuelve una lista de referencias rotas
Public Function ListarReferenciasRotas() As String
    Dim ref As Reference
    Dim salida As String

    For Each ref In Application.References
        If ref.IsBroken Then
            salida = salida & ref.Name & " (" & ref.FullPath & ")" & vbCrLf
        End If
    Next

    ListarReferenciasRotas = salida
End Function


' ------------------------------------------------------------
' VALIDACIÓN DE MÓDULOS Y OBJETOS
' ------------------------------------------------------------

' Comprueba si un módulo existe
Public Function ExisteModulo(nombre As String) As Boolean
    On Error Resume Next
    ExisteModulo = (Not Application.Modules(nombre) Is Nothing)
End Function

' Comprueba si un formulario existe
Public Function ExisteFormulario(nombre As String) As Boolean
    On Error Resume Next
    ExisteFormulario = (Not CurrentProject.AllForms(nombre) Is Nothing)
End Function

' Comprueba si una tabla existe
Public Function ExisteTabla(nombre As String) As Boolean
    On Error Resume Next
    ExisteTabla = (Not CurrentData.AllTables(nombre) Is Nothing)
End Function


' ------------------------------------------------------------
' REGISTRO AUTOMÁTICO DE ERRORES
' ------------------------------------------------------------

' Registra un error en el log del Inspector
Public Sub RegistrarError(errObj As ErrObject, origen As String)
    On Error Resume Next

    Dim mensaje As String
    mensaje = "ERROR en " & origen & vbCrLf & _
              "Descripción: " & errObj.Description & vbCrLf & _
              "Número: " & errObj.Number & vbCrLf & _
              "Fuente: " & errObj.Source

    RegistrarEvento mensaje
End Sub


' ------------------------------------------------------------
' DIAGNÓSTICO COMPLETO DEL ENTORNO
' ------------------------------------------------------------

' Ejecuta un diagnóstico general y lo registra en el log
Public Sub DiagnosticoInspector()
    On Error Resume Next

    RegistrarEvento "=== Diagnóstico del entorno ==="
    RegistrarEvento "Idioma del sistema: " & IdiomaSistema()
    RegistrarEvento "Modo runtime: " & IIf(EsModoRuntime(), "Sí", "No")
    RegistrarEvento "Ruta Documentos: " & RutaDocumentos()
    RegistrarEvento "Ruta AppData: " & RutaAppData()

    If HayReferenciasRotas() Then
        RegistrarEvento "Referencias rotas detectadas:"
        RegistrarEvento ListarReferenciasRotas()
    Else
        RegistrarEvento "No hay referencias rotas."
    End If

    RegistrarEvento "Diagnóstico completado."
End Sub


