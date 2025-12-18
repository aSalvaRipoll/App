Attribute VB_Name = "modInicioEntornoUsuario"

Option Compare Database
Option Explicit

' ============================================================
'  MÓDULO CENTRAL DE INICIALIZACIÓN DEL INSPECTOR
'  Se ejecuta al abrir los formularios principales del sistema.
'  Prepara cintas, logs, entorno y dependencias.
' ============================================================


' ------------------------------------------------------------
' Inicialización del formulario principal (frmInspector)
' ------------------------------------------------------------
Public Sub InicializarInspectorPrincipal(frm As Form)
    On Error Resume Next

    ' --- Cinta contextual ---
    frm.RibbonName = ""
    frm.RibbonContextualTab = "TabSetForm"

    ' --- Inicializar logs ---
    PrepararCarpetaLogs
    RegistrarEvento "Inspector iniciado"

    ' --- Cargar preferencias del usuario ---
    CargarPreferenciasInspector frm

    ' --- Validar dependencias del entorno ---
    ValidarEntornoInspector

End Sub


' ------------------------------------------------------------
' Inicialización del subformulario de exportación
' ------------------------------------------------------------
Public Sub InicializarInspectorExportacion(frm As Form)
    On Error Resume Next

    ' --- Cinta contextual ---
    frm.RibbonName = ""
    frm.RibbonContextualTab = "TabSetSubform"

    ' --- Cargar preferencias de exportación ---
    CargarPreferenciasExportacion frm

End Sub


' ============================================================
'  SECCIÓN: LOGS
' ============================================================

' Crea la carpeta de logs si no existe
Public Sub PrepararCarpetaLogs()
    Dim carpeta As String
    carpeta = CurrentProject.Path & "\Logs"

    If Dir(carpeta, vbDirectory) = "" Then
        MkDir carpeta
    End If
End Sub

' Registra un evento simple en el log diario
Public Sub RegistrarEvento(texto As String)
    Dim ruta As String
    ruta = CurrentProject.Path & "\Logs\Inspector_" & Format(Date, "yyyy-mm-dd") & ".log"

    Dim f As Integer
    f = FreeFile

    Open ruta For Append As #f
    Print #f, Format(Now, "hh:nn:ss") & " - " & texto
    Close #f
End Sub


' ============================================================
'  SECCIÓN: PREFERENCIAS DEL INSPECTOR
' ============================================================

' Carga preferencias generales del Inspector
Public Sub CargarPreferenciasInspector(frm As Form)
    On Error Resume Next

    ' Ejemplo: restaurar tamaño o estado
    If Nz(GetSetting("InspectorVBA", "Preferencias", "VentanaMaximizada"), "0") = "1" Then
        DoCmd.Maximize
    End If

    ' Ejemplo: restaurar última ruta usada
    frm!txtUltimaRuta = GetSetting("InspectorVBA", "Preferencias", "UltimaRuta", "")
End Sub

' Carga preferencias específicas del panel de exportación
Public Sub CargarPreferenciasExportacion(frm As Form)
    On Error Resume Next

    frm!cboFormato = GetSetting("InspectorVBA", "Exportacion", "Formato", "TXT")
    frm!cboEstilo = GetSetting("InspectorVBA", "Exportacion", "Estilo", "Basico")
    frm!txtRutaDestino = GetSetting("InspectorVBA", "Exportacion", "Ruta", CurrentProject.Path)
End Sub


' ============================================================
'  SECCIÓN: VALIDACIÓN DEL ENTORNO
' ============================================================

' Comprueba que el entorno del Inspector está en condiciones
Public Sub ValidarEntornoInspector()
    On Error Resume Next

    ' Ejemplo: comprobar referencias rotas
    If ReferenciasRotas() Then
        RegistrarEvento "Advertencia: Se detectaron referencias rotas."
    End If

    ' Ejemplo: comprobar existencia de módulos clave
    If Not ExisteModulo("modAuxAnalisis") Then
        RegistrarEvento "Error: Falta el módulo modAuxAnalisis."
    End If
End Sub

' Comprueba si hay referencias rotas
Private Function ReferenciasRotas() As Boolean
    Dim ref As Reference
    For Each ref In Application.References
        If ref.IsBroken Then
            ReferenciasRotas = True
            Exit Function
        End If
    Next
End Function

' Comprueba si un módulo existe
Private Function ExisteModulo(nombre As String) As Boolean
    On Error Resume Next
    ExisteModulo = (Not Application.Modules(nombre) Is Nothing)
End Function

