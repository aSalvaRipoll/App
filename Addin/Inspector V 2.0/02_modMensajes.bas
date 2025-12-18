Attribute VB_Name = "02_modMensajes"

Option Compare Database
Option Explicit

'===============================================================
' Módulo centralizado de mensajes del Inspector
'   - Mensajes para análisis, reparación, exportación, errores
'   - Totalmente extensible
'   - Desacopla UI y lógica
'===============================================================

Private mMensajesReparacion As Scripting.Dictionary
Private mMensajesAnalisis As Scripting.Dictionary
Private mMensajesExportacion As Scripting.Dictionary



'---------------------------------------------------------------
' Inicializar diccionario de mensajes de reparación
'---------------------------------------------------------------
Public Sub MensajesReparacion_Inicializar()

    Set mMensajesReparacion = New Scripting.Dictionary

    mMensajesReparacion.Add ReparacionNoEjecutada, _
        "No hay resultados para reparar."

    mMensajesReparacion.Add ReparacionEjecutada, _
        "Reparación completada correctamente."

    mMensajesReparacion.Add ReparacionConErrores, _
        "Reparación completada con incidencias."
End Sub



'---------------------------------------------------------------
' Obtener mensaje asociado a un estado de reparación
'---------------------------------------------------------------
Public Function MensajeReparacion(estado As EstadoReparacion) As String

    If mMensajesReparacion Is Nothing Then
        MensajesReparacion_Inicializar
    End If

    If mMensajesReparacion.Exists(estado) Then
        MensajeReparacion = mMensajesReparacion(estado)
    Else
        MensajeReparacion = "Estado de reparación desconocido."
    End If
End Function



'---------------------------------------------------------------
' Inicializar diccionario de mensajes de análisis
'---------------------------------------------------------------
Public Sub MensajesAnalisis_Inicializar()

    Set mMensajesAnalisis = New Scripting.Dictionary

    mMensajesAnalisis.Add AnalisisNoEjecutado, _
        "No se ha ejecutado ningún análisis."

    mMensajesAnalisis.Add AnalisisEjecutado, _
        "Análisis completado correctamente."

    mMensajesAnalisis.Add AnalisisConErrores, _
        "Análisis completado con incidencias."
End Sub



'---------------------------------------------------------------
' Obtener mensaje asociado a un estado de análisis
'---------------------------------------------------------------
Public Function MensajeAnalisis(estado As EstadoAnalisis) As String

    If mMensajesAnalisis Is Nothing Then
        MensajesAnalisis_Inicializar
    End If

    If mMensajesAnalisis.Exists(estado) Then
        MensajeAnalisis = mMensajesAnalisis(estado)
    Else
        MensajeAnalisis = "Estado de análisis desconocido."
    End If
End Function



'---------------------------------------------------------------
' Inicializar diccionario de mensajes de exportación
'---------------------------------------------------------------
Public Sub MensajesExportacion_Inicializar()

    Set mMensajesExportacion = New Scripting.Dictionary

    mMensajesExportacion.Add ExportacionNoEjecutada, _
        "No hay resultados para exportar."

    mMensajesExportacion.Add ExportacionEjecutada, _
        "Exportación completada correctamente."

    mMensajesExportacion.Add ExportacionConErrores, _
        "Exportación completada con incidencias."
End Sub



'---------------------------------------------------------------
' Obtener mensaje asociado a un estado de exportación
'---------------------------------------------------------------
Public Function MensajeExportacion(estado As EstadoExportacion) As String

    If mMensajesExportacion Is Nothing Then
        MensajesExportacion_Inicializar
    End If

    If mMensajesExportacion.Exists(estado) Then
        MensajeExportacion = mMensajesExportacion(estado)
    Else
        MensajeExportacion = "Estado de exportación desconocido."
    End If
End Function


