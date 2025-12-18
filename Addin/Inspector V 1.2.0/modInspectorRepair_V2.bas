Attribute VB_Name = "modInspectorRepair_V2"

Option Compare Database
Option Explicit

'===============================================================
' Módulo: modInspectorRepair
' Motor moderno de reparación del Inspector VBA
'===============================================================

'---------------------------------------------------------------
' Ejecuta reparaciones sobre una colección de resultados
'---------------------------------------------------------------
Public Sub RepararResultados(resultados As Collection)
    Dim res As clsResultadoAnalisis
    Dim n As Long

    If Not AsegurarReferenciaVBIDE() Then
        Debug.Print "No se puede ejecutar la reparación sin VBIDE."
        Exit Sub
    End If

    If resultados Is Nothing Then
        Debug.Print "No hay resultados para reparar."
        Exit Sub
    End If

    Debug.Print
    Debug.Print "==============================================="
    Debug.Print "   INICIO DE REPARACIÓN"
    Debug.Print "==============================================="

    For Each res In resultados
        If res.esReparable Then
            RepararResultado res
            n = n + 1
        End If
    Next res

    Debug.Print "-----------------------------------------------"
    Debug.Print "Reparaciones realizadas: "; n
    Debug.Print "==============================================="
    Debug.Print "   FIN DE REPARACIÓN"
    Debug.Print "==============================================="
End Sub

'---------------------------------------------------------------
' Ejecuta la reparación asociada a un resultado
'---------------------------------------------------------------
Private Sub RepararResultado(res As clsResultadoAnalisis)
    Debug.Print "Reparando: "; res.Formatear

    Select Case True

        Case res.codigoReparacion = "ADD_OPTION_EXPLICIT"
            RepararOptionExplicit res.nombreElemento

        Case Left$(res.codigoReparacion, 14) = "FIX_REFERENCE:"
            RepararReferenciaEspecifica Mid$(res.codigoReparacion, 15)

        Case Else
            Debug.Print "  >> No hay rutina de reparación definida para: "; res.codigoReparacion
    End Select
End Sub

'---------------------------------------------------------------
' Reparación: añadir Option Explicit
'---------------------------------------------------------------
Private Sub RepararOptionExplicit(nombreModulo As String)
    Dim vbProj As VBIDE.VBProject
    Dim comp As VBIDE.VBComponent
    Dim cm As VBIDE.CodeModule
    Dim insertLinea As Long
    Dim i As Long, linea As String

    If Len(nombreModulo) = 0 Then
        Debug.Print "  >> No se puede reparar: nombre de módulo vacío."
        Exit Sub
    End If

    Set vbProj = Application.VBE.ActiveVBProject

    For Each comp In vbProj.VBComponents
        If comp.Name = nombreModulo Then
            Set cm = comp.CodeModule

            ' Evitar duplicados
            If InStr(1, cm.Lines(1, cm.CountOfLines), "Option Explicit", vbTextCompare) > 0 Then
                Debug.Print "  >> Option Explicit ya existe en "; nombreModulo
                Exit Sub
            End If

            ' Buscar primera línea no vacía ni comentario
            insertLinea = 1
            For i = 1 To cm.CountOfLines
                linea = Trim$(cm.Lines(i, 1))
                If Len(linea) = 0 Or Left$(linea, 1) = "'" Then
                    insertLinea = i + 1
                Else
                    Exit For
                End If
            Next i

            On Error GoTo ErrHandler
            cm.InsertLines insertLinea, "Option Explicit"
            Debug.Print "  >> Option Explicit añadido en "; nombreModulo
            Exit Sub

ErrHandler:
            Debug.Print "  >> Error al insertar Option Explicit en "; nombreModulo; ": "; Err.Description
            Exit Sub
        End If
    Next comp

    Debug.Print "  >> No se encontró el módulo: "; nombreModulo
End Sub

'---------------------------------------------------------------
' Reparación: referencias específicas
'---------------------------------------------------------------
Private Sub RepararReferenciaEspecifica(nombreRef As String)
    If Len(nombreRef) = 0 Then
        Debug.Print "  >> Código de referencia vacío."
        Exit Sub
    End If

    Select Case UCase$(nombreRef)

        Case "VBIDE"
            If AsegurarReferenciaVBIDE() Then
                Debug.Print "  >> Referencia VBIDE reparada."
            Else
                Debug.Print "  >> No se pudo reparar VBIDE automáticamente."
            End If

        Case Else
            Debug.Print "  >> Reparación no implementada para referencia: "; nombreRef

    End Select
End Sub

