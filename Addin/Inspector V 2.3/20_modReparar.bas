Attribute VB_Name = "20_modReparar"

Option Compare Database
Option Explicit

'===============================================================
' Módulo: 20_modReparar
' Motor moderno de reparación del Inspector VBA
'===============================================================

'---------------------------------------------------------------
' Punto de entrada global: repara todo el proyecto
'---------------------------------------------------------------
Public Sub RepararProyecto(cat As clsCatalogoInspector)
    If cat Is Nothing Then
        Debug.Print "No hay catálogo para reparar."
        Exit Sub
    End If

    RepararResultados cat.resultados
End Sub

'---------------------------------------------------------------
' Ejecuta reparaciones sobre una colección de resultados
'---------------------------------------------------------------
Public Sub RepararResultados(resultados As Collection)
    Dim res As clsResultadoAnalisis
    Dim n As Long

    ' El motor de reparación necesita acceso al VBIDE
    If Not gVBIDEDisponible Then
        Debug.Print "No se puede ejecutar la reparación: el VBIDE no está disponible."
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

        ' Espacio para futuras reparaciones
        ' Case res.codigoReparacion = "REMOVE_UNUSED_IMPORTS"
        '     RepararImports res.nombreElemento

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
' Reparación: referencias específicas del PROYECTO ANALIZADO
'---------------------------------------------------------------
Private Sub RepararReferenciaEspecifica(nombreRef As String)
    If Len(nombreRef) = 0 Then
        Debug.Print "  >> Código de referencia vacío."
        Exit Sub
    End If

    Select Case UCase$(nombreRef)

        Case "VBIDE"
            ' Si el proyecto analizado tiene referencia VBIDE y está rota, se repara.
            RepararReferenciaDelProyecto "VBIDE"

        Case Else
            RepararReferenciaDelProyecto nombreRef

    End Select
End Sub

'---------------------------------------------------------------
' Reparación genérica de referencias del proyecto analizado
'---------------------------------------------------------------
Private Sub RepararReferenciaDelProyecto(nombreRef As String)
    Dim ref As Reference
    Dim vbProj As VBIDE.VBProject

    Set vbProj = Application.VBE.ActiveVBProject

    On Error Resume Next

    For Each ref In vbProj.References
        If UCase$(ref.Name) = UCase$(nombreRef) Then

            If ref.IsBroken Then
                Debug.Print "  >> Reparando referencia rota: "; nombreRef

                ' Intento de reparación automática
                ref.FullPath = ref.FullPath

                If ref.IsBroken Then
                    Debug.Print "  >> No se pudo reparar la referencia: "; nombreRef
                Else
                    Debug.Print "  >> Referencia reparada correctamente: "; nombreRef
                End If

            Else
                Debug.Print "  >> La referencia no está rota: "; nombreRef
            End If

            Exit Sub
        End If
    Next ref

    Debug.Print "  >> No se encontró la referencia en el proyecto: "; nombreRef
End Sub

