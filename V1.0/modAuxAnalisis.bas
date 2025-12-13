Attribute VB_Name = "modAuxAnalisis"

Option Compare Database
Option Explicit


Public Function DetectarMiembros(code As VBIDE.CodeModule) As Collection
    Dim col As New Collection
    Dim i As Long
    Dim linea As String
    Dim m As clsMiembro
    Dim kind As VBIDE.vbext_ProcKind

    For i = 1 To code.CountOfLines
        linea = Trim(code.Lines(i, 1))

        If EsInicioMiembro(linea) Then
            Set m = New clsMiembro
            m.nombre = ExtraerNombreMiembro(linea)
            m.Tipo = ExtraerTipoMiembro(linea)
            m.Ambito = ExtraerAmbitoMiembro(linea)

            kind = code.ProcKind(m.nombre)
            m.lineaInicio = code.ProcStartLine(m.nombre, kind)
            m.NumLineas = code.ProcCountLines(m.nombre, kind)
            m.LineaFin = m.lineaInicio + m.NumLineas - 1

            CalcularMetricasMiembro code, m

            m.EsPropiedad = (Left$(m.Tipo, 8) = "Property")
            m.EsFuncion = (m.Tipo = "Function")
            m.EsProcedimiento = (m.Tipo = "Sub")
            m.EsEvento = EsMiembroEvento(m.nombre, code, m.lineaInicio)

            col.Add m
        End If
    Next i

    Set DetectarMiembros = col
End Function

Private Sub CalcularMetricasMiembro(code As VBIDE.CodeModule, m As clsMiembro)
    Dim i As Long
    Dim linea As String

    For i = m.lineaInicio To m.LineaFin
        linea = code.Lines(i, 1)
        Select Case ClasificarLinea(linea)
            Case "Codigo":        m.NumLineasCodigo = m.NumLineasCodigo + 1
            Case "Comentario":    m.NumLineasComentario = m.NumLineasComentario + 1
            Case "Vacia":         m.NumLineasVacias = m.NumLineasVacias + 1
        End Select
    Next i
End Sub

Private Function ClasificarLinea(linea As String) As String
    Dim t As String
    t = Trim(linea)

    If t = "" Then
        ClasificarLinea = "Vacia"
    ElseIf Left$(t, 1) = "'" Or Left$(t, 3) = "Rem" Then
        ClasificarLinea = "Comentario"
    Else
        ClasificarLinea = "Codigo"
    End If
End Function

Private Function EsMiembroEvento(nombre As String, code As VBIDE.CodeModule, lineaInicio As Long) As Boolean
    ' Heurística básica: nombre contiene "_"
    ' (Form_Load, cmdAceptar_Click, etc.)
    If InStr(1, nombre, "_") > 0 Then
        EsMiembroEvento = True
    Else
        EsMiembroEvento = False
    End If
End Function

