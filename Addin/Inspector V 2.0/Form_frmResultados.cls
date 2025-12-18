VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmResultados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

Private ultimaColumnaOrden As String
Private ultimoAsc As Boolean

'---------------------------------------------------------------
' Formulario: frmResultados
' Interfaz visual del Inspector VBA
'---------------------------------------------------------------

Private Sub Form_Load()

    ' Asegura que el motor está cargado
'    DoCmd.OpenForm "frmInicio", , , , , acHidden

    lstResultados.RowSource = ""
End Sub

Private Sub Form_Open(Cancel As Integer)
'    Me.RibbonName = ""
'    Me.RibbonContextualTab = "TabSetForm"
    
    InicializarInspectorPrincipal Me

End Sub

Private Sub Form_Close()
    GuardarPreferenciasInspector Me
End Sub

'---------------------------------------------------------------

Private Sub cmdAbrirLogs_Click()
    Call Inspector_AbrirCarpetaLogs
End Sub

Private Sub cmdAnalizar_Click()

    Me.subExport.Visible = False
    lblEstado.Caption = "Analizando proyecto..."
    DoEvents

    Dim estado As EstadoAnalisis
    estado = Inspector_Analizar()

    ' Si el análisis no generó resultados, no mostramos nada
    If gResultadosInspector Is Nothing Then
        lblEstado.Caption = MensajeAnalisis(estado)
        Exit Sub
    End If

    ' Mostrar resultados en el formulario
    MostrarResultados gResultadosInspector

    Me.subExport.Visible = True
    lblEstado.Caption = MensajeAnalisis(estado)

End Sub


Private Sub cmdReparar_Click()
    lblEstado.Caption = "Reparando proyecto..."
    DoEvents

    Inspector_Reparar

    lblEstado.Caption = "Reparación completada"
End Sub

Private Sub cmdCerrar_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub lstResultados_DblClick(Cancel As Integer)
    If Not gResultadosInspector Is Nothing Then
        Dim clave As String
        clave = lstResultados.Column(0)
        gResultadosInspector.NavegarAResultado clave
    End If
End Sub

'Private Sub lstResultados_DblClick(Cancel As Integer)
'    If Not gResultadosInspector Is Nothing Then
'        gResultadosInspector.NavegarAResultado lstResultados.Value
'    End If
'End Sub
'Private Sub lstResultados_DblClick(Cancel As Integer)
'    Inspector.NavegarAResultado lstResultados.Value
'End Sub

Private Sub MostrarResultados(res As clsResultados)

    Dim item As clsResultadoAnalisis
    Dim col As Collection
    Dim i As Long
    Dim linea As String
    Dim severidad As String
    Dim elemento As String
    Dim miembro As String

    '-----------------------------------------------------------
    ' Vaciar lista
    '-----------------------------------------------------------
    For i = lstResultados.ListCount - 1 To 0 Step -1
        lstResultados.RemoveItem i
    Next i

    Set col = res.Todos

    '-----------------------------------------------------------
    ' Rellenar lista
    '-----------------------------------------------------------
    For Each item In col

        '-------------------------------------------
        ' Severidad (icono + texto)
        '-------------------------------------------
        severidad = IconoSeveridad(item.severidad)

        '-------------------------------------------
        ' Elemento (icono + nombre)
        '-------------------------------------------
        elemento = FormatoElemento(item.nombreElemento, item.tipoElemento)

        '-------------------------------------------
        ' Miembro (icono + nombre)
        '-------------------------------------------
        miembro = FormatoMiembro(item.nombreMiembro, item.tipoMiembro)

        '-------------------------------------------
        ' Línea (vacía si no aplica)
        '-------------------------------------------
        If item.linea > 0 Then
            linea = CStr(item.linea)
        Else
            linea = ""
        End If

        '-------------------------------------------
        ' Añadir al ListBox (5 columnas)
        '   0 = clave (oculta)
        '   1 = severidad
        '   2 = elemento
        '   3 = miembro
        '   4 = línea
        '-------------------------------------------
        lstResultados.AddItem _
            item.clave & ";" & _
            severidad & ";" & _
            elemento & ";" & _
            miembro & ";" & _
            linea

    Next item

End Sub

'Private Sub MostrarResultados(res As clsResultados)
'    Dim item As clsResultadoAnalisis
'    Dim col As Collection
'    Dim i As Long
'    Dim linea As String
'    Dim severidad As String
'
'    ' Vaciar lista
'    For i = lstResultados.ListCount - 1 To 0 Step -1
'        lstResultados.RemoveItem i
'    Next i
'
'    Set col = res.Todos
'
'    For Each item In col
'
'        Select Case item.severidad
'            Case sevInfo:  severidad = IconoUnicode("Info") & " INFO"
'            Case sevAviso: severidad = IconoUnicode("Aviso") & " AVISO"
'            Case sevError: severidad = IconoUnicode("Error") & " ERROR"
'        End Select
'
'        If item.linea > 0 Then
'            linea = CStr(item.linea)
'        Else
'            linea = ""
'        End If
'
'        lstResultados.AddItem _
'            item.clave & ";" & _
'            severidad & ";" & _
'            item.nombreElemento & ";" & _
'            item.nombreMiembro & ";" & _
'            linea
'
'    Next item
'End Sub

Private Sub lblSeveridad_Click()
    Dim asc As Boolean
    asc = AlternarOrden("Severidad")

    gResultadosInspector.OrdenarPor "Severidad", asc
    ActualizarIndicadores "Severidad", asc
    MostrarResultados gResultadosInspector
End Sub

Private Sub lblElemento_Click()
    Dim asc As Boolean
    asc = AlternarOrden("Elemento")

    gResultadosInspector.OrdenarPor "Elemento", asc
    ActualizarIndicadores "Elemento", asc
    MostrarResultados gResultadosInspector
End Sub

Private Sub lblMiembro_Click()
    Dim asc As Boolean
    asc = AlternarOrden("Miembro")

    gResultadosInspector.OrdenarPor "Miembro", asc
    ActualizarIndicadores "Miembro", asc
    MostrarResultados gResultadosInspector
End Sub

Private Sub lblLinea_Click()
    Dim asc As Boolean
    asc = AlternarOrden("Linea")

    gResultadosInspector.OrdenarPor "Linea", asc
    ActualizarIndicadores "Linea", asc
    MostrarResultados gResultadosInspector
End Sub


Private Function AlternarOrden(campo As String) As Boolean

    If ultimaColumnaOrden = campo Then
        ultimoAsc = Not ultimoAsc
    Else
        ultimoAsc = True   ' primera vez: ascendente
    End If

    ultimaColumnaOrden = campo
    AlternarOrden = ultimoAsc

End Function

Private Sub ActualizarIndicadores(campo As String, asc As Boolean)

    ' Limpiar todos
    lblSeveridad.Caption = "Severidad"
    lblElemento.Caption = "Elemento"
    lblMiembro.Caption = "Miembro"
    lblLinea.Caption = "Línea"

    Dim flecha As String
    flecha = IIf(asc, " " & IconoFlechaArriba(), " " & IconoFlechaAbajo())

    Select Case campo
        Case "Severidad": lblSeveridad.Caption = "Severidad" & flecha
        Case "Elemento":  lblElemento.Caption = "Elemento" & flecha
        Case "Miembro":   lblMiembro.Caption = "Miembro" & flecha
        Case "Linea":     lblLinea.Caption = "Línea" & flecha
    End Select

End Sub

