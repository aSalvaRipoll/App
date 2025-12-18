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
        gResultadosInspector.NavegarAResultado lstResultados.Value
    End If
End Sub
'Private Sub lstResultados_DblClick(Cancel As Integer)
'    Inspector.NavegarAResultado lstResultados.Value
'End Sub

Private Sub MostrarResultados(col As Collection)
    Dim item As Variant

    lstResultados.Clear

    For Each item In col
        lstResultados.AddItem item
    Next item
End Sub

