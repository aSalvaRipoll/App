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

Private Sub cmdAnalizar_Click()
    Me.subExport.Visible = False
    lblEstado.Caption = "Analizando proyecto..."
    DoEvents

    Dim resultados As Collection
    Set resultados = Inspector.AnalizarProyectoActual

    Set gResultadosInspector = resultados

    Call MostrarResultados(resultados)

    Me.subExport.Visible = True

    lblEstado.Caption = "Análisis completado"
End Sub

Private Sub cmdReparar_Click()
    lblEstado.Caption = "Reparando proyecto..."
    DoEvents

    Inspector.RepararProyectoActual

    lblEstado.Caption = "Reparación completada"
End Sub

Private Sub cmdCerrar_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub lstResultados_DblClick(Cancel As Integer)
    Inspector.NavegarAResultado lstResultados.Value
End Sub

Private Sub MostrarResultados(col As Collection)
    Dim item As Variant

    lstResultados.Clear

    For Each item In col
        lstResultados.AddItem item
    Next item
End Sub

