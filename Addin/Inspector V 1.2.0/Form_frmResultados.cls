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

 With Me.cboFormato
        .AddItem "Resultados a TXT"
        .AddItem "Resultados a Excel"
        .AddItem "Símbolos no usados a TXT"
        .AddItem "Símbolos no usados a Excel"
        .AddItem "Exportar TODO a TXT"
        .AddItem "Exportar TODO a Excel"
    End With

    Me.fraResultados.Visible = False


    ' Asegura que el motor está cargado
    DoCmd.OpenForm "frmInicio", , , , , acHidden

    lstResultados.RowSource = ""
End Sub

Private Sub cmdAnalizar_Click()
    lblEstado.Caption = "Analizando proyecto..."
    DoEvents

    Dim resultados As Collection
    Set resultados = Inspector.AnalizarProyectoActual

    Call MostrarResultados(resultados)

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

