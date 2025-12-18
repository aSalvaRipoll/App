VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_subExportarInspector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Database
Option Explicit


'---------------------------------------------------------------
' Cambio de formato ? habilitar/deshabilitar selector de estilo HTML
'---------------------------------------------------------------
Private Sub cboFormato_AfterUpdate()
    Dim formato As String
    formato = Nz(Me.cboFormato.Value, "")

    ' Solo HTML requiere estilo
    If formato = "Exportar TODO a HTML" Then
        Me.cboHtmlEstilo.Enabled = True

        ' Si no hay selección, aplicar Tema claro
        If Me.cboHtmlEstilo.ListIndex = -1 Then
            Me.cboHtmlEstilo.ListIndex = 0
        End If
    Else
        Me.cboHtmlEstilo.Enabled = False
    End If
End Sub



'---------------------------------------------------------------
' Botón Examinar ? seleccionar ruta de guardado
'---------------------------------------------------------------
Private Sub cmdExaminar_Click()
    Dim dlg As FileDialog
    Dim ruta As String
    Dim formato As String
    Dim nombreDefecto As String

    formato = Nz(Me.cboFormato.Value, "")

    ' Nombre por defecto según formato
    Select Case formato
        Case "Resultados a TXT", "Símbolos no usados a TXT", "Exportar TODO a TXT"
            nombreDefecto = "InformeInspector.txt"

        Case "Resultados a Excel", "Símbolos no usados a Excel", "Exportar TODO a Excel"
            nombreDefecto = "InformeInspector.xlsx"

        Case "Exportar TODO a HTML"
            nombreDefecto = "InformeInspector.html"

        Case Else
            nombreDefecto = "InformeInspector"
    End Select

    ' Si el usuario ya escribió algo, usarlo
    ruta = Nz(Me.txtRuta.Value, "")
    If ruta = "" Then ruta = nombreDefecto

    Set dlg = Application.FileDialog(msoFileDialogSaveAs)
    dlg.InitialFileName = ruta

    If dlg.Show Then
        Me.txtRuta.Value = dlg.SelectedItems(1)
    End If
End Sub



'---------------------------------------------------------------
' Botón Exportar ? ejecutar exportación seleccionada
'---------------------------------------------------------------
Private Sub cmdExportar_Click()
    Dim formato As String
    Dim ruta As String

    formato = Nz(Me.cboFormato.Value, "")
    ruta = Nz(Me.txtRuta.Value, "")

    ' Validaciones básicas
    If formato = "" Then
        MsgBox "Selecciona un formato de exportación.", vbExclamation
        Exit Sub
    End If

    If ruta = "" Then
        MsgBox "Indica una ruta de archivo.", vbExclamation
        Exit Sub
    End If

    ' Exportación según formato
    Select Case formato

        Case "Resultados a TXT"
            ExportarResultadosAArchivo gResultadosInspector, ruta

        Case "Resultados a Excel"
            ExportarResultadosAExcel gResultadosInspector, ruta

        Case "Símbolos no usados a TXT"
            ExportarSimbolosNoUsadosTXT ruta

        Case "Símbolos no usados a Excel"
            ExportarSimbolosNoUsadosExcel ruta

        Case "Exportar TODO a TXT"
            ExportarTodoATXT gResultadosInspector, ruta

        Case "Exportar TODO a Excel"
            ExportarTodoAExcel gResultadosInspector, ruta

        Case "Exportar TODO a HTML"
            ' Validación adicional
            If Me.cboHtmlEstilo.ListIndex = -1 Then
                MsgBox "Selecciona un estilo HTML.", vbExclamation
                Exit Sub
            End If

            Dim Estilo As EstiloHtml
            Estilo = Me.cboHtmlEstilo.Column(0)   ' valor numérico de la enumeración

            ExportarTodoAHTML gResultadosInspector, ruta, Estilo

    End Select

    MsgBox "Exportación completada correctamente.", vbInformation
End Sub



'---------------------------------------------------------------
' Inicialización del formulario
'---------------------------------------------------------------
Private Sub Form_Load()

    ' Formatos disponibles
    With Me.cboFormato
        .AddItem "Resultados a TXT"
        .AddItem "Resultados a Excel"
        .AddItem "Símbolos no usados a TXT"
        .AddItem "Símbolos no usados a Excel"
        .AddItem "Exportar TODO a TXT"
        .AddItem "Exportar TODO a Excel"
        .AddItem "Exportar TODO a HTML"
    End With

    ' Estilos HTML (valor;texto)
    With Me.cboHtmlEstilo
        .ColumnCount = 2
        .ColumnWidths = "0cm;4cm"
        .AddItem "0;Tema claro"
        .AddItem "1;Tema oscuro"
        .AddItem "2;Tema sepia"
        .AddItem "3;Alto contraste"
        .AddItem "4;Minimalista"
    End With

    ' Selección por defecto ? Tema claro
    Me.cboHtmlEstilo.ListIndex = 0

    ' Deshabilitado hasta que el usuario elija HTML
    Me.cboHtmlEstilo.Enabled = False

End Sub

