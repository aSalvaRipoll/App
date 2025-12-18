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
    Dim formato As FormatoExportacion

    If Me.cboFormato.ListIndex = -1 Then Exit Sub

    formato = Me.cboFormato.Column(0)

    ' Solo HTML requiere estilo
    If formato = ExpTodoHTML Then
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
    Dim formato As FormatoExportacion
    Dim nombreDefecto As String

    If Me.cboFormato.ListIndex = -1 Then
        MsgBox "Selecciona un formato antes de elegir la ruta.", vbExclamation
        Exit Sub
    End If

    formato = Me.cboFormato.Column(0)

    Select Case formato

        ' TXT
        Case ExpResultadosTXT, ExpSimbolosTXT, ExpTodoTXT
            nombreDefecto = "InformeInspector.txt"

        ' Excel
        Case ExpResultadosExcel, ExpSimbolosExcel, ExpTodoExcel
            nombreDefecto = "InformeInspector.xlsx"

        ' HTML
        Case ExpTodoHTML
            nombreDefecto = "InformeInspector.html"

        ' Futuros formatos
        Case Else
            nombreDefecto = "InformeInspector"
    End Select

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
    Dim formato As FormatoExportacion
    Dim ruta As String
    Dim estilo As EstiloHtml

    If Me.cboFormato.ListIndex = -1 Then
        MsgBox "Selecciona un formato de exportación.", vbExclamation
        Exit Sub
    End If

    ruta = Nz(Me.txtRuta.Value, "")
    If ruta = "" Then
        MsgBox "Indica una ruta de archivo.", vbExclamation
        Exit Sub
    End If

    formato = Me.cboFormato.Column(0)

'    Select Case formato
'
'        '=======================================================
'        ' TXT
'        '=======================================================
'        Case ExpResultadosTXT
'            ExportarResultadosAArchivo gResultadosInspector, ruta
'
'        Case ExpSimbolosTXT
'            ExportarSimbolosNoUsadosTXT ruta
'
'        Case ExpTodoTXT
'            ExportarTodoATXT gResultadosInspector, ruta
'
'
'        '=======================================================
'        ' EXCEL
'        '=======================================================
'        Case ExpResultadosExcel
'            ExportarResultadosAExcel gResultadosInspector, ruta
'
'        Case ExpSimbolosExcel
'            ExportarSimbolosNoUsadosExcel ruta
'
'        Case ExpTodoExcel
'            ExportarTodoAExcel gResultadosInspector, ruta
'
'
'        '=======================================================
'        ' HTML
'        '=======================================================
'        Case ExpTodoHTML
'
'            If Me.cboHtmlEstilo.ListIndex = -1 Then
'                MsgBox "Selecciona un estilo HTML.", vbExclamation
'                Exit Sub
'            End If
'
'            estilo = Me.cboHtmlEstilo.Column(0)
'            ExportarTodoAHTML gResultadosInspector, ruta, estilo
'
'
'        '=======================================================
'        ' FUTUROS FORMATOS
'        '=======================================================
'        Case Else
'            MsgBox "Este formato aún no está implementado.", vbInformation
'            Exit Sub
'
'    End Select
'
'    MsgBox "Exportación completada correctamente.", vbInformation
End Sub


'---------------------------------------------------------------
' Inicialización del formulario
'---------------------------------------------------------------
Private Sub Form_Load()

    '===========================================================
    ' Cargar formatos (valor;texto)
    '===========================================================
    With Me.cboFormato
        .ColumnCount = 2
        .ColumnWidths = "0cm;5cm"

        ' TXT
        .AddItem CStr(ExpResultadosTXT) & ";Resultados a TXT"
        .AddItem CStr(ExpSimbolosTXT) & ";Símbolos no usados a TXT"
        .AddItem CStr(ExpTodoTXT) & ";Exportar TODO a TXT"

        ' Excel
        .AddItem CStr(ExpResultadosExcel) & ";Resultados a Excel"
        .AddItem CStr(ExpSimbolosExcel) & ";Símbolos no usados a Excel"
        .AddItem CStr(ExpTodoExcel) & ";Exportar TODO a Excel"

        ' HTML
        .AddItem CStr(ExpTodoHTML) & ";Exportar TODO a HTML"
    End With


    '===========================================================
    ' Cargar estilos HTML (valor;texto)
    '===========================================================
    With Me.cboHtmlEstilo
        .ColumnCount = 2
        .ColumnWidths = "0cm;4cm"
        .AddItem "0;Tema claro"
        .AddItem "1;Tema oscuro"
        .AddItem "2;Tema sepia"
        .AddItem "3;Alto contraste"
        .AddItem "4;Minimalista"
    End With

    ' Deshabilitado hasta que el usuario elija HTML
    Me.cboHtmlEstilo.Enabled = False

End Sub



'---------------------------------------------------------------
' Establecer valores por defecto cuando el formulario ya está listo
'---------------------------------------------------------------
Private Sub Form_Current()
    ' Asegurar que el estilo HTML tiene un valor válido
    If Me.cboHtmlEstilo.ListCount > 0 Then
        If Me.cboHtmlEstilo.ListIndex = -1 Then
            Me.cboHtmlEstilo.ListIndex = 0   ' Tema claro
        End If
    End If
End Sub

