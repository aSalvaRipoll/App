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

Private Sub Form_Open(Cancel As Integer)
    InicializarInspectorExportacion Me
End Sub

'---------------------------------------------------------------
' Establecer valores por defecto cuando el formulario ya está listo
'---------------------------------------------------------------
Private Sub Form_Current()
    If Me.cboHtmlEstilo.ListCount > 0 Then
        If Me.cboHtmlEstilo.ListIndex = -1 Then
            Me.cboHtmlEstilo
