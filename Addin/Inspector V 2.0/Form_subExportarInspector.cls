VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_subExportarInspector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Option Compare Database
'Option Explicit
'
''---------------------------------------------------------------
'' Inicialización del formulario
''---------------------------------------------------------------
'Private Sub Form_Load()
'
'    '===========================================================
'    ' Cargar formatos (valor;texto)
'    '===========================================================
'    With Me.cboFormato
'        .ColumnCount = 2
'        .ColumnWidths = "0cm;5cm"
'
'        ' TXT
'        .AddItem CStr(ExpResultadosTXT) & ";Resultados a TXT"
'        .AddItem CStr(ExpSimbolosTXT) & ";Símbolos no usados a TXT"
'        .AddItem CStr(ExpTodoTXT) & ";Exportar TODO a TXT"
'
'        ' Excel
'        .AddItem CStr(ExpResultadosExcel) & ";Resultados a Excel"
'        .AddItem CStr(ExpSimbolosExcel) & ";Símbolos no usados a Excel"
'        .AddItem CStr(ExpTodoExcel) & ";Exportar TODO a Excel"
'
'        ' HTML
'        .AddItem CStr(ExpTodoHTML) & ";Exportar TODO a HTML"
'    End With
'
'
'    '===========================================================
'    ' Cargar estilos HTML (valor;texto)
'    '===========================================================
'    With Me.cboHtmlEstilo
'        .ColumnCount = 2
'        .ColumnWidths = "0cm;4cm"
'        .AddItem "0;Tema claro"
'        .AddItem "1;Tema oscuro"
'        .AddItem "2;Tema sepia"
'        .AddItem "3;Alto contraste"
'        .AddItem "4;Minimalista"
'    End With
'
'    ' Deshabilitado hasta que el usuario elija HTML
'    Me.cboHtmlEstilo.Enabled = False
'
'End Sub
'
'Private Sub Form_Open(Cancel As Integer)
''    Me.RibbonName = ""
''    Me.RibbonContextualTab = "TabSetSubform"
'
'    InicializarInspectorExportacion Me
'
'End Sub
'
''---------------------------------------------------------------
'' Establecer valores por defecto cuando el formulario ya está listo
''---------------------------------------------------------------
'Private Sub Form_Current()
'    ' Asegurar que el estilo HTML tiene un valor válido
'    If Me.cboHtmlEstilo.ListCount > 0 Then
'        If Me.cboHtmlEstilo.ListIndex = -1 Then
'            Me.cboHtmlEstilo.ListIndex = 0   ' Tema claro
'        End If
'    End If
'End Sub
'
'Private Sub Form_Close()
'    GuardarPreferenciasExportacion Me
'End Sub
'
''---------------------------------------------------------------
'
'
''---------------------------------------------------------------
'' Cambio de formato ? habilitar/deshabilitar selector de estilo HTML
''---------------------------------------------------------------
'Private Sub cboFormato_AfterUpdate()
'    Dim formato As FormatoExportacion
'
'    If Me.cboFormato.ListIndex = -1 Then Exit Sub
'
'    formato = Me.cboFormato.Column(0)
'
'    ' Solo HTML requiere estilo
'    If formato = ExpTodoHTML Then
'        Me.cboHtmlEstilo.Enabled = True
'
'        ' Si no hay selección, aplicar Tema claro
'        If Me.cboHtmlEstilo.ListIndex = -1 Then
'            Me.cboHtmlEstilo.ListIndex = 0
'        End If
'    Else
'        Me.cboHtmlEstilo.Enabled = False
'    End If
'End Sub
'
'
'
''---------------------------------------------------------------
'' Botón Examinar ? seleccionar ruta de guardado
''---------------------------------------------------------------
'Private Sub cmdExaminar_Click()
'    Dim dlg As FileDialog
'    Dim ruta As String
'
'    If Me.cboFormato.ListIndex = -1 Then
'        MsgBox "Selecciona un formato antes de elegir la ruta.", vbExclamation
'        Exit Sub
'    End If
'
'    ruta = Nz(Me.txtRuta.Value, "")
'    If ruta = "" Then ruta = "InformeInspector"
'
'    Set dlg = Application.FileDialog(msoFileDialogSaveAs)
'    dlg.InitialFileName = ruta
'
'    If dlg.Show Then
'        Me.txtRuta.Value = dlg.SelectedItems(1)
'    End If
'End Sub
'
'
'
''---------------------------------------------------------------
'' Botón Exportar ? ejecutar exportación seleccionada
''---------------------------------------------------------------
'Private Sub cmdExportar_Click()
'    Dim formato As FormatoExportacion
'    Dim ruta As String
'    Dim estilo As EstiloHtml
'
'    If Me.cboFormato.ListIndex = -1 Then
'        MsgBox "Selecciona un formato de exportación.", vbExclamation
'        Exit Sub
'    End If
'
'    ruta = Nz(Me.txtRuta.Value, "")
'    If ruta = "" Then
'        MsgBox "Indica una ruta de archivo.", vbExclamation
'        Exit Sub
'    End If
'
'    formato = Me.cboFormato.Column(0)
'
'    ' Si es HTML, validar estilo
'    If formato = ExpTodoHTML Then
'        If Me.cboHtmlEstilo.ListIndex = -1 Then
'            MsgBox "Selecciona un estilo HTML.", vbExclamation
'            Exit Sub
'        End If
'        estilo = Me.cboHtmlEstilo.Column(0)
'    Else
'        estilo = TemaClaro   ' Valor por defecto
'    End If
'
'    ' Delegar en el módulo principal
'    Call Inspector_Exportar(formato, ruta, estilo)
'
'    Call GuardarPreferenciasExportacion(Me)
'
'End Sub


Option Compare Database
Option Explicit

' ============================================================
' Módulo de prueba para subExportarInspector
' Permite validar la cinta contextual sin lógica completa
' ============================================================

' --- Variables internas ---
Private mFormatoActual As String
Private mEstiloActual As String

' ============================================================
' Propiedades requeridas por la cinta
' ============================================================

Public Property Get FormatoActual() As String
    If Len(mFormatoActual) = 0 Then
        mFormatoActual = "TXT"   ' Valor por defecto
    End If
    FormatoActual = mFormatoActual
End Property

Public Property Let FormatoActual(ByVal valor As String)
    mFormatoActual = valor
End Property

Public Property Get EstiloActual() As String
    If Len(mEstiloActual) = 0 Then
        mEstiloActual = "Claro"   ' Valor por defecto
    End If
    EstiloActual = mEstiloActual
End Property

Public Property Let EstiloActual(ByVal valor As String)
    mEstiloActual = valor
End Property

' ============================================================
' Métodos llamados desde la cinta
' ============================================================

Public Sub CambiarFormato(ByVal nuevoFormato As String)
    mFormatoActual = nuevoFormato
    Debug.Print "Formato cambiado a: "; nuevoFormato
End Sub

Public Sub CambiarEstilo(ByVal nuevoEstilo As String)
    mEstiloActual = nuevoEstilo
    Debug.Print "Estilo cambiado a: "; nuevoEstilo
End Sub

Public Sub cmdExaminar_Click()
    MsgBox "Simulación: Examinar ruta…", vbInformation
End

