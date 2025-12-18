VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmInicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

'===============================================================
' Form_frmInicio
' Formulario oculto que gestiona el ciclo de vida del Inspector
'===============================================================

Private iniciado As Boolean

Private Sub Form_Open(Cancel As Integer)
    ' Garantiza que el formulario nunca se muestre
    Me.Visible = False
End Sub

Private Sub Form_Activate()
    If iniciado Then Exit Sub

    ' Iniciar entorno del Inspector (menús, VBIDE, clases, etc.)
    Call InicioInspector

    iniciado = True
End Sub

Private Sub Form_Close()
    ' Finalizar entorno del Inspector (limpieza, menús, objetos, etc.)
    Call FinInspector

    iniciado = False
End Sub


