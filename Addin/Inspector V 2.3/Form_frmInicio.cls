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
    Me.Visible = False
End Sub

Private Sub Form_Activate()
    If iniciado Then Exit Sub

    ' Verificar acceso interno al VBIDE (Add-In)
    If Not VerificaVBIDEinterno() Then
        Debug.Print "Inspector: No se pudo verificar el acceso al VBIDE. El Inspector funcionará en modo limitado."
        ' Aquí podrías deshabilitar funciones del Ribbon si lo deseas
    Else
        Debug.Print "Inspector: Acceso al VBIDE verificado correctamente."
    End If

    ' Cargamos la lista de palabras reservadas
    Call CargarPalabrasReservadas
    ' Iniciar entorno del Inspector (menús, clases, etc.)
    Call InicioInspector

    iniciado = True
End Sub

Private Sub Form_Close()
    Call FinInspector
    iniciado = False
End Sub

