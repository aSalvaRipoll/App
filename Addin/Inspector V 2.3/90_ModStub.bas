Attribute VB_Name = "90_ModStub"

Option Compare Database
Option Explicit

' ============================================================
' 90_ModStub Módulo de prueba para subExportarInspector
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
End Sub


