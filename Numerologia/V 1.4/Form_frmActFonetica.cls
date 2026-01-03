VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmActFonetica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdGuardar_Click()
    Dim tabla As String
    Dim sql As String
    
    If IsNull(cboTipo) Or IsNull(txtOriginal) Or IsNull(txtFonetica) Or IsNull(cboIdioma) Then
        MsgBox "Faltan datos por completar.", vbExclamation
        Exit Sub
    End If
    
    ' Determinar tabla
    If cboTipo = "Nombres" Then
        tabla = "tblNombresExcepciones"
    Else
        tabla = "tblApellidosExcepciones"
    End If
    
    ' Insertar registro
    sql = "INSERT INTO " & tabla & " (Original, Fonetica, Idioma) " & _
          "VALUES ('" & txtOriginal & "', '" & txtFonetica & "', '" & cboIdioma & "')"
    
    CurrentDb.Execute sql, dbFailOnError
    
    MsgBox "Entrada guardada correctamente.", vbInformation
    
    ' Limpiar campos
    txtOriginal = ""
    txtFonetica = ""
    cboIdioma = Null
End Sub


