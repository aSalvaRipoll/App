Attribute VB_Name = "basTest"
Option Compare Database

Public Sub RevisarReferencias()
    Dim ref As Reference
    For Each ref In Application.References
        If ref.IsBroken Then
            Debug.Print "Referencia rota: " & ref.Name & " ? " & ref.FullPath
        End If
    Next ref
    MsgBox "Revisión de referencias completada. Mira la ventana inmediato (Ctrl+G).", vbInformation
End Sub

'Colocar el siguiente código en los eventos de un formulario que contiene un listbox y un botón
Private Sub cmdCargarLista_Click()
     Call ListaComplementosAccess
End Sub

'Private Sub lstLista_DblClick(Cancel As Integer)
'Dim objAddin As clsAddin
'    Set objAddin = colAddins(Me.lstLista)
'        MsgBox "Complemento:" & vbCrLf & objAddin.addin_Name & _
'                vbNewLine & _
'                "Librería:" & vbCrLf & _
'                objAddin.library
'    Set objAddin = Nothing
'End Sub


Sub buscaColeciones()

Dim salida As String
Dim com As Object

For Each com In CurrentProject.AllModules
    If com.Name <> "bastest" Then
        Debug.Print com.Name
    End If
Next


End Sub
