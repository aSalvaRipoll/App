Attribute VB_Name = "modMotor_V_2_2_Aux"
Option Compare Database
Option Explicit

Public Function MarcarTonicas(ByVal silabas As String, ByVal tonicas As String) As String

    Dim partes() As String
    Dim idx() As String
    Dim i As Long, j As Long

    partes = Split(silabas, "-")
    idx = Split(tonicas, ",")

    For j = LBound(idx) To UBound(idx)
        i = CLng(idx(j)) - 1
        If i >= 0 And i <= UBound(partes) Then
            partes(i) = "*" & partes(i) & "*"
        End If
    Next j

    MarcarTonicas = Join(partes, "-")

End Function

Public Function RevisarSilabas_EnFormulario( _
        ByVal texto As String, _
        ByVal silabas As String _
    ) As String

    DoCmd.OpenForm "frmRevisionSilabas", , , , , acHidden ', WindowMode:=acDialog

    With Forms!frmRevisionSilabas
        .TextoOriginal = texto
        .SilabasOriginal = silabas
        .lblOriginal.Caption = texto
        '.txtSilabasOriginal.Value = .ResaltarTonicaHTML(Silabas) '.InsertarEspaciosEnSilabas(texto, Silabas)
        .txtSilabasOriginal.Value = silabas '.NormalizarYResaltarTonica(silabas)
        .txtSilabas = .txtSilabasOriginal
        
        .Visible = True
    End With

    ' Espera a que el formulario se cierre
    'Do While CurrentProject.AllForms("frmRevisionSilabas").IsLoaded
    Do While Forms!frmRevisionSilabas.Visible
        DoEvents
    Loop

    If Forms!frmRevisionSilabas.Cancelado Then
        RevisarSilabas_EnFormulario = silabas
    Else
        RevisarSilabas_EnFormulario = Forms!frmRevisionSilabas.SilabasFinal
    End If
    
    ' Cerrar el formulario aquí, cuando ya hemos leído los datos
    DoCmd.Close acForm, "frmRevisionSilabas"
    
End Function


