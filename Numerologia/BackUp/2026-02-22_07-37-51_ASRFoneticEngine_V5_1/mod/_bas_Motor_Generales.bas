Attribute VB_Name = "_bas_Motor_Generales"
Option Compare Database
Option Explicit

Public FonemasBase As Object
Public FonemasValor As Object


Public Sub CargarFonemasBase()
    Dim rs As DAO.Recordset
    Set FonemasBase = CreateObject("Scripting.Dictionary")

    Set rs = CurrentDb.OpenRecordset("SELECT Grafema, ID FROM qryFonemasBase")

    Do While Not rs.EOF
        FonemasBase(LCase$(rs!grafema)) = rs!id
        rs.MoveNext
    Loop

    rs.Close
End Sub


Public Sub CargarFonemasValor()
    Dim rs As DAO.Recordset
    Set FonemasValor = CreateObject("Scripting.Dictionary")

    Set rs = CurrentDb.OpenRecordset("SELECT ID, Grafema, Valor FROM qryFonemasValor")

    Do While Not rs.EOF
        FonemasValor(rs!id & "|" & LCase$(rs!grafema)) = rs!valor
        rs.MoveNext
    Loop

    rs.Close
End Sub

'-------------------------------------------------------------
'-------------------------------------------------------------

Private Sub CargaCombos(frm As Form)

    With frm

        ' Limpiar combos
        .cmbVocales.Value = ""
        .cmbTrataH.Value = ""
        .cmbGenero.Value = ""
        .cmbNumero.Value = ""
        ' ... los que correspondan

        ' Cargar valores según idioma
        ' Ejemplo para Castellano
        .cmbVocales.Clear
        .cmbVocales.AddItem "0;Ninguna"
        .cmbVocales.AddItem "1;Acentos Castellanos"
        .cmbVocales.AddItem "2;Acentos Graves y Agudos"
        .cmbVocales.AddItem "3;Todos los acentos y marcadores"

        ' ... resto de combos según idioma

        ' Cargar el ListBox con SQL dinámico
        Dim strSQL As String
        strSQL = CreaSQL("Castellano")
        .lstDigrafos.RowSource = strSQL

    End With

End Sub

Sub limpiarPrefijos()

    Dim rsI As DAO.Recordset
    Dim rsO As DAO.Recordset
    Dim f As DAO.Field
    Dim d As String, v As String
    
    CurrentDb.Execute "Delete from tbmPrefijos_Temp"
    
    'Set rsI = CurrentDb.OpenRecordset("tbtPrefijos")
    'Set rsI = CurrentDb.OpenRecordset("tbtPrefijoides")
    Set rsI = CurrentDb.OpenRecordset("tbtFormantes")
    
    Set rsO = CurrentDb.OpenRecordset("tbmPrefijos_Temp")
    
    While Not rsI.EOF
        DoEvents
        rsO.AddNew
        For Each f In rsI.Fields
            If f.Name = "prefijo" Then
            
                d = Trim(rsI(f.Name))
                v = vbNullString
                
                Debug.Print d
                
                If Right(d, 1) = "-" Then
                    d = Left(d, Len(d) - 1)
                End If
                
                If InStr(d, "(") Then
                    v = Trim(Mid(d, InStr(d, "(")))
                    d = Trim(Replace(d, v, ""))
                    
                    v = Replace(v, "(", "")
                    v = Replace(v, ")", "")
                    
                    'rsO(f.Name) = d
                    rsO("vocal") = Trim(v)
                End If
                rsO(f.Name) = d
                
            Else
                rsO(f.Name) = rsI(f.Name)
            End If
        Next
        rsO.Update
        rsI.MoveNext
    Wend
    
End Sub

