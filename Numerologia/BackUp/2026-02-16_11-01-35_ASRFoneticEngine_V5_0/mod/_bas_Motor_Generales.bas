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

