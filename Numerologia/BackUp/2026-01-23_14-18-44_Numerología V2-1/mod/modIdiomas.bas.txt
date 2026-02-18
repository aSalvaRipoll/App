Attribute VB_Name = "modIdiomas"
' ------------------------------------------------------
' Nombre:    modIdiomas
' Tipo:      Módulo
' Propósito:
' Autor:     asalv
' Fecha:     15/01/2026
' ------------------------------------------------------

' modIdiomas.bas
Option Compare Database
Option Explicit



Public Sub CargarIdiomas()
    Dim rs As DAO.Recordset
    Dim obj As clsIdioma

    Set colIdiomas = New Collection

    Set rs = CurrentDb.OpenRecordset( _
        "SELECT * FROM tbmIdiomas " & _
        "ORDER BY (IDIdioma = 0), NomIdioma")

    Do While Not rs.EOF
        Set obj = New clsIdioma
        obj.Init rs!IDIdioma, rs!Abreviado, rs!NomIdioma, Nz(rs!Notas, "")
        colIdiomas.Add obj, CStr(obj.IDIdioma)
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
End Sub

Public Function CargarIdiomaDesdeID(ByVal id As Byte) As clsIdioma

    If id = 0 Then Exit Function

    Dim rs As DAO.Recordset
    Dim sql As String
    Dim i As New clsIdioma

    sql = "SELECT * FROM tbmIdiomas WHERE IDIdioma = " & id
    Set rs = CurrentDb.OpenRecordset(sql)

    If Not rs.EOF Then
        i.Init rs!IDIdioma, rs!Abreviado, rs!NomIdioma, Nz(rs!Notas, "")
        Set CargarIdiomaDesdeID = i
    End If

    rs.Close
    Set rs = Nothing

End Function

Public Function CargarIdiomaDesdeAbrev(ByVal Abrev As String) As clsIdioma

    If Abrev = "" Then Exit Function

    Dim rs As DAO.Recordset
    Dim sql As String
    Dim i As New clsIdioma

    sql = "SELECT * FROM tbmIdiomas WHERE Abreviado like '" & Abrev & "'"
    Set rs = CurrentDb.OpenRecordset(sql)

    If Not rs.EOF Then
        i.Init rs!IDIdioma, rs!Abreviado, rs!NomIdioma, Nz(rs!Notas, "")
        Set CargarIdiomaDesdeAbrev = i
    End If

    rs.Close
    Set rs = Nothing

End Function


