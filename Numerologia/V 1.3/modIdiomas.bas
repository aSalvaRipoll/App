Attribute VB_Name = "modIdiomas"

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
        obj.Init rs!IDIdioma, rs!Abreviado, rs!NomIdioma, Nz(rs!notas, "")
        colIdiomas.Add obj, CStr(obj.IDIdioma)
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
End Sub


